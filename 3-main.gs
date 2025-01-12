/***********************************************
 * Steps:
 * 1. Fetch unread emails and filter them for processing based on the search query.
 * 2. Skip already processed emails unless overridden in the configuration (RERUN_READ_MAILS).
 * 3. Extract email data and determine the source (email or SMS).
 * 4. Identify the applicable rules for parsing transaction details or handle duplicate emails.
 * 5. Sanitize the email body/SMS content and extract key transaction details:
 *    - Date, accounts (from/to), amount, type (debit/credit), merchant, and metadata.
 * 6. Validate mandatory fields (date, amount, category, account) and classify transactions:
 *    - Debit, credit, or transfer.
 * 7. Handle "Transfer" transactions by creating corresponding debit and credit entries.
 * 8. Prepare transaction data for appending to the Google Sheet.
 * 9. Append successfully processed transactions to the spreadsheet.
 * 10. Label processed emails and update their status as read.
 * 11. Generate a summary URL and send an email with the transaction details.
 ***********************************************/

function processTransactionEmails(e) {
  try {
    // Step 1: Fetch unread emails and filter them for processing
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    transactionSheet = spreadsheet.getSheetByName(DEV_CONFIG.OUTPUT_SHEET_TITLE);

    const emailThreads = getEmailThreads(e);
    Logger.log('[Info] Number of email threads found: ' + emailThreads.length);

    const transactions = []; // Holds all processed transactions
    const metadata = []; // Holds source and silentErrors

    // Get Label;
    const labelRequest = getLabelRequest(LABELS.PROCESSED);

    // Each thread can have multiple emails. Each email will have one transaction related info.
    emailThreads.forEach(function (thread) {

      // Emails in each thread
      var emails = thread.getMessages();
      emails.forEach(function (email) {
        
        // Step 2: Skip already processed emails unless overridden in config
        if (email.isUnread() || DEV_CONFIG.RERUN_READ_MAILS) {
          ProcessedCount.TOTAL++;

          // Step 3: Extract email data and determine the source (email or SMS).
          populateEmailData(email);
          const { emailId, emailSubject } = emailData;  // These are used multiple times below.
          Logger.log(`[START] Processing email ID: ${emailId}, Subject: ${emailSubject}`);

          if (isEmailAlreadyProcessed()) {
            Logger.log(`[SKIP] ${emailId} is Already processed. ${emailSubject}`);
            ProcessedCount.SKIPPED++;
            return;
          }

          // Logger.log(`[INFO] messageBody : ${emailData.messageBody}`);
          currentTransactionSilentErrors = [];

          // Step 4: Identify the applicable rules for parsing transaction details or handle duplicate emails.
          const applicableRegexMap = getApplicableRegexMap(emailData.isSMS);
          if (!applicableRegexMap) {
            logError(ErrorType.NO_RULE, `(${emailData.source}) No ruleId found for ${emailSubject}`, true); // Stopping error
            createFailureRecord("No applicable rule found.");
            return;
          }
          if (applicableRegexMap == CONSTANTS.DUPLICATE) return;

          const regexMap = combineWithDefaultRegexMap(applicableRegexMap);

          // Step 5: Sanitize the email body/SMS content and extract key transaction details:
          //          - Date, accounts (from/to), amount, type (debit/credit), merchant, and metadata.
          var sanitizedText = cleanEmailBody(emailData.messageBody);
          const transactionDate = extractTransactionDate(sanitizedText, regexMap.dateRegex, emailData.emailDate);
          const fromAccount = getAccountName(extractFullMatchOrFirstCaptureGroupString(sanitizedText, regexMap.fromAccountRegex), sanitizedText);
          const toAccount = getAccountName(extractFullMatchOrFirstCaptureGroupString(sanitizedText, regexMap.toAccountRegex), sanitizedText);
          const transactionAmount = getTransactionAmount(sanitizedText, regexMap.amountRegex); // Remove everything except digits, decimal & - sign
          const isDebit = isDebitTransaction(sanitizedText, regexMap.isDebit, toAccount, transactionAmount); // Update to isIncome ?

          const transactionType = getTransactionType(toAccount, isDebit);
          const merchant = extractFullMatchOrFirstCaptureGroupString(sanitizedText, regexMap.merchantRegex);
          const notes = extractNotes(sanitizedText, regexMap.notesRegex);
          var {categoryType, category, subcategory} = getTransactionCategory(extractFullMatchOrCaptureGroupsArray(sanitizedText, regexMap.categoryRegex), sanitizedText, isDebit);
          subcategory = getTransactionSubcategory(sanitizedText, regexMap.subcategoryRegex, subcategory);

          const adjustedAmount = isDebit ? -transactionAmount : transactionAmount;
          const title = generateTitle(merchant, sanitizedText);

          // Step 6: Validate mandatory fields and classify transactions: Debit, credit, or transfer.
          if (!validateMandatoryFields(transactionDate, transactionAmount, category, fromAccount)) return;

          // Step 7: Handle "Transfer" transactions by creating corresponding debit and credit entries from a single email
          // if (transactionType === "Transfer") {
          //   const debitTransaction = {...transactionPayload, amount: -transactionAmount, transactionType: "Debit"};
          //   const creditTransaction = {...transactionPayload, amount: transactionAmount, transactionType: "Credit"};
          //   transactions.push(debitTransaction, creditTransaction);
          // } else {
          // Normal transaction - Separate email for each transaction
          var transactionPayload = filterPayload({
            amount: adjustedAmount,
            title,
            notes,
            date: moment(transactionDate).format(DATE_FORMATS.CASHEW_FORMAT),
            category,
            subcategory,
            account: fromAccount
          });
          var silentErrors = currentTransactionSilentErrors.join("; ");
          transactions.push({...transactionPayload});
          // Aggregating these separately since these're not contributing to transaction URL.
          metadata.push({source : emailData.source, silentErrors});
          // }

          // Step 8: Prepare transaction data for appending to the spreadsheet
          var rowData = [
            CONSTANTS.SUCCESS_STATUS,
            moment(transactionDate).format(DATE_FORMATS.DISPLAY.DATETIME),
            `${adjustedAmount}`,
            transactionType,
            fromAccount,
            category || "",
            subcategory || "",
            merchant || "",
            title || "",
            notes || "",
            silentErrors,
            emailId,
            getEmailThreadLink(),
            createSingleTransactionUrl(transactionPayload),
            emailSubject,
            getToday()
          ];

          // Step 9: Append successfully processed transactions to the Google Sheet
          Logger.log('[SUCCESS] Appending to Google Sheet: ' + rowData);
          transactionSheet.appendRow(rowData);
          ProcessedCount.SUCCESS++;

          // Step 10: Label processed emails and update their status
          if (DEV_CONFIG.MARK_AS_PROCESSED) {
            // Approach 1 : Apply label to thread. (Doesn't use GMAIL API)
            // Even if some emails in the thread are not processed, the tag is still applied on all emails and they're not picked on rerun of script.
            // let nestedLabelPath = "Txs/✅";
            // let label = GmailApp.getUserLabelByName(nestedLabelPath) //|| GmailApp.createLabel(nestedLabelPath);
            // thread.addLabel(label);

            // Approach 2 [Recommended] : Apply label to email. (Uses GMAIL API) - Advantage in Readme Appendix
            Gmail.Users.Messages.modify(labelRequest, 'me', emailId); // Apply the label
            Logger.log(`Label '${LABELS.PROCESSED}' applied to email with ID: ${emailId}`);    
            email.markRead();
          }
        }
      });
    });

    // Step 11: Generate a summary URL and send an email with the transaction details.
    printProcessingSummary(); 
    if (transactions.length > 0) {
      var finalTransactionUrl = createFinalTransactionUrl({transactions});
      sendSuccessEmail(finalTransactionUrl, transactions, metadata)
      return {transactions, ProcessedCount};  // Response for /dev & /exec
    }
  } catch (e) {
    Logger.log("[ERROR] Unhandled error in mainFunction: " + e.message);
    Logger.log(e.stack); // Logs full stack trace
    throw e; // Re-throw the error after logging
  }
}