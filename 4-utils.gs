// ██████   █████  ██████  ███████ ███████     ███████ ███    ███  █████  ██ ██      
// ██   ██ ██   ██ ██   ██ ██      ██          ██      ████  ████ ██   ██ ██ ██      
// ██████  ███████ ██████  ███████ █████       █████   ██ ████ ██ ███████ ██ ██      
// ██      ██   ██ ██   ██      ██ ██          ██      ██  ██  ██ ██   ██ ██ ██      
// ██      ██   ██ ██   ██ ███████ ███████     ███████ ██      ██ ██   ██ ██ ███████ 

function getLabelRequest(labelName){
  const allLabels = Gmail.Users.Labels.list('me');
  let label = allLabels.labels.find(l => l.name === labelName) || Gmail.Users.Labels.create({ name: labelName }, 'me');
  return { addLabelIds: [label.id], labelName };
}

function populateEmailData(email){
  const isSMS = new RegExp(`^(${CONSTANTS.TRANSACTION_SMS}|${CONSTANTS.BACKUP_SMS})`).test(email.getSubject());
  emailData = {
    emailId : email.getId(),
    emailSubject : email.getSubject(),
    messageBody : email.getPlainBody(),
    emailDate : email.getDate(), // Returns a Date object
    isSMS,
    source : isSMS ? Source.SMS : Source.EMAIL
  }
}

function isEmailAlreadyProcessed() {
  if(DEV_CONFIG.RERUN_PROCESSED_IN_SHEET) return false;
  var emailIdRange = transactionSheet.getRange(2, 12, transactionSheet.getLastRow(), 1).getValues(); // Column for Email ID
  return emailIdRange.some(row => row[0] === emailData.emailId);
}

function labelEmail(labelRequest, email) {
  if (DEV_CONFIG.MARK_AS_PROCESSED) {
    // Approach 1 : Apply label to thread. (Doesn't use GMAIL API)
    // Even if some emails in the thread are not processed, the tag is still applied on all emails and they're not picked on rerun of script.
    // let nestedLabelPath = "Txs/✅";
    // let label = GmailApp.getUserLabelByName(nestedLabelPath) //|| GmailApp.createLabel(nestedLabelPath);
    // thread.addLabel(label);

    // Approach 2 [Recommended] : Apply label to email. (Uses GMAIL API) - https://developers.google.com/gmail/api/quickstart/apps-script#configure_the_script
    // The advantage of labeling email instead of thread is that in the edge case when few messages in thread are passing and rest are not, the label and read status
    // gets updated only on the email which actually got processed. (Gmail > Settings > Uncheck "Conversation view"  to see each eamil in thread individually).
    // When the script reruns, since some emails don't contain the LABEL.PROCESSED, it picks threads for those emails. Among all emails of threads, ignores read ones and processed the unread ones.
    Gmail.Users.Messages.modify(labelRequest, 'me', emailData.emailId); // Apply the label
    // TODO: Add failed subjects to Success email kanpilotID(ut8hsuyx5du0tljkcxnc7ixw)
    Logger.log(`Label '${labelRequest.labelName}' applied to email with ID: ${emailData.emailId}`);
    email.markRead();
  }
}

function getMandatoryFields(transactionDate, transactionAmount, category, fromAccount) {
  return { 
    Date: transactionDate, 
    Amount: transactionAmount, 
    Category: category, 
    Account: fromAccount 
  };
}

function validateMandatoryFields(mandatoryFields, email, labelRequestFailed) {
  const missingFields = Object.keys(mandatoryFields).filter(field => !mandatoryFields[field]);
  
  if (missingFields.length > 0) {
    const errorMessage = `Missing mandatory fields - ${missingFields.join(", ")}`;
    logError(ErrorType.MISSING_FIELDS, `${emailData.emailSubject} : ${errorMessage}. Skipping this message.`, true); // Stopping error
    createFailureRecord(errorMessage);
    labelEmail(labelRequestFailed, email);
    return false; // Indicates validation failure
  }
  return true; // Indicates validation success
}

// Utility function to clean email body text
function cleanEmailBody(text) {
  return text.replace(/\r?\n|\r/g, " ").replace(/\*/g, "").replace(/\s+/g, " ").trim();
}

// ███████ ██ ███    ██ ██████      ██████  ██    ██ ██████  ██      ██  ██████  █████  ████████ ███████ ███████     
// ██      ██ ████   ██ ██   ██     ██   ██ ██    ██ ██   ██ ██      ██ ██      ██   ██    ██    ██      ██          
// █████   ██ ██ ██  ██ ██   ██     ██   ██ ██    ██ ██████  ██      ██ ██      ███████    ██    █████   ███████     
// ██      ██ ██  ██ ██ ██   ██     ██   ██ ██    ██ ██      ██      ██ ██      ██   ██    ██    ██           ██     
// ██      ██ ██   ████ ██████      ██████   ██████  ██      ███████ ██  ██████ ██   ██    ██    ███████ ███████    

// Utility to normalize field values
const normalizeField = value => (value ? value.toString() : '');

// Utility to format date for logging
const formatDate = date =>
  moment(date).isValid() ? moment(date).format('DD-MM-YY HH:mm:ss') : '';

// Utility to extract the month as a short name (e.g., "Jan", "Feb")
const extractMonth = date => moment(date).format('MMM');

// Utility to chunk large log output to avoid truncation
const logInChunks = (header, rows, chunkSize = 20) => {
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize).join('\n');
    Logger.log(`${header}\n${chunk}`);
    header = ''; // Ensure header is only logged once
  }
};

// Function to calculate max column widths for alignment
const calculateColumnWidths = (rows, headers) => {
  const allRows = [headers, ...rows];
  return allRows[0].map((_, colIndex) =>
    Math.max(...allRows.map(row => (row[colIndex] || '').toString().length))
  );
};

// Function to pad columns for alignment
const padRowColumns = (row, columnWidths) =>
  row.map((cell, index) => (cell || '').toString().padEnd(columnWidths[index], ' ')).join('\t');

// Logs existing transactions in a tabular format (filtered by month)
const logExistingRowsTable = (normalizedRows, tableHeaders, transactionMonth) => {
  const monthFilteredRows = normalizedRows.filter(row => extractMonth(row[0]) === transactionMonth);

  if (monthFilteredRows.length === 0) {
    Logger.log(`[DEBUG] No transactions found for ${transactionMonth}`);
    return;
  }

  // Format rows for logging
  const formattedRows = monthFilteredRows.map((row, index) => [
    (index + 1).toString(), formatDate(row[0]), ...row.slice(1),
  ]);

  // Compute column widths
  const columnWidths = calculateColumnWidths(formattedRows, tableHeaders);
  const headerRow = padRowColumns(tableHeaders, columnWidths);
  const dataRows = formattedRows.map(row => padRowColumns(row, columnWidths));

  // Log header separately before chunking rows
  Logger.log(`\n[DEBUG] Existing Transactions for ${transactionMonth} (Total: ${monthFilteredRows.length})\n${headerRow}`);
  logInChunks('', dataRows);  // Header already logged, pass empty string
};

// Compares a row with the transaction payload
const compareRowWithPayload = (row, transactionPayload, rowIndex, source) => {
  const rowDate = moment(new Date(row[0])); // Spreadsheet date
  const payloadDate = moment(transactionPayload.date, DATE_FORMATS.CASHEW_FORMAT); // Payload date

  if (!rowDate.isValid() || !payloadDate.isValid()) {
    Logger.log(`[ERROR] Invalid date comparison at Row Index: ${rowIndex + 2}`);
    return { rowIndex: rowIndex + 2, status: '⚠️', rowData: row };
  }

  // If source is SMS, compare only date (ignore time)
  const isDateMatch = (source === Source.SMS) ? rowDate.isSame(payloadDate, 'day') : rowDate.unix() === payloadDate.unix();

  const isMatch =
    isDateMatch &&
    parseFloat(row[1]) === parseFloat(transactionPayload.amount) &&
    row[3] === normalizeField(transactionPayload.account) &&
    row[4] === normalizeField(transactionPayload.category) &&
    row[5] === normalizeField(transactionPayload.subcategory) &&
    row[7] === normalizeField(transactionPayload.title) &&
    row[8] === normalizeField(transactionPayload.notes);

  return { rowIndex: rowIndex + 2, status: isMatch ? '✅' : '❌', rowData: row };
};

// Logs the comparison table
const buildComparisonTable = (comparisonResults, tableHeaders, payloadIndex, transactionMonth) => {
  const monthFilteredResults = comparisonResults.filter(({ rowData }) => extractMonth(rowData[0]) === transactionMonth);
  
  if (monthFilteredResults.length === 0) {
    Logger.log(`[DEBUG] No comparison data for ${transactionMonth}`);
    return;
  }

  const formattedRows = monthFilteredResults.map(({ rowIndex, status, rowData }) => [
    rowIndex, status, formatDate(rowData[0]), rowData[1], rowData[2], rowData[3],
    rowData[4], rowData[5], rowData[6], rowData[7], rowData[8]
  ]);

  // Compute column widths
  const columnWidths = calculateColumnWidths(formattedRows, tableHeaders);
  const headerRow = padRowColumns(tableHeaders, columnWidths);
  const dataRows = formattedRows.map(row => padRowColumns(row, columnWidths));

  // Log header separately before chunking rows
  Logger.log(`\n[DEBUG] Comparison Table for Payload #${payloadIndex + 1} (${transactionMonth})\n${headerRow}`);
  logInChunks('', dataRows);
};

// Checks if a transaction is a duplicate
function isDuplicateTransaction(transactionPayload, existingRows, payloadIndex, source) {
  if (!DEV_CONFIG.IDENTIFY_DUPLICATES) return false;

  // Define headers for the tables
  const tableHeaders = ['Index', 'Date', 'Amount', 'Type', 'Account', 'Category', 'Subcategory', 'Merchant', 'Title', 'Notes'];

  // Filter and normalize existing rows
  const normalizedRows = existingRows
    .filter(row => row[0] === CONSTANTS.SUCCESS_STATUS) // Keep only "Success" rows
    .map(row => row.slice(1, 10).map(normalizeField));

  // Extract the month from transactionPayload
  const transactionMonth = extractMonth(transactionPayload.date);

  // Log existing transactions table
  logExistingRowsTable(normalizedRows, tableHeaders, transactionMonth);

  const comparisonResults = [];

  // Compare each row with the transaction payload
  normalizedRows.forEach((row, rowIndex) => {
    const result = compareRowWithPayload(row, transactionPayload, rowIndex, source);
    comparisonResults.push(result);
  });

  // Add the payload row to the comparison results
  comparisonResults.unshift({
    rowIndex: `#${payloadIndex + 1}`,
    status: '-',
    rowData: [
      transactionPayload.date,
      transactionPayload.amount,
      transactionPayload.accountingType || '',
      transactionPayload.account,
      transactionPayload.category,
      transactionPayload.subcategory,
      transactionPayload.merchant || '',
      transactionPayload.title,
      transactionPayload.notes,
    ].map(normalizeField),
  });

  // Add 'Status' column only in the comparison table
  const comparisonHeaders = [...tableHeaders];
  comparisonHeaders.splice(1, 0, 'Status');

  // Log the comparison table
  buildComparisonTable(comparisonResults, comparisonHeaders, payloadIndex, transactionMonth);

  // Return whether a duplicate was found
  return comparisonResults.some(row => row.status === '✅');
}

// ,------.                                    ,------.        ,--.          ,--.            ,--. 
// |  .--. ' ,---.  ,---.  ,---. ,--.  ,--.    |  .--. ' ,---. |  | ,--,--.,-'  '-. ,---.  ,-|  | 
// |  '--'.'| .-. :| .-. || .-. : \  `'  /     |  '--'.'| .-. :|  |' ,-.  |'-.  .-'| .-. :' .-. | 
// |  |\  \ \   --.' '-' '\   --. /  /.  \     |  |\  \ \   --.|  |\ '-'  |  |  |  \   --.\ `-' | 
// `--' '--' `----'.`-  /  `----''--'  '--'    `--' '--' `----'`--' `--`--'  `--'   `----' `---'  

function handleStaticValueOrMatch(text, regex) {
  if (!(regex instanceof RegExp)) return regex; // Handle static values.
  return text.match(regex) || []; // Return matches or an empty array for no match.
}

function extractFullMatchOrFirstCaptureGroupString(text, regex) {
  const matches = handleStaticValueOrMatch(text, regex);
  if (!Array.isArray(matches)) return matches;  // static value case

  // Return first capture groups otherwise full match as string. Useful for account, amount, merchant, category, subcategory etc where only single value is expected.
  if(matches.length > 2) 
    logError(ErrorType.REGEX_SPILLOVER, `Excess matches found (${matches.length}) for : ${regex} - ${matches.join("--")}`, false);

  return matches[1] || matches[0] || "";
}

function extractFullMatchOrCaptureGroupsArray(text, regex) {
  const matches = handleStaticValueOrMatch(text, regex);
  if (!Array.isArray(matches)) return matches;

  // Return all capture groups (excluding the full match) or the full match.
  return matches.length > 1 ? matches.slice(1) : matches;
}

function getApplicableRegexMap(isSMS){
  var messageContent = isSMS ? emailData.messageBody : emailData.emailSubject;
  var ruleId = isSMS ? identifyMessageRule(messageContent) : emailSubjectToRuleIdMap[messageContent];

  if(!ruleId || ruleId == CONSTANTS.DUPLICATE) return ruleId; // Don't process SMS for which email is already received.
  Logger.log(`[RULE] ${ruleId} matched`);
  return isSMS ? smsRuleMap[ruleId] : emailRuleMap[ruleId];
}

function combineWithDefaultRegexMap(applicableRegexMap){
  return {
    ...applicableRegexMap,
    ...Object.fromEntries(
      Object.entries(DefaultRegexMap).filter(([key]) => !(key in applicableRegexMap))
    )
  }
}

function identifyMessageRule(messageContent) {
  const contentMatch = messageContent.match(/Transaction SMS\s*:\s*\n\s*([\s\S]*?)(?=\s*Received\s*:)/);
  if (!contentMatch) return null;  // Ideally should never execute unless SMS format is changed in Automate
  
  const smsContent = contentMatch[1].replace(/\s*\n\s*/g, ' ').replace(/\s+/g, ' ').trim(); // Remove newlines & extra spaces
  
  for (const [ruleId, patterns] of Object.entries(smsReferencePatterns)) {
    if (patterns.some(pattern => pattern.test(smsContent))) {
      return ruleId;
    }
  }
  
  return null;
}

// ___   _ _____ ___   _____ ___ __  __ ___   ___ ___  ___ __  __   _ _____ _____ ___ _  _  ___ 
// |   \ /_\_   _| __| |_   _|_ _|  \/  | __| | __/ _ \| _ \  \/  | /_\_   _|_   _|_ _| \| |/ __|
// | |) / _ \| | | _|    | |  | || |\/| | _|  | _| (_) |   / |\/| |/ _ \| |   | |  | || .` | (_ |
// |___/_/ \_\_| |___|   |_| |___|_|  |_|___| |_| \___/|_|_\_|  |_/_/ \_\_|   |_| |___|_|\_|\___|
                                                                                              
// Date/Time Format Constants
const DATE_FORMATS = deepFreeze({
  DISPLAY: {
    DATE: 'MMM DD YYYY',
    TIME: 'h:mm:ss A',
    DATETIME: "MMM DD YYYY, h:mm:ss A"
  },
  PARSE: {
    // CAUTION !! Keep YY variant before YYYY otherwise 24 become 0024
    DATE: ['MMM DD, YY', 'MMM DD, YYYY', 'DD-MM-YY', 'DD-MM-YYYY', 'DD/MM/YY', 'DD/MM/YYYY', 'MMM DD' ],
    TIME: ['HH:mm:ss', 'h:mm:ss A']
  },
  EMAIL: {
    DATE: 'MMM DD, YYYY',
    TIME: 'HH:mm:ss'
  },
  CASHEW_FORMAT: "YYYY-MM-DDTHH:mm:ss",
  // SUPPORTED FORMATS - https://github.com/jameskokoska/Cashew/blob/5.2.3%2B328/budget/lib/struct/commonDateFormats.dart
});

/**
 * Extracts date and time components from a text using the provided regex pattern.
 * Falls back to the provided date if extraction fails.
 * 
 * @param {string} text - The text to extract date/time from
 *                       e.g., "Transaction occurred on Dec 06, 2024 at 01:18:50"
 * @param {RegExp} dateRegex - Regex with two capture groups for date and time
 *                            e.g., /(\w{3} \d{2}, \d{4})\s*at\s*(\d{2}:\d{2}:\d{2})/
 * @param {moment.Moment} fallbackDate - Moment object to use if extraction fails
 * @returns {Object} Object containing datePart and timePart strings
*/
function extractDateComponents(text, dateRegex, fallbackDate) {
  const defaultComponents = {
    datePart: moment(fallbackDate).format(DATE_FORMATS.EMAIL.DATE),
    timePart: moment(fallbackDate).format(DATE_FORMATS.EMAIL.TIME)
  };

  if (!(dateRegex instanceof RegExp)) {
    logError(ErrorType.DATETIME_REGEX_INVALID, `Invalid Regex: ${dateRegex}`, false);
    return defaultComponents;
  }

  const matches = text.match(dateRegex);
  if (!matches) {
    logError(ErrorType.DATETIME_REGEX_INVALID, `No matches for dateRegex: ${dateRegex} while matching "${text}"`, false);
    return defaultComponents;
  }

  const [, extractedDate = '', extractedTime = ''] = matches;
  // removes full match, keeps capture groups
  // For regex: /(\w{3} \d{2}, \d{4})\s*at\s(\d{2}:\d{2}:\d{2})/
  // "Dec 06, 2024 at 01:18:50" gives:
  // matches[0]: "Dec 06, 2024 at 01:18:50"
  // matches[1]: "Dec 06, 2024"
  // matches[2]: "01:18:50"

  return {
    datePart: extractedDate || defaultComponents.datePart,
    timePart: extractedTime || defaultComponents.timePart
  };
}

/**
 * Validates the extracted date against supported formats and email date.
 * 
 * @param {string} datePart - The date string to validate e.g., "Dec 06, 2024"
 * @param {string} emailDateFormatted - Formatted email date for comparison
 * @param {moment.Moment} fallbackDate - Moment object to use if validation fails
 * @returns {moment.Moment} Valid moment object representing the date
 */
function validateDate(datePart, emailDateFormatted, fallbackDate) {
  const parsedDate = moment(datePart, DATE_FORMATS.PARSE.DATE, true);
  if (!parsedDate.isValid()) {
    logError(ErrorType.DATE_INVALID, `Invalid date format: ${datePart} expected ${DATE_FORMATS.PARSE.DATE}`, false);
    return fallbackDate;
  }

  if (parsedDate.format(DATE_FORMATS.EMAIL.DATE) !== emailDateFormatted) {
    logError(ErrorType.DATE_MISMATCH, `Date mismatch: Extracted(${parsedDate.format(DATE_FORMATS.EMAIL.DATE)}) vs Email(${emailDateFormatted})`, false);
  }

  return parsedDate;
}

/**
 * Validates the extracted time against supported formats.
 * 
 * @param {string} timePart - The time string to validate e.g., "01:18:50"
 * @param {moment.Moment} fallbackDate - Moment object to use if validation fails
 * @returns {moment.Moment} Valid moment object representing the time
 */
function validateTime(timePart, fallbackDate) {
  for (const format of DATE_FORMATS.PARSE.TIME) {
    const parsedTime = moment(timePart, format, true);
    if (parsedTime.isValid()) {
      return parsedTime;
    }
  }
  
  logError(ErrorType.TIME_INVALID, `Invalid time format: ${timePart} expected ${DATE_FORMATS.PARSE.TIME}`, false);
  return fallbackDate;
}

/**
 * Combines validated date and time into a formatted datetime string.
 * 
 * @param {moment.Moment} date - Validated date moment object
 * @param {moment.Moment} time - Validated time moment object
 * @returns {string} Formatted datetime string e.g., "Dec 06 2024, 1:18:50 AM"
 */

function formatDateTime(date, time) {
  // Ensure both date and time are Moment objects
  const momentDate = moment.isMoment(date) ? date : moment(date);
  const momentTime = moment.isMoment(time) ? time : moment(time);
  // Logger.log("[DEBUG] Date value: " + momentDate.valueOf());
  // Logger.log("[DEBUG] Time value: " + momentTime.valueOf());

  // Extract time components
  const hour = momentTime.get('hour');
  const minute = momentTime.get('minute');
  const second = momentTime.get('second');
  // Logger.log("[DEBUG] Extracted time: " + hour + ":" + minute + ":" + second);

  // Create a new moment object with the combined date and time
  const combinedDateTime = moment(momentDate).hour(hour).minute(minute).second(second);
  // Logger.log("[DEBUG] Combined DateTime: " + combinedDateTime.format());
  return combinedDateTime.format();
}

/**
 * Extracts and validates transaction datetime from text.
 * 
 * @param {string} text - Text containing date/time information
 *                       e.g., "Transaction occurred on Dec 06, 2024 at 01:18:50"
 * @param {RegExp} dateRegex - Regex to extract date and time components
 * @param {moment.Moment|string|Date} emailDate - Reference date for validation
 * @returns {string} Formatted datetime string e.g., "Dec 06 2024, 1:18:50 AM"
 */
function extractTransactionDate(text, dateRegex, emailDate) {
  const fallbackDate = moment(emailDate);
  const emailDateFormatted = fallbackDate.format(DATE_FORMATS.EMAIL.DATE);
  const { datePart, timePart } = extractDateComponents(text, dateRegex, fallbackDate);
  
  const validatedDate = validateDate(datePart, emailDateFormatted, fallbackDate);
  const validatedTime = validateTime(timePart, fallbackDate);
  // Logger.log("[DEBUG] Validated Date: " + validatedDate.format());
  // Logger.log("[DEBUG] Validated Time: " + validatedTime.format());
  
  return formatDateTime(validatedDate, validatedTime);
}

//    ______                   ______                                 __  _                ______                 __  _                 
//   / ____/___  ________     /_  __/________ _____  _________ ______/ /_(_)___  ____     / ____/_  ______  _____/ /_(_)___  ____  _____
//  / /   / __ \/ ___/ _ \     / / / ___/ __ `/ __ \/ ___/ __ `/ ___/ __/ / __ \/ __ \   / /_  / / / / __ \/ ___/ __/ / __ \/ __ \/ ___/
// / /___/ /_/ / /  /  __/    / / / /  / /_/ / / / (__  ) /_/ / /__/ /_/ / /_/ / / / /  / __/ / /_/ / / / / /__/ /_/ / /_/ / / / (__  ) 
// \____/\____/_/   \___/    /_/ /_/   \__,_/_/ /_/____/\__,_/\___/\__/_/\____/_/ /_/  /_/    \__,_/_/ /_/\___/\__/_/\____/_/ /_/____/  

function getTransactionAmount(text, amountRegex){
  var amountStr = extractFullMatchOrFirstCaptureGroupString(text, amountRegex);
  return parseFloat(amountStr.replace(/[^\d.-]/g, '')) || 0;
}

function getTransactionType(toAccount, isDebit){
  return toAccount ? TransactionType.TRANSFER : isDebit ? TransactionType.DEBIT : TransactionType.CREDIT;
}

function getTransactionCategory(data, emailBody, isDebit) {
  // Clean and combine the data and email body
  var combinedText = (data + " " + emailBody).toLowerCase();

  // Select the appropriate section (expenses or incomes)
  var categoryType = isDebit ? CategoryType.EXPENSES : CategoryType.INCOMES;
  var categoryMap = categorySubcategoryKeywordMap[categoryType];

  // Iterate over categories
  for (var category in categoryMap) {
    var categoryData = categoryMap[category];

    // Check subcategories first, if they exist
    if (categoryData.subcategories) {
      for (var subcategory in categoryData.subcategories) {
        var subcategoryKeywords = categoryData.subcategories[subcategory];
        if (subcategoryKeywords.some(keyword => combinedText.includes(keyword.toLowerCase()))) {
          Logger.log(`[CATEGORY] Matched ${categoryType} - Category: ${category}, Subcategory: ${subcategory}`);
          return { categoryType, category, subcategory };
        }
      }
    }

    // Check category-level keywords if no subcategory match
    if (categoryData.keywords && categoryData.keywords.some(keyword => combinedText.includes(keyword.toLowerCase()))) {
      Logger.log(`[CATEGORY] Matched ${categoryType} - Category: ${category}, No Subcategory`);
      return { categoryType, category, subcategory: null };
    }
  }

  // Default if no match is found
  logError(ErrorType.NO_CATEGORY, `No matching ${categoryType} category found for ${data}. Using default category ${isDebit ? USER_DEFAULTS.EXPENSE_CATEGORY : USER_DEFAULTS.INCOME_CATEGORY}`, false);    // silent error
  return {
    categoryType,
    category: isDebit ? USER_DEFAULTS.EXPENSE_CATEGORY : USER_DEFAULTS.INCOME_CATEGORY,
    subcategory: null
  };
}

// Utility function to determine if the transaction is debit or credit
function isDebitTransaction(emailBody, debitRegex, toAccount, transactionAmount) {
  if (typeof debitRegex === 'boolean') {
    return debitRegex; // If isDebit is directly set to true/false, return it
  }
  
  if (debitRegex instanceof RegExp) {
    return debitRegex.test(emailBody); // Match against regex and return true if it matches
  }

  if(toAccount) {
    return true;
  } else {
    return transactionAmount < 0;
  }
}

/**
 * Utility to get the account name based on keywords in the data or email body.
 * If no match is found, defaults to "Wallet".
 * 
 * @param {string} data - Extracted data from transaction.
 * @param {string} emailBody - Email body text.
 * @return {string} - Account name.
 */
function getAccountName(data, emailBody) {
  if (!data) return data; // Fail fast for empty "toAccount";
  
  var combinedText = (data + " " + emailBody).toLowerCase();

  if (!accountKeywordMap || Object.keys(accountKeywordMap).length === 0) {
    Logger.log(`[ERROR] accountKeywordMap is empty or undefined!`);
    return USER_DEFAULTS.ACCOUNT;
  }

  let bestMatch = null; 

  // First pass: Look for exact matches first
  for (var accountName in accountKeywordMap) {
    var keywords = accountKeywordMap[accountName];

    if (keywords.includes(data)) {  // Exact match with `data`
      Logger.log(`[ACCOUNT] Exact match found for: ${data} -> ${accountName}`);
      return accountName;  // Prioritize exact matches
    }
  }

  // Second pass: Look for keyword matches in combined text
  for (var accountName in accountKeywordMap) {
    var keywords = accountKeywordMap[accountName];

    if (keywords.some(keyword => combinedText.includes(keyword.toLowerCase()))) {
      Logger.log(`[ACCOUNT] Matched Account: ${accountName}`);
      bestMatch = accountName;
    }
  }

  if (bestMatch) {
    return bestMatch;  // Use first keyword match found
  }

  logError(ErrorType.NO_ACCOUNT, `No matching account found for ${data}. Defaulting to ${USER_DEFAULTS.ACCOUNT}.`, false);
  return USER_DEFAULTS.ACCOUNT;
}

// Generate transaction title
function generateTitle(merchant, text) {
  // Remove special characters
  return merchant && merchant.replace(/[^a-zA-Z0-9 ]/g, '');
  // return merchant || "Transaction at " + text.split(" ")[0];
}

function extractNotes(text, regex){
  var notes = extractFullMatchOrCaptureGroupsArray(text, regex);
  // Handling multiple capture groups OR nonRegexMatch.
  return Array.isArray(notes) ? notes.join("-") : notes;
}

function getTransactionSubcategory(sanitizedText, subcategoryRegex, subcategory){
  if(subcategory) return subcategory;
  subcategory = extractFullMatchOrFirstCaptureGroupString(sanitizedText, subcategoryRegex);
  if(!subcategory) return subcategory;  // If no match is found then no use of checking keywords.

  for (var subcategoryName in subcategoryKeywordMap) {
    var keywords = subcategoryKeywordMap[subcategoryName];
    if (keywords.some(keyword => subcategory.toLowerCase().includes(keyword.toLowerCase()))) {
      Logger.log(`[Subcategory] Matched subcategory: ${subcategoryName}`);
      return subcategoryName;
    }
  }

  // If subcategory is not found in match then it'll surface while adding the transaction. Might also be a valid case of nonRegexMatch
  return subcategory;
}

//  _   _ ___ _       ___                       _   _          
// | | | | _ \ |     / __|___ _ _  ___ _ _ __ _| |_(_)___ _ _  
// | |_| |   / |__  | (_ / -_) ' \/ -_) '_/ _` |  _| / _ \ ' \ 
//  \___/|_|_\____|  \___\___|_||_\___|_| \__,_|\__|_\___/_||_|

function createFinalTransactionUrl(filteredPayload) {
  var encodedPayload = encodeURIComponent(JSON.stringify(filteredPayload));
  return `https://${CONFIG.WEB_DOMAIN}/${CONFIG.ADD_ROUTE}?JSON=${encodedPayload}`;
}

function createSingleTransactionUrl(filteredPayload) {
  // Encode individual query parameters
  const queryString = Object.entries(filteredPayload)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');

  // Return the constructed URL
  return `https://${CONFIG.WEB_DOMAIN}/${CONFIG.EDIT_ROUTE}?${queryString}`;
}

// Create the transaction URL with JSON payload
function filterPayload(payload){
  var filteredPayload = {};
  for (var key in payload) {
    if (payload[key] !== null && payload[key] !== "") {
      filteredPayload[key] = payload[key];
    }
  }
  return filteredPayload;
}

// ██       ██████   ██████   ██████  ██ ███    ██  ██████  
// ██      ██    ██ ██       ██       ██ ████   ██ ██       
// ██      ██    ██ ██   ███ ██   ███ ██ ██ ██  ██ ██   ███ 
// ██      ██    ██ ██    ██ ██    ██ ██ ██  ██ ██ ██    ██ 
// ███████  ██████   ██████   ██████  ██ ██   ████  ██████  

function logError(errorType, errorMessage, sendErrorEmailFlag) {
  const { key, isStopping } = errorType;

  Logger.log(`[ERROR] (${key}) : ${errorMessage}`);

  // Increment specific error count
  ProcessedCount.errored.details[key] = (ProcessedCount.errored.details[key] || 0) + 1;
  ProcessedCount.errored.totalErrors++;

  // Increment stopping or silent error count based on the enum
  if (isStopping) {
    ProcessedCount.errored.stoppingErrors++;
  } else {
    currentTransactionSilentErrors.push(key);
    ProcessedCount.errored.silentErrors++;
  }

  if(sendErrorEmailFlag){
    var subject = `TPS Error [${key}] (${isStopping ? "Stopping" : "Silent"})`
    // Minimal version with plain text.
    // var body = `Error processing Email: ${getEmailThreadLink()}\nSubject: ${emailData.emailSubject}\nEmail Date: ${moment(emailData.emailDate).format(DATE_FORMATS.DISPLAY.DATETIME)}\nSource : ${emailData.source}\nError : ${errorMessage}\n\nOriginal Body : \n---\n${emailData.messageBody}\n---\n`;
    // GmailApp.sendEmail(CONFIG.EMAIL, subject, body);
    sendErrorEmail(subject, errorMessage);
  }
}

function printProcessingSummary(logError = true) {
  let summary = '';
  
  summary += `[INFO]  Total Processed: ${ProcessedCount.TOTAL}.\n`;
  summary += `Total Success: ${ProcessedCount.SUCCESS}, Skipped: ${ProcessedCount.SKIPPED}\n`;
  summary += `[ERROR] Total Errors: ${ProcessedCount.errored.totalErrors}\n`;

  if (ProcessedCount.errored.stoppingErrors) {
    summary += `[ERROR] Stopping Errors: ${ProcessedCount.errored.stoppingErrors}\n`;
    Object.entries(ErrorType)
      .filter(([_, error]) => error.isStopping)
      .filter(([_, error]) => (ProcessedCount.errored.details[error.key] || 0) > 0)
      .forEach(([key, error]) => {
        summary += `  - ${key}: ${ProcessedCount.errored.details[error.key]}\n`;
      });
  }

  if (ProcessedCount.errored.silentErrors) {
    summary += `[ERROR] Silent Errors: ${ProcessedCount.errored.silentErrors}\n`;
    Object.entries(ErrorType)
      .filter(([_, error]) => !error.isStopping)
      .filter(([_, error]) => (ProcessedCount.errored.details[error.key] || 0) > 0)
      .forEach(([key, error]) => {
        summary += `  - ${key}: ${ProcessedCount.errored.details[error.key]}\n`;
      });
  }

  if(logError) {
    Logger.log(summary);
  } else {
    summary = summary.replaceAll('\n', '<br/>');
  }
  return summary;
}

function createFailureRecord(status) {
  if(!DEV_CONFIG.CREATE_FAILURE_RECORD) return;
  // Get the last row in the sheet to append data
  var lastRow = transactionSheet.getLastRow() + 1; // Row number for the next empty row
  
  // Write the status, emailId, and emailSubject in the next row
  transactionSheet.getRange(lastRow, 1).setValue(status); // Error
  transactionSheet.getRange(lastRow, 2).setValue(moment(emailData.emailDate).format(DATE_FORMATS.DISPLAY.DATETIME)); // Email date
  transactionSheet.getRange(lastRow, 3).setValue(emailData.messageBody); // View email body inline
  transactionSheet.getRange(lastRow, 12).setValue(emailData.emailId); // So that it is not picked again unless manually fixed & removed.
  transactionSheet.getRange(lastRow, 13).setValue(getEmailThreadLink());  // Direct link to email thread
  transactionSheet.getRange(lastRow, 14).setValue("-"); // Transaction link could not be created
  transactionSheet.getRange(lastRow, 15).setValue(emailData.emailSubject); // Subject
  transactionSheet.getRange(lastRow, 16).setValue(getToday()); // Processed Date
}

//  _   _ _   _ _ _ _          ___             _   _               
// | | | | |_(_) (_) |_ _  _  | __|  _ _ _  __| |_(_)___ _ _  ___  
// | |_| |  _| | | |  _| || | | _| || | ' \/ _|  _| / _ \ ' \(_-<  
//  \___/ \__|_|_|_|\__|\_, | |_| \_,_|_||_\__|\__|_\___/_||_/__/  
//                      |__/                                       

const getToday = () => {
  const now = new Date();
  const options = { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false };
  return now.toLocaleString('en-US', options).replace(',', '');
};

function getEmailThreadLink(){
  return CONFIG.GMAIL_URL_PREFIX + emailData.emailId;
}

// WORKING But takes time to open - hence not using currently
// Search email using labels.
function getMessageId(emailId = emailData.emailId) {
  var email = GmailApp.getMessageById(emailId);
  var messageId = email.getHeader("Message-ID");
  return encodeURI(CONFIG.GMAIL_MSGID_PREFIX + messageId);
}

// Recursive deep freeze for objects and arrays
function deepFreeze(obj) {
  // Check if the object is a non-null object (including arrays)
  if (obj && typeof obj === 'object') {
    // Freeze the object itself
    Object.freeze(obj);

    // Iterate over all properties in the object
    Object.keys(obj).forEach(key => {
      // If the property is an object or array, recursively freeze it
      deepFreeze(obj[key]);
    });
  }

  return obj;
}

// ,---.                               ,--.      ,--.   ,--.                                              
// '   .-'  ,---.  ,---.,--.--. ,---. ,-'  '-.    |   `.'   | ,--,--.,--,--,  ,--,--. ,---.  ,---. ,--.--. 
// `.  `-. | .-. :| .--'|  .--'| .-. :'-.  .-'    |  |'.'|  |' ,-.  ||      \' ,-.  || .-. || .-. :|  .--' 
// .-'    |\   --.\ `--.|  |   \   --.  |  |      |  |   |  |\ '-'  ||  ||  |\ '-'  |' '-' '\   --.|  |    
// `-----'  `----' `---'`--'    `----'  `--'      `--'   `--' `--`--'`--''--' `--`--'.`-  /  `----'`--'    
//                                                                                   `---'                

/**
 * Sets secrets from a JSON object into Script Properties
 * @param {Object} secrets - JSON object containing secrets
 */
function storeSecrets(secrets) {
  const properties = PropertiesService.getScriptProperties();
  
  for (const [key, value] of Object.entries(secrets)) {
    // Convert arrays to JSON strings for storage
    const valueToStore = Array.isArray(value) ? JSON.stringify(value) : value;
    properties.setProperty(key, valueToStore);
  }
  Logger.log("[RESPONSE] Secrets stored/updated in properties.");
}

/**
 * Retrieves all secrets from Script Properties and returns them as a JSON object
 * @returns {Object} JSON object containing all secrets
 */
function fetchSecrets() {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  const secrets = {};
  
  for (const [key, value] of Object.entries(allProperties)) {
    try {
      // Attempt to parse as JSON (for arrays)
      secrets[key] = JSON.parse(value);
    } catch (e) {
      // If parsing fails, store as regular string
      secrets[key] = value;
    }
  }
  Logger.log("[SETUP] Secrets retrieved from properties.");
  return secrets;
}

// Example usage:
// function example() {
//   // Sample secrets JSON
//   const mySecrets = {
//     EMAIL: "user1@example.com",
//     SPREADSHEET_ID: "1234",
//     ACCOUNT_IDENTIFIERS: ["XXXX1234", "ACCOUNT_KEYWORD"],
//   };
  
//   // Store secrets
//   storeSecrets(mySecrets);
  
//   // Retrieve secrets
//   const retrievedSecrets = fetchSecrets();
//   Logger.log(retrievedSecrets.EMAIL); // "user1@example.com"
//   Logger.log(retrievedSecrets.ACCOUNT_IDENTIFIERS); // ["XXXX1234", "ACCOUNT_KEYWORD"]
// }

/**
 * Enriches an account keyword map with additional identifiers from secrets and normalizes account keys
 * 
 * @param {Object} baseAccountIdsMap - Base map of account keys to their identifier arrays
 *   Example: { "ACCOUNT_SAVINGS": ["Bank Name"], "CC": ["Credit Card"] }
 * 
 * @param {Object} secrets - Secrets object containing additional account identifiers
 *   Each account's identifiers should be stored with key format: ACCOUNT_ID_<ACCOUNT_KEY>
 *   Example: { "ACCOUNT_ID_ACCOUNT_SAVINGS": ["additional", "identifiers"] }
 * 
 * @param {Object} accountIdsToAccountNameMap - Mapping of account keys to their normalized names
 *   Example: { "ACCOUNT_SAVINGS": "acc1" }
 * 
 * @returns {Object} Enriched and immutable account keyword map
*/

function enrichMapWithSecrets(baseAccountIdsMap, accountIdsToAccountNameMap) {
  // Create a new object for the enriched map
  const enrichedMap = {};
  
  // Iterate through the base map keys
  Object.keys(baseAccountIdsMap).forEach(accountKey => {
    // Determine the final key using the name mapping or keep original
    const finalKey = accountIdsToAccountNameMap[accountKey] || accountKey;
    
    // Initialize the array with values from base map
    enrichedMap[finalKey] = [...baseAccountIdsMap[accountKey]];
    
    // Construct the secret key
    const secretKey = `ACCOUNT_ID_${accountKey.toUpperCase()}`;
    
    // Add values from secrets if they exist
    if (SECRETS[secretKey]) {
      enrichedMap[finalKey].push(...SECRETS[secretKey]);
    }
    
    // Deduplicate the array
    enrichedMap[finalKey] = [...new Set(enrichedMap[finalKey])];
  });
  
  // Return the enriched map
  // Note: Object.freeze() will be applied by the caller
  return enrichedMap;
}

// ███████ ███    ███  █████  ██ ██      
// ██      ████  ████ ██   ██ ██ ██      
// █████   ██ ████ ██ ███████ ██ ██      
// ██      ██  ██  ██ ██   ██ ██ ██      
// ███████ ██      ██ ██   ██ ██ ███████ 

// Utility function for shared styles
function getEmailStyles() {
  return `
    /* Shared styles for both success and error templates */
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      margin: 0;
      padding: 20px;
    }

    table {
      border-collapse: collapse;
      width: 100%;
      margin: 20px 0;
    }

    th, td {
      border: 1px solid #ddd;
      padding: 12px;
      text-align: left;
    }

    th {
      background-color: #f5f5f5;
    }

    /* Success template specific styles */
    .success-row-credit {
      background-color: #e8f5e9;
    }

    .success-row-debit {
      background-color: #fff3e0;
    }

    .success-row-duplicate {
      background-color: #ecb4b4;
    }

    /* Error template specific styles */
    .error-message {
      background-color: #ffebee;
      border: 1px solid #ffcdd2;
      padding: 15px;
      margin: 10px 0;
      border-radius: 4px;
    }

    /* Button styles for both templates */
    .action-buttons {
      margin: 15px 0;
      display: flex;
      gap: 15px;
      align-items: center;
    }

    .button-label {
      font-weight: bold;
      margin-right: 10px;
    }

    .action-link {
      display: inline-block;
      padding: 8px 15px;
      background-color: #4caf50;
      color: white;
      text-decoration: none;
      border-radius: 4px;
      margin-right: 10px;
    }

    .edit-link {
      background-color: #2196f3;
    }

    /* Mobile responsive styles */
    @media screen and (max-width: 768px) {
      body {
        padding: 0;
      }

      table {
        font-size: 14px;
      }

      th, td {
        padding: 8px;
      }

      /* Max widths for specific columns on mobile */
      .date-column, 
      .category-column, 
      .notes-column, 
      .errors-column {
        max-width: 150px;
        white-space: normal;
        word-wrap: break-word;
      }
    }
  `;
}

function getSuccessEmailHtml(finalTransactionUrl, transactions, metadata) {
  const tableRows = transactions.map((transaction, index) => {
    const { source, silentErrors, txnType } = metadata[index];
    let classId = `success-row-${txnType.toLowerCase()}`; // success-row-<debit/credit/duplicate>
    return `
      <tr class="${classId}">
        <td>${source}</td>
        <td class="date-column">${moment(transaction.date).format(DATE_FORMATS.DISPLAY.DATETIME)}</td>
        <td>${transaction.amount}</td>
        <td>${transaction.account}</td>
        <td>${transaction.title || ""}</td>
        <td class="category-column">${transaction.category}</td>
        <td>${transaction.subcategory || ""}</td>
        <td class="notes-column">${transaction.notes || ""}</td>
        <td class="errors-column">${silentErrors || ""}</td>
      </tr>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>${getEmailStyles()}</style>
    </head>
    <body>
      <p>Below are the details of the transactions processed:</p>
      
      <div class="action-buttons">
        <span class="button-label">PC:</span>
        <a href="${finalTransactionUrl}" class="action-link">✅ Approve</a>
        <a href="${finalTransactionUrl.replace(CONFIG.ADD_ROUTE, CONFIG.EDIT_ROUTE)}" class="action-link edit-link">Edit ✍️</a>
      </div>

      <table>
        <thead>
          <tr>
            <th>Source</th>
            <th>Date</th>
            <th>Amount</th>
            <th>Account</th>
            <th>Title</th>
            <th>Category</th>
            <th>Subcategory</th>
            <th>Notes</th>
            <th>Silent Errors</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      <div class="action-buttons">
        <span class="button-label">Mobile:</span>
        <a href="${finalTransactionUrl.replace(CONFIG.WEB_DOMAIN, CONFIG.MOBILE_DOMAIN)}" class="action-link">✅ Approve</a>
        <a href="${finalTransactionUrl.replace(CONFIG.WEB_DOMAIN, CONFIG.MOBILE_DOMAIN).replace(CONFIG.ADD_ROUTE, CONFIG.EDIT_ROUTE)}" class="action-link edit-link">Edit ✍️</a>
      </div>
      <pre style="white-space: pre-wrap; word-wrap: break-word; background: #f5f5f5; padding: 15px; border-radius: 4px;">${printProcessingSummary(false)}</pre>
    </body>
    </html>
  `;
}

function getErrorEmailHtml(errorMessage) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>${getEmailStyles()}</style>
    </head>
    <body>
      <h2>Transaction Processing Error</h2>
      <div class="error-message">
        <p><strong>Error Message:</strong> ${errorMessage}</p>
      </div>

      <h3>Email Details</h3>
      <table>
        <tr>
          <td><strong>Email Link:</strong></td>
          <td><a href="${getEmailThreadLink()}">View Original Email</a></td>
        </tr>
        <tr>
          <td><strong>Subject:</strong></td>
          <td>${emailData.emailSubject}</td>
        </tr>
        <tr>
          <td><strong>Date:</strong></td>
          <td>${moment(emailData.emailDate).format(DATE_FORMATS.DISPLAY.DATETIME)}</td>
        </tr>
        <tr>
          <td><strong>Source:</strong></td>
          <td>${emailData.source}</td>
        </tr>
      </table>

      <h3>Original Message Body</h3>
      <pre style="white-space: pre-wrap; word-wrap: break-word; background: #f5f5f5; padding: 15px; border-radius: 4px;">
${emailData.messageBody}
      </pre>
    </body>
    </html>
  `;
}

// Updated send functions
function sendSuccessEmail(finalTransactionUrl, transactions, metadata) {
  const htmlBody = getSuccessEmailHtml(finalTransactionUrl, transactions, metadata);
  GmailApp.sendEmail(CONFIG.EMAIL, CONSTANTS.SUCCESS_SUBJECT, "", {
    htmlBody: htmlBody
  });
  Logger.log('[INFO] Success Email Sent.');
}

function sendErrorEmail(subject, errorMessage) {
  if (!DEV_CONFIG.SEND_ERROR_MAIL) return;

  const htmlBody = getErrorEmailHtml(errorMessage);
  GmailApp.sendEmail(CONFIG.EMAIL, subject, "", {
    htmlBody: htmlBody
  });
  Logger.log('[INFO] Error Email Sent.');
}