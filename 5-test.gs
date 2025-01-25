// Dev Config
const DEV_CONFIG = {
    // Send Mail & make entry in sheet
    SEND_ERROR_MAIL: true, // send email for stopping/silent error
    CREATE_FAILURE_RECORD: true, // Create failure record in sheet.

    // There are multiple filters. 
    // 1. Any email will only process if it is unread and 
    // 2. Any LABELS.INCLUDE label is applied and Any LABELS.EXCLUDE is not applied and
    // 3. It is not already present in current sheet.
    RERUN_READ_MAILS: false, // Bypass condition 1. True - Process already Read mails again. 
    RERUN_PROCESSED_IN_SHEET: true, // Bypass condition 3. False - Skip email if emailID is present in OUTPUT Sheet. True - rerun without deleting failed entry.
    RERUN_PROCESSED_IN_EMAIL: false, // Alter condition 2 to also include LABELS.PROCESSED emails. 
    // False - Exclude LABELS.PROCESSED in search query. True - [CAUTION !!] Pick entire history.

    // After email is processed if we mark it as read & apply LABELS.PROCESSED then it can't be rerun.
    MARK_AS_PROCESSED: true, // True - Mark as read and apply LABELS.PROCESSED. False - Keep it unchanged to allow rerun.
    IDENTIFY_DUPLICATES: true, // True - check all existing entries and skip entry from final URL if same details are already present. False - Ignore existing.

    SANITY_TESTS_RUN: false, // During sanity testing, Only LABELS.TESTCASES are run.
    OUTPUT_SHEET_TITLE: CONFIG.MAIN_SHEET_NAME, // Which sheet shall the result be written to.
}

// To Mock WebApp API hit
// https://developers.google.com/apps-script/guides/web#request_parameters
function testDoGet() {
    const mockRequest = {
        parameter: {
            param1: "value1"
        },
        mockUrl: "https://script.google.com/macros/s/your-script-id/dev" // Simulate /dev or /exec
    };

    const response = doGet(mockRequest);
    Logger.log(response.getContent());
}

// Explicitly run for these
// Ignore these values for time driven Trigger - https://developers.google.com/apps-script/guides/triggers/events#time-driven-events
function getTestThreadsOrQuery() {
    DEV_CONFIG.IDENTIFY_DUPLICATES = false;
    // return ['193f3d90e07f3123'];
    // return null; // Default: No test data available

    // Uncomment as needed for testing:
    // return 'subject:(Transaction SMS [12 Dec])';
    // return 'subject:(bob world - transaction alert) after:2024/10/01 before:2024/10/30';
    // return 'subject:(bob world - transaction alert) after:2024/11/15 before:2024/11/16';
    // return ['193ba6804fedec28', '193ba6804fedec29']; // Example email IDs
}


// Webapp API endoint for /exec & /dev GET requests
function doGet(e) {
    const mockUrl = e?.mockUrl; // Use mockUrl if provided
    const isDev = mockUrl ? mockUrl.includes("/dev") : isDevEnvironment(e);

    if (isDev) {
        Logger.log("[SETUP] Running /dev route with latest saved changes. Reprocessing read ones");
        DEV_CONFIG.RERUN_READ_MAILS = true;
        DEV_CONFIG.RERUN_PROCESSED_IN_SHEET = true;
        DEV_CONFIG.MARK_AS_PROCESSED = false;
    } else {
        Logger.log("[SETUP] Running /exec route with last deployed changes (On HTTP GET)");
    }
    try {
        // Call your existing method to process emails
        const result = processTransactionEmails();

        // Return a success message
        return ContentService.createTextOutput(
            JSON.stringify({
                mode: isDev ? "Development Mode" : "Production Mode",
                status: "success",
                message: "Emails processed",
                data: result
            })
        ).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        // Return error information
        return ContentService.createTextOutput(
            JSON.stringify({
                status: "error",
                message: error.message
            })
        ).setMimeType(ContentService.MimeType.JSON);
    }
}

function isDevEnvironment(e) {
    const url = ScriptApp.getService().getUrl(); // Base script URL
    const isDev = url.includes('/dev'); // Check if the URL contains '/dev'
    // Logger.log("[INFO] Parameter 1: " + e.parameter.param1); // Access query params like this.
    return isDev;
}

// Example: Returning an HTML Page
// function doGet(e) {
//     return HtmlService.createHtmlOutput("<h1>Welcome to My App</h1>");
//   }

// Google Apps Script only supports a single doGet or doPost. Using ?route= is the simplest way to define routes
// function doGet(e) {
//     const route = e.parameter.route || "default"; // e.parameter contains the query parameters
//     switch (route) {
//       case "testEndpoint":
//         return handleTestEndpoint();
//       default:
//         return handleDefault();
//     }
// }

// Sample Post route. UseCases - API for Drive/Sheets, Webhook etc.
// function doPost(e) {
//     // Log the incoming POST request for debugging
//     Logger.log(e.postData.contents);

//     // Parse JSON if the incoming payload is in JSON format
//     const requestData = JSON.parse(e.postData.contents);

//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
//     sheet.appendRow([requestData.id, requestData.name, requestData.value]);
//     return ContentService.createTextOutput(
//         JSON.stringify({
//             success: true,
//             message: "Row added successfully"
//         })
//     ).setMimeType(ContentService.MimeType.JSON);
// }
// curl -X POST -H "Content-Type: application/json" \
// -d '{"param1":"value1","param2":"value2"}' \
// https://script.google.com/macros/s/your-script-id/exec     

// To run for all "LABELS.TESTCASES" emails, store their results in CONFIG.RESULTS_SHEET_TITLE & Compare results with CONFIG.EXPECTED_SHEET_TITLE
function runSanityTest() {
    // Explicitly set dev config for sanity tests.
    setSanityDevConfig();

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const expectedSheet = spreadsheet.getSheetByName(CONFIG.EXPECTED_SHEET_TITLE);
    const outputSheet = spreadsheet.getSheetByName(CONFIG.RESULTS_SHEET_TITLE);

    // Clear previous output
    outputSheet.clear();

    // Set headers in Output sheet
    const expectedHeaders = expectedSheet.getRange(1, 1, 1, expectedSheet.getLastColumn()).getValues()[0];
    outputSheet.appendRow(expectedHeaders);

    // Process emails and write to Output sheet
    processTransactionEmails();

    // Fetch all rows from Output and Expected sheets
    const outputData = outputSheet.getDataRange().getValues();
    const expectedData = expectedSheet.getDataRange().getValues();

    // Match and compare rows
    compareRows(outputData, expectedData, outputSheet, expectedSheet);
}

// ,--------.              ,--.      ,--.  ,--.       ,--.                             
// '--.  .--',---.  ,---.,-'  '-.    |  '--'  | ,---. |  | ,---.  ,---. ,--.--. ,---.  
//    |  |  | .-. :(  .-''-.  .-'    |  .--.  || .-. :|  || .-. || .-. :|  .--'(  .-'  
//    |  |  \   --..-'  `) |  |      |  |  |  |\   --.|  || '-' '\   --.|  |   .-'  `) 
//    `--'   `----'`----'  `--'      `--'  `--' `----'`--'|  |-'  `----'`--'   `----'  
//                                                        `--'                         
/**
 * Retrieves email threads based on different execution contexts:
 * - Time-driven trigger : Unread & (LABELS.INCLUDED - LABELS.EXCLUDED)
 * - Sanity test run : LABELS.TESTCASES
 * - Specific email IDs : [threadIds]
 * - Search query : specific search query
 * 
 * @param {Object} e - Trigger event object (optional)
 * @returns {GmailThread[]} Array of Gmail threads
*/

function getEmailThreads(e) {
    let threads;
    // Ignore testQuery if time-driven trigger
    if (e?.triggerUid) {
        Logger.log('[SETUP] Running for Unread & (LABELS.INCLUDED - LABELS.EXCLUDED) since time-driven trigger');
        DEV_CONFIG.RERUN_READ_MAILS = false;
        DEV_CONFIG.RERUN_PROCESSED_IN_SHEET = false;
        DEV_CONFIG.RERUN_PROCESSED_IN_EMAIL = false;
        DEV_CONFIG.MARK_AS_PROCESSED = true;
        threads = GmailApp.search(getSearchQuery());
    } else if (DEV_CONFIG.SANITY_TESTS_RUN) {
        // Also ignore in case of sanity test RUN as test
        Logger.log(`[SETUP] Running for ${LABELS.TESTCASES} during sanity test.`);
        threads = GmailApp.search(getSearchQuery());
    } else {
        const threadsQuery = getTestThreadsOrQuery();
        
        if (Array.isArray(threadsQuery)) {
            Logger.log('[INFO] Fetching theads for explicitly provided emailIDs');
            threads = getEmailThreadsByIds(threadsQuery);
        } else {
            const searchQuery = threadsQuery || getSearchQuery();
            Logger.log(`[INFO] Fetching emails using ${threadsQuery ? 'default' : 'modified'} query: ${searchQuery}`);
            threads = GmailApp.search(searchQuery);
        }
    }

    return threads;
}

function getEmailThreadsByIds(emailIds) {
    const emailThreads = new Set();

    emailIds.forEach(function (emailId) {
        try {
            // Retrieve the email message and associated thread
            const email = GmailApp.getMessageById(emailId);
            const thread = email.getThread();

            // Add the thread to the set
            emailThreads.add(thread);
        } catch (error) {
            Logger.log(`[Error] Failed to retrieve thread for emailId: ${emailId}. Error: ${error.message}`);
        }
    });

    // Convert the Set back to an array before returning
    return Array.from(emailThreads);
}

function getSearchQuery() {
    if (DEV_CONFIG.SANITY_TESTS_RUN) return `label:"${LABELS.TESTCASES}"`;

    var unreadCondition = DEV_CONFIG.RERUN_READ_MAILS ? '' : 'is:unread';
    if (DEV_CONFIG.RERUN_PROCESSED_IN_EMAIL) {
        LABELS.EXCLUDE.shift();
    }

    var searchIncludeLabels = LABELS.INCLUDE.map(label => `label:"${label}"`).join(" OR ");
    var searchExcludeLabels = LABELS.EXCLUDE.map(label => `label:"${label}"`).join(" OR ");
    return `${unreadCondition} (${searchIncludeLabels}) NOT (${searchExcludeLabels})`;
}

function setSanityDevConfig() {
    DEV_CONFIG.SEND_ERROR_MAIL = false; // Manually check what test cases failed.
    DEV_CONFIG.CREATE_FAILURE_RECORD = true; // Output results must have entry to compare with expected.
    DEV_CONFIG.RERUN_READ_MAILS = true; // Process already read emails again & Don't mark as read.
    DEV_CONFIG.RERUN_PROCESSED_IN_SHEET = true; // Ignore if emailId is already present in output sheet. Ideally not required since emailIds are unique.
    DEV_CONFIG.MARK_AS_PROCESSED = false; // Don't change read status or apply labels.
    DEV_CONFIG.SANITY_TESTS_RUN = true; // To pick only LABEL.TESTCASES in searchQuery.
    DEV_CONFIG.OUTPUT_SHEET_TITLE = CONFIG.RESULTS_SHEET_TITLE; // Write in TestResults sheet instead of Main sheet.
}

const TestResultsBgColors = deepFreeze({
    EXACT_MATCH: "#00FF00", // green
    MISMATCH: "#FF5733",    // red
    NEW_MATCH: "#ADD8E6",   // blue
    NEW_FAILED: "#CF9FFF",  // purple
});

function compareRows(outputData, expectedData, outputSheet, expectedSheet) {
    // TODO: Create array of headers and find column from there. kanpilotID(pk4rmg2qyufx8j971pjw3t5d)
    const emailIdIndex = 12;

    // Create a map of expected rows by emailId for quick lookup
    const expectedMap = new Map();
    expectedData.slice(1).forEach((row) => {
        const key = row.slice(0, -1); // Exclude last column (current date)
        expectedMap.set(row[emailIdIndex], key);
    });

    // Compare rows and apply formatting
    for (let i = 1; i < outputData.length; i++) {
        const processedTime = outputData[i].pop(); // Exclude last column (current date)
        const outputRow = outputData[i];
        const emailId = outputRow[emailIdIndex];
        const expectedRow = expectedMap.get(emailId);

        const range = outputSheet.getRange(i + 1, 1, 1, outputRow.length + 1); // Include last column in formatting

        if (expectedRow) {
            // Compare with Expected
            const isExactMatch = JSON.stringify(outputRow) === JSON.stringify(expectedRow);

            if (isExactMatch) {
                range.setBackground(TestResultsBgColors.EXACT_MATCH);
            } else {
                range.setBackground(TestResultsBgColors.MISMATCH);
                for (let j = 0; j < outputRow.length; j++) {
                    if (isMismatch(outputRow[j], expectedRow[j])) {
                        // Logger.log("[ERROR] Mismatch values "+outputRow[j]+" "+expectedRow[j]);
                        range.getCell(1, j + 1).setFontWeight("bold");
                    }
                }
            }
        } else {
            // Handle extra rows
            if (outputRow[0] === CONSTANTS.SUCCESS_STATUS) {
                range.setBackground(TestResultsBgColors.NEW_MATCH);
                expectedSheet.appendRow(outputRow.concat(processedTime)); // Add back removed date
            } else {
                range.setBackground(TestResultsBgColors.NEW_FAILED);
            }
        }
    }
}

function isMismatch(outputCell, expectedCell) {
    return formatCell(outputCell) !== formatCell(expectedCell);
}

function formatCell(data) {
    if (moment(data, moment.ISO_8601, true).isValid()) {
        // Strip any time zone and convert to UTC
        return moment(data).utc().format("YYYY-MM-DD HH:mm:ss"); // Format to "YYYY-MM-DD HH:mm:ss"
    }
    return data; // Return as is if not a valid date
}
