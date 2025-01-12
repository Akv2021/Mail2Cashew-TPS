//   ___ ___  _  _ ___ _____ _   _  _ _____ ___ 
// / __/ _ \| \| / __|_   _/_\ | \| |_   _/ __|
//| (_| (_) | .` \__ \ | |/ _ \| .` | | | \__ \
// \___\___/|_|\_|___/ |_/_/ \_\_|\_| |_| |___/

// Caution !! The order of files matter so Constants should be at top.
// Secrets from GScript Properties
const SECRETS = deepFreeze(fetchSecrets());

// Libraries
const moment = Moment.load();

// Data Structures
let emailData = {};
let transactionSheet;
let currentTransactionSilentErrors = [];

// Configuration Constants
const CONFIG = deepFreeze({
    get EMAIL() {
        return SECRETS.EMAIL || PropertiesService.getScriptProperties().getProperty('EMAIL');
    },
    get SPREADSHEET_ID() {
        return SECRETS.SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    },
    MAIN_SHEET_NAME: "Transactions",
    EXPECTED_SHEET_TITLE: "TestsExpected",
    RESULTS_SHEET_TITLE: "TestResults",
    DOMAIN: "budget-track.web.app",
    MOBILE_DOMAIN: "cashewapp.web.app",
    ROUTE: "addTransaction",
    EDIT_ROUTE: "addTransactionRoute",
    GMAIL_MSGID_PREFIX: "https://mail.google.com/mail/u/0/#search/rfc822msgid:",
    GMAIL_URL_PREFIX: "https://mail.google.com/mail/u/0/?source=sync&tf=1&view=pt&search=all&th=",
});

// Static Values
const CONSTANTS = deepFreeze({
    DUPLICATE: "Duplicate",
    BACKUP_SMS: "Backup SMS", // Subject of Email when offline SMS is retried. So emailDate and transaction date can have huge difference.
    TRANSACTION_SMS: "Transaction SMS",
    SUCCESS_STATUS: "Success",
    SUCCESS_SUBJECT: "Upload Transactions"
});

const Source = deepFreeze({
    SMS: "SMS",
    EMAIL: "Email"
});

const CategoryType = deepFreeze({
    EXPENSES: "expenses",
    INCOMES: "incomes"
});

const TransactionType = deepFreeze({
    DEBIT: "Debit",
    CREDIT: "Credit",
    TRANSFER: "Transfer"
});

// Labels
const LABELS = {
    EMAIL: 'Txs/💳',        // Tag All emails related to transactions from banks.
    SINGLE_SMS: 'Txs/💬',   // Tag emails which contain SMS sent using automate as soon as it is received.
    BACKUP_SMS: 'Txs/🛜',   // Tag emails which contain SMS received offline and mailed later.
    TESTCASES: 'Txs/🧪',    // Tag emails for running in sanity tests usecase.
    IGNORED: 'Txs/❌',      // Tag emails which are not valid transaction emails. Can be used on top of above filters to skip specific emails.
    PROCESSED: 'Txs/✅',    // Tag emails which have been processed by Script successfully.
};

LABELS.INCLUDE = [LABELS.EMAIL, LABELS.SINGLE_SMS, LABELS.BACKUP_SMS, LABELS.TESTCASES];
LABELS.EXCLUDE = [LABELS.PROCESSED, LABELS.IGNORED];
deepFreeze(LABELS); // Optionally, freeze the object to make it immutable

// ERROR & LOGGING
// "key" shows up in error mail & sheet.
const ErrorType = deepFreeze({
    NO_RULE: { key: "noRule", isStopping: true },
    MISSING_FIELDS: { key: "missingFields", isStopping: true },
    NO_ACCOUNT: { key: "noAccount", isStopping: false },
    NO_CATEGORY: { key: "noCategory", isStopping: false },
    REGEX_SPILLOVER: { key: "regexSpillOver", isStopping: false },
    DATE_INVALID: { key: "dateInvalid", isStopping: false },
    TIME_INVALID: { key: "timeInvalid", isStopping: false },
    DATETIME_REGEX_INVALID: { key: "dateTimeRegexInvalid", isStopping: false },
    DATE_MISMATCH: { key: "dateMismatch", isStopping: false },
});

// Processed Counts
const ProcessedCount = {
    TOTAL: 0,
    SUCCESS: 0,
    SKIPPED: 0,
    errored: {
        totalErrors: 0,
        stoppingErrors: 0,
        silentErrors: 0,
        details: Object.fromEntries(Object.keys(ErrorType).map(key => [ErrorType[key].key, 0]))
    }
};