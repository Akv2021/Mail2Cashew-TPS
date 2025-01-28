// Secrets to store in GScript Properties. Don't keep actual values here.
const mySecrets = {
    EMAIL: "user1@example.com",
    SPREADSHEET_ID: "1234",
    ACCOUNT_ID_ACCOUNT_SAVINGS: ["XXXX1234", "additional identifier"],
};

// Store secrets after any update to mySecrets
storeSecrets(mySecrets);

// Regex Maps
// DefaultRegexMap contains any common fallback regex.
const DefaultRegexMap = deepFreeze({
    dateRegex: /Received : (\d{2}\/\d{2}\/\d{2}) (\d{2}:\d{2}:\d{2})/, // datePart = dd/MM/YYYY & timePart = hh:mm:ss 
    toAccountRegex: null,   // Single match i.e. 1 Capture Group OR just full match
    merchantRegex: null,    // Single match
    categoryRegex: null,    // Single match
    subcategoryRegex: null, // Single match
    notesRegex: null,       // Multiple matches supported - Join all with "-"
    isDebit: true           // Static value
});

// Email Rules
const emailSubjectToRuleIdMap = deepFreeze({
    "Subject1": "Account-Savings-Debit",
    "Subject2": "CC-Debit"
});

const emailRuleMap = deepFreeze({
    // Your ABC Bank Credit Card XX1234 has been used for a transaction of INR 363.00 on Dec 06, 2024 at 01:18:50. Info: ABC SERVICE. The Available Credit Limit
    "CC-Debit": {
        dateRegex: /\b([A-Za-z]{3}\s\d{2},\s\d{4})\s*at\s(\d{2}:\d{2}:\d{2})\b/, // '"Dec 06, 2024" at "01:18:50"' (i.e. Full Match with 2 Groups - Dec 06, 2024 & 01:18:50)
        fromAccountRegex: /Credit Card\s([A-Za-z0-9]+)\s/, // 'Credit Card "XX1234"'
        amountRegex: /transaction of\s[A-Za-z]+\s([\d,]+\.\d{2})/, // 'transaction of INR "363.00"' - Optional Comma, Required Decimal
        merchantRegex: /Info:\s*(.+?)\. The Available Credit Limit/ // 'Info: "ABC SERVICE". The Available Credit Limit'
    },
    // Rs. 200 debited from A/c no. XXXXXXXXXX1234 on 16-11-2024 and transferred to bank A/c no. XXXXXXXXXX5678 with transaction Id 12341234 - Bank Name.   
    "Account-Savings-Debit": {
        dateRegex: /\b(\d{2}-\d{2}-\d{4})\b/,               // '"16-11-2024"'
        fromAccountRegex: /A\/c no\.\s(XXXXXXXXXX\d+)\s/,   // 'A/c no. "XXXXXXXXXX1234" '
        amountRegex: /Rs\.\s([\d,]+\.\d{2})\sdebited/,      // 'Rs. "2,000.00" debited' - Optional Comma, Required Decimal
        merchantRegex: "Account Transfer",                  // Static value
        notesRegex: /bank A\/c no\.\s(XXXXXXXXXX\d+\swith\stransaction\sId\s\d+)/i, // 'bank A/c no. "XXXXXXXXXX5678 with transaction Id 12341234"'
        categoryRegex: /(transferred to bank)/i // '"transferred to bank"'
    }
});

// SMS Rules
const smsReferencePatterns = deepFreeze({
    // Rs 2500.00 debited from A/C XXXXXX1234 and credited to 1234@abc
    "Account-Savings-Debit": [
        /Rs\s+(\d+(?:\.\d{2})?)\s+debited\s+from\s+A\/C\s+XXXXXX(\d{4})\s+and\s+credited\s+to\s+([^\s]+)/
    ],
    "CC-Reversal-Credit": [
        // Reversal of Rs 6.07 credited to ABC Bank Credit Card XX1234 on 14-DEC-24.
        /Reversal of Rs \d+\.\d{2} credited to [A-Za-z ]+ [A-Za-z ]+ XX\d{4} on \d{2}-[A-Z]{3}-\d{2}\.*/,
    ],
    // If any transactions are reported both via email & sms then add pattern below ignore those messages.
    "Duplicate": []
});

const smsRuleMap = deepFreeze({
    // Rs 10.00 debited from A/C XXXXXX1234 and credited to 1234@abc
    "Account-Savings-Debit": {
        // dateRegex moved to DefaultRegexMap since it is common for all SMS.
        // dateRegex: /^Received\s*:\s*(\w{3} \d{1,2})\s+(\d{2}:\d{2}:\d{2})/, // Matches date & time separately.
        fromAccountRegex: /A\/C\s*(XXXXXX\d{4})/, // Matches "XXXXXX1234"
        amountRegex: /Rs\s*(\d+\.?\d*)/, // Matches "10.00" from "Rs 10.00"
        notesRegex: /credited to\s*([A-Za-z0-9]+@[A-Za-z0-9]+).*?/, // Matches "1234@abc"
        categoryRegex: USER_DEFAULTS.EXPENSE_CATEGORY, // Static category
        subcategoryRegex: "SubCategory1" // Static subcategory
    }, 
    // Reversal of Rs 6.07 credited to ABC Bank Credit Card XX1234 on 14-DEC-24.
    "CC-Reversal-Credit": {
        fromAccountRegex: /ABC Bank Credit Card\s+(XX\d+)/,
        amountRegex: /Reversal of Rs\s*(\d+(?:,\d+)*(?:\.\d{2})?)/,
        notesRegex: "Fuel Surcharge Reversal",
        categoryRegex: USER_DEFAULTS.INCOME_CATEGORY, // Static category
        subcategoryRegex: "SubCategory2", // Static subcategory
        isDebit: false
    }
});

// Category and Subcategory Maps
const categorySubcategoryKeywordMap = deepFreeze({
    expenses: {
        "Transit": {
            keywords: ["fuel", "travel", "commute"],
            subcategories: {
                "CNG": ["STATION 1", "Gas Station"],
                "Public Transport": ["bus", "metro", "train"]
            }
        },
        "Food": {
            subcategories: {
                "Restaurants": ["restaurant", "dining", "cafe"],
                "Groceries": ["grocery", "supermarket", "vegetables", "fruits"]
            }
        },
        "Transfer": {
            keywords: ["transferred to self-linked", "self-linked"]
        },
        "Balance Correction": {
            keywords: ["Transfer Out"],
            subcategories: null
        },
        "Uncategorized Expense": {
            keywords: [USER_DEFAULTS.EXPENSE_CATEGORY]     // Manually set in regex
        }
    },
    incomes: {
        "Income": {
            subcategories: {
                "Salary": ["salary", "paycheck", "wages"],
                "Rent": ["bonus", "incentive"]
            }
        },
        "Rewards": {
            subcategories: {
                "Cashback": ["Reversal of fuel Surcharge"],
                "Interest": ["stocks", "mutual funds", "shares"]
            }
        },
        "Balance Correction": {
            keywords: ["Transfer In"],
            subcategories: null
        }
    }
});

const subcategoryKeywordMap = deepFreeze({
    "SubCategory1": ["Sub category identifier"],
});

// Account Keyword Map
const baseAccountIdsMap = {
    "ACCOUNT_SAVINGS": ["Bank Name"], 
    "CC": ["Credit Card"],
};

const accountIdsToDisplayNameMap = {
    "ACCOUNT_SAVINGS" : "account Name in Cashew"
}

const accountKeywordMap = deepFreeze(enrichMapWithSecrets(baseAccountIdsMap, accountIdsToDisplayNameMap));

// Enrich accountKeywordMap with values from SECRETS i.e. push values of ACCOUNT_IDENTIFIERS_BOB to BOB
/*
 * @assumptions
 * 1. Secret keys are automatically converted to uppercase internally
 * 2. Base map keys are used to construct secret keys in format: ACCOUNT_ID_<ACCOUNT_KEY>
 * 3. All arrays in the returned map are deduplicated
 * 4. The function creates a new object and doesn't modify the input maps
 * 5. Account keys in baseAccountIdsMap will be replaced with values from accountIdsToAccountNameMap
 * 
 * @example
 * const baseMap = { "ACCOUNT_SAVINGS": ["Bank Name"], "CC": ["Credit Card"] }
 * const secrets = { "ACCOUNT_ID_ACCOUNT_SAVINGS": ["additional", "identifiers"] }
 * const nameMap = { "ACCOUNT_SAVINGS": "account Name in Cashew" }
 * const enriched = enrichMapWithSecrets(baseMap, secrets, nameMap);
 * // Returns: { "account Name in Cashew": ["Bank Name", "additional", "identifiers"], "CC": ["Credit Card"] }
*/

const USER_DEFAULTS = deepFreeze({
    EXPENSE_CATEGORY: "Uncategorized Expense",
    INCOME_CATEGORY: "Uncategorized Income",
    ACCOUNT: "Cash",
})