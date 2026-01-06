
import Papa from 'papaparse';
import { getAccounts, getExpenseList, getMatchingRules, getAllTransactionIds, MatchSet, getAmbiguousItems } from './lookups';
import { AddToTable } from './excel-helpers';

export class Transaction {
    'id': string | null;
    'Account Type': string; // useless rbc field
    'Account Number': string;
    'Account Name': string;
    'Transaction Date': Date;
    'Cheque Number': string | null; // useless rbc field
    'Description 1': string;
    'Description 2': string;
    'CAD$': number | null;
    'USD$': number | null;

    Month: number;
    Year: number;
    Account: string;
    Date: string;
    Description: string;
    Amount: number;
    Expense: string;
    Memo: string;

    static fromCsv(csvString: string): Transaction[] {
        const result = Papa.parse(csvString, { header: true });
        const transactions = result.data as Transaction[];

        transactions.forEach(async transaction => {
            // remove transaction if it has no description
            if (!transaction['Description 1']) {
                transactions.splice(transactions.indexOf(transaction), 1);
                return;
            }

            transaction['id'] = await createTransactionId(transaction);

            transaction['Transaction Date'] = new Date(transaction['Transaction Date']);
            transaction.Month = transaction['Transaction Date'].getMonth() + 1;
            transaction.Year = transaction['Transaction Date'].getFullYear();
            transaction.Date = transaction['Transaction Date'].toLocaleDateString();

            transaction['CAD$'] = transaction['CAD$'] ? parseFloat(transaction['CAD$'].toString()) : null;
            transaction['USD$'] = transaction['USD$'] ? parseFloat(transaction['USD$'].toString()) : null;
            transaction.Amount = transaction['USD$'] || transaction['CAD$'] || 0;

            transaction.Description = transaction['Description 1'] + ' ' + transaction['Description 2'];
        });

        return transactions;
    }
}

function encodeTransaction(transaction: Transaction) {
    return new TextEncoder().encode(`
    ${transaction['Transaction Date']}
    ${transaction['Account Number']}
    ${transaction['Description 1']}
    ${transaction['Description 2']}
    ${transaction['CAD$']}
    ${transaction['USD$']}
  `);
}

async function createTransactionId(transaction: Transaction) {
    const buffer = encodeTransaction(transaction);
    const hashArrayBuffer = await crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashArrayBuffer));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex;
}

export async function ProcessTransactions(fileContent) {
    // console.log('File content:', fileContent);

    var fileTransactions = Transaction.fromCsv(fileContent);
    console.log('Incoming Transactions:', fileTransactions);

    const expenses = await getExpenseList();
    console.log('Expenses:', expenses);

    const expenseMatches = await getMatchingRules();
    console.log('Matching Rules:', expenseMatches);

    const ambiguousItems = await getAmbiguousItems();
    console.log('Ambiguous Items:', ambiguousItems);

    const accountNames = await getAccounts();
    console.log('accountNames:', accountNames);

    const existingTransactions = await getAllTransactionIds();
    // console.log('existingTransactions:', existingTransactions);

    const newTransactions: Transaction[] = [];

    // Strategy for efficiency: first attempt an exact match, then iterate over wildcard matches
    const indexedLookup = new Map();
    const wildcardLookup = new Array();

    // Iterate over matching ruleset to build these two look up methods
    expenseMatches.forEach(lookupRule => {

        const key = `${lookupRule['Match 1']}${lookupRule['Match 2']}`;
        // future todo: add support on 'amount' to scenario of same vendor can be split into multiple expense categories via charge amount
        indexedLookup.set(key, lookupRule['Expense Type']);

        if (!lookupRule['Match 2']) {
            // This look up rule isnt a candidate for an exact match, so it's assumed the first part is a wildcard matches
            wildcardLookup.push({ match: lookupRule['Match 1'], expense: lookupRule['Expense Type'] });
        }
    });

    // Iterate over rows in files to fully build model and apply rules
    fileTransactions.forEach(transaction => {

        // Duplication filter
        if (transaction['id'] && existingTransactions.includes(transaction['id'])) {
            // console.log('Transaction already exists for', transaction);
            return;
        }

        newTransactions.push(transaction);

        // Resolve account name from account number
        transaction['Account Name'] = accountNames[transaction['Account Number']];

        // Provide friendly account name with last four digits of account number
        transaction.Account = `${transaction['Account Name']} (${transaction['Account Number'].slice(-4)})`;

        // Check for Ambiguity
        const isAmbiguous = ambiguousItems.some(item => 
            transaction['Description 1'].toLowerCase().includes(item.Item.toLowerCase())
        );

        if (isAmbiguous) {
            console.warn('Transaction flagged as ambiguous:', transaction['Description 1']);
            transaction['Expense'] = 'Ambiguous';
            return;
        }

        // Auto expense type categorization:

        // Attempt to find an exact rule match to identify type of expense
        let exactMatch = indexedLookup.get(`${transaction['Description 1']}${transaction['Description 2']}`);
        if (exactMatch) {
            transaction['Expense'] = exactMatch;
            return;
        }

        // Iterate over rules to find match by substring (Aggressive)
        for (let rule of wildcardLookup) {
            if (transaction['Description 1'].toLowerCase().includes(rule.match.toLowerCase())) {
                transaction['Expense'] = rule.expense;
                return;
            }
        }

        console.warn('No match found for transaction:', transaction['Description 1'], transaction['Description 2']);
        
        // Removed auto-add to MatchingRules to prevent pollution. 
        // User should manually review and add rules or we can implement a "Learning" feature later.
    });

    console.log('New Transactions:', newTransactions);

    // Transform object to match the Excel table structure
    return newTransactions.map(transaction => [
        transaction.id,
        transaction.Month,
        transaction.Year,
        transaction.Date,
        transaction.Account,
        transaction.Expense,
        transaction.Amount,
        transaction.Description,
        transaction.Memo
    ]);
}