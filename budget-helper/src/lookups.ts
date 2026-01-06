import { filter, map, of, reduce, switchMap, tap, toArray } from 'rxjs';
import { NamedRangeValues$, SetNamedRangeValue, EnsureTableExists, TableRows$ } from './excel-helpers';
import { RolloverEntry } from './rollover';
import { Transaction } from './transaction';
import * as rollover from './rollover';

export async function initializeSchema() {
    try {
        await EnsureTableExists('AmbiguousItems', ['Item', 'IsAmbiguous', 'OverrideCount', 'Confidence']);
        // Add other schema checks here if needed
    } catch (error) {
        console.error("Schema initialization failed:", error);
    }
}

export async function getLastUpdateDate(): Promise<Date | null> {
    return new Promise((resolve, reject) => {
        NamedRangeValues$('LastRolloverUpdate').pipe(
            toArray(),
            tap(values => {
                if (values.length > 0 && values[0]) {
                    resolve(new Date(values[0]));
                } else {
                    resolve(null);
                }
            })
        ).subscribe({
            error(err) { 
                // If range doesn't exist or error, resolve null to be safe
                console.warn('Could not fetch LastRolloverUpdate:', err);
                resolve(null); 
            },
        });
    });
}

export async function setLastUpdateDate(date: Date): Promise<void> {
    const dateString = date.toISOString().split('T')[0]; // YYYY-MM-DD
    return SetNamedRangeValue('LastRolloverUpdate', dateString);
}

export async function getAccounts(): Promise<{ [key: string]: string }> {
    // returns collection of key value pair
    return new Promise((resolve, reject) => {
        TableRows$('Accounts').pipe(
            map(row => {
                var obj = {};
                obj[row['Number']] = row['Name'];
                return obj;
            }),
            reduce((acc, obj) => ({ ...acc, ...obj }), {})
        ).subscribe({
            next: (obj) => resolve(obj),
            error: (err) => reject(err),
        });
    });
}

export class MatchSet {
    'Match 1': string;
    'Match 2': string;
    'Amount': string;
    'Expense Type': string;
}

export async function getMatchingRules(): Promise<MatchSet[]> {
    return new Promise((resolve, reject) => {
        TableRows$('MatchingRules').pipe(
            toArray()
        ).subscribe({
            next: (rows) => resolve(rows),
            error: (err) => reject(err),
        });
    });
}

export async function getExpenseList(): Promise<string[]> {
    return new Promise((resolve, reject) => {
        NamedRangeValues$('Expenses').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getTransactions(month: number, year: number, expense: string): Promise<Transaction[]> {
    return new Promise((resolve, reject) => {
        TableRows$('Transactions').pipe(
            // tap((transaction: Transaction) => console.log(transaction)),
            filter((transaction: Transaction) => transaction.Month === month && transaction.Year === year),
            filter((transaction: Transaction) => transaction['Expense Type'] === expense),
            toArray(),
            tap(transactions => resolve(transactions))
        ).subscribe({
            error: (err) => reject(err),
        });
    });

}

export async function getAllTransactions(): Promise<Transaction[]> {
    return new Promise((resolve, reject) => {
        TableRows$('Transactions').pipe(
            toArray(),
            tap(transactions => resolve(transactions))
        ).subscribe({
            error: (err) => reject(err),
        });
    });
}

export async function getAllTransactionIds(): Promise<string[]> {
    return new Promise((resolve, reject) => {
        NamedRangeValues$('TransactionIds').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getRollovers(): Promise<RolloverEntry[]> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            toArray()
        ).subscribe({
            next: (rows) => resolve(rows),
            error: (err) => reject(err),
        });
    });
}

export class AmbiguousItem {
    Item: string;
    IsAmbiguous: string;
    OverrideCount: number;
    Confidence: number;
}

export async function getAmbiguousItems(): Promise<AmbiguousItem[]> {
    return new Promise((resolve, reject) => {
        TableRows$('AmbiguousItems').pipe(
            toArray()
        ).subscribe({
            next: (rows) => resolve(rows),
            error: (err) => reject(err),
        });
    });
}

export async function getInitialAmount(expense: string): Promise<number> {
    return new Promise((resolve, reject) => {
        TableRows$('ExpenseData').pipe(
            filter(row => row['Expense Type'] === expense),
            map(row => row['Init']),
            toArray(),
            tap(values => resolve(parseFloat(values[0])))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getAllExpenseData(): Promise<any[]> {
    return new Promise((resolve, reject) => {
        TableRows$('ExpenseData').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getAllBudgetHistory(): Promise<any[]> {
    return new Promise((resolve, reject) => {
        TableRows$('BudgetHistory').pipe(
            toArray(),
            tap(values => resolve(values))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getBudget(expense: string, month: number | null = null, year: number | null = null): Promise<number> {
    return new Promise((resolve, reject) => {

        const currentBudget$ = TableRows$('ExpenseData').pipe(
            filter(row => row['Expense Type'] === expense),
            map(row => row.Budget)
        );

        let dataSource$ = currentBudget$;

        if (month && year) {
            // Request for specific month and year, look first into change history
            dataSource$ = TableRows$('BudgetHistory').pipe(
                filter(row => row['Expense'] === expense),
                filter(row => (row['Month Start'] <= month && month <= row['Month End'])
                    && (row['Year Start'] <= year && year <= row['Year End'])
                ),
                toArray(),
                switchMap(rows => {

                    if (rows.length > 1) {
                        console.warn("Unexpected multiple rollover entries found for the same month, year, and expense", rows);
                        debugger;
                    } else if (rows.length === 1) {
                        // Match found, continue
                        return of(rows[0].Amount);
                    }

                    // If no historical entry are found, resume fetching from current budget
                    return currentBudget$;
                })
            );
        }

        dataSource$.pipe(
            toArray(),
            tap(values => resolve(parseFloat(values[0])))
        ).subscribe({
            error(err) { reject(err); },
        });
    });
}

export async function getRollover(month: number, year: number, expense: string): Promise<RolloverEntry> {
    return rollover.getRollover(month, year, expense);
}