
import { filter, forkJoin, from, map, of, reduce, switchMap, tap, toArray, throwError } from "rxjs";
import { AddToTable, TableRows$, UpdateTableRow, UpdateTableRows, WriteToTable } from "./excel-helpers";
import { getBudget, getExpenseList, getInitialAmount, getTransactions, getAllTransactions, getAllExpenseData, getAllBudgetHistory, getRollovers, setLastUpdateDate } from "./lookups";
import { selectBudgetHistoryEntry, parseBudgetAmount, BudgetHistoryEntry } from "./budget-history";
import { Transaction } from "./transaction";


export class RolloverEntry {
    Month: number;
    Year: number;
    Expense: string;
    Expenses: number;
    BOM: number;
    EOM: number;
}

/**
 * Retrieves a rollover entry for the requested month/year/expense.
 * If more than one entry exists, the first is returned while the dupes are logged.
 * When no entry exists yet, we compute totals and bootstrap a new row via `AddToTable`.
 */
export async function getRollover(month: number, year: number, expense: string): Promise<RolloverEntry> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            filter((entry: RolloverEntry) => entry.Month === month && entry.Year === year),
            filter((entry: RolloverEntry) => entry.Expense === expense),
            toArray(),
            switchMap(rows => {

                if (rows.length > 1) {
                    console.warn("Unexpected multiple rollover entries found for the same month, year, and expense", rows);
                    return of(rows[0]);
                }

                if (rows.length === 1) {
                    // Match found, continue
                    return of(rows[0]);
                }

                // No results, add in new entry instead
                // First retrieve the total expenses for the month
                const transactions$ = TableRows$('Transactions').pipe(
                    filter((transaction: Transaction) => transaction.Month === month && transaction.Year === year),
                    filter((transaction: Transaction) => transaction.Expense === expense),
                    reduce((acc, transaction) => { return acc + transaction.Amount; }, 0)
                );

                // And the expense's initial amount
                const initialAmount$ = from(getInitialAmount(expense));

                const budget$ = from(getBudget(expense, month, year));

                return forkJoin([transactions$, initialAmount$, budget$]).pipe(
                    switchMap(([monthlyExpenses, initialAmount, budget]) => {
                        const newEntry: RolloverEntry = {
                            Month: month,
                            Year: year,
                            Expense: expense,
                            Expenses: monthlyExpenses,
                            BOM: initialAmount,
                            EOM: initialAmount
                        };

                        // Convert the Promise returned by AddToTable into an Observable
                        return from(AddToTable('Rollovers', newEntry)).pipe(map(() => newEntry));
                    })
                );
            }),
            tap(entry => resolve(entry))
        ).subscribe({
            error: (err) => reject(err),
        });
    });
}

/**
 * Updates an existing rollover row by locating the matching row index and delegating
 * to `UpdateTableRow`. Rejects if no matching row exists so callers can respond immediately.
 */
export async function updateRollover(entry: RolloverEntry): Promise<void> {
    return new Promise((resolve, reject) => {
        TableRows$('Rollovers').pipe(
            toArray(), // Collect all rows into an array
            switchMap((rows) => {
                const rowIndex = rows.findIndex(row =>
                    row.Month === entry.Month &&
                    row.Year === entry.Year &&
                    row.Expense === entry.Expense
                );

                if (rowIndex === -1) {
                    console.error("No matching row found to update.");
                    return throwError(() => new Error("No matching row found."));
                }

                return from(UpdateTableRow('Rollovers', rowIndex, entry));
            })
        ).subscribe({
            next: () => resolve(),
            error: (err) => reject(err),
        });
    });
}

/**
 * Recalculates rollovers from the given starting month/year through the current date.
 * Optimized to batch data fetching and writing to improve performance.
 */
export async function resetRollover(startingMonth: number, startingYear: number, expense: string | null = null): Promise<void> {

    let expenses: string[];
    const logDetails = expense !== null;

    if (expense) {
        expenses = [expense];
    } else {
        expenses = await getExpenseList();
    }

    let today = new Date();
    let todaysMonth = today.getMonth() + 1;
    let todaysYear = today.getFullYear();

    // 1. Bulk Fetch Data
    const [allRollovers, allTransactions, expenseData, budgetHistory] = await Promise.all([
        getRollovers(),
        getAllTransactions(),
        getAllExpenseData(),
        getAllBudgetHistory()
    ]);

    // 2. Index Data for O(1) Lookup
    // Map: "Month-Year-Expense" -> { entry: RolloverEntry, index: number }
    const rolloverMap = new Map<string, { entry: RolloverEntry, index: number }>();
    allRollovers.forEach((entry, index) => {
        const key = `${entry.Month}-${entry.Year}-${entry.Expense}`;
        rolloverMap.set(key, { entry, index });
    });

    // Map: "Month-Year-Expense" -> { total, count }
    const transactionMap = new Map<string, { total: number, count: number }>();
    allTransactions.forEach(t => {
        // Ensure Transaction has Month/Year/Expense populated (assuming they come from Excel table correctly)
        // If not, we might need to parse Date. For now assuming properties exist as in original logic.
        const key = `${t.Month}-${t.Year}-${t['Expense Type']}`;
        const current = transactionMap.get(key) || { total: 0, count: 0 };
        transactionMap.set(key, {
            total: current.total + (t.Amount || 0),
            count: current.count + 1,
        });
    });

    // Map: "Expense" -> { Budget: number, Init: number }
    const expenseDataMap = new Map<string, { Budget: number, Init: number }>();
    expenseData.forEach(row => {
        expenseDataMap.set(row['Expense Type'], {
            Budget: parseFloat(row['Budget'] || 0),
            Init: parseFloat(row['Init'] || 0)
        });
    });

    // 3. Prepare Updates and New Entries
    const updates: { rowIndex: number, data: RolloverEntry }[] = [];
    const newEntries: RolloverEntry[] = [];

    // Helper to get budget from history or default
    const getBudgetInMemory = (expense: string, month: number, year: number): number => {
        const { entry, matches } = selectBudgetHistoryEntry(
            budgetHistory as BudgetHistoryEntry[],
            expense,
            month,
            year
        );

        if (matches.length > 1 && logDetails) {
            console.warn(`Multiple BudgetHistory matches for ${expense} ${month}/${year}:`, matches);
        }

        if (entry) {
            return parseBudgetAmount(entry.Amount);
        }

        // Default
        return expenseDataMap.get(expense)?.Budget || 0;
    };

    const getInitialAmountInMemory = (expense: string): number => {
        return expenseDataMap.get(expense)?.Init || 0;
    };

    for (const expense of expenses) {
        let loopLimit = 24;
        let month = startingMonth;
        let year = startingYear;

        if (year > todaysYear || (year === todaysYear && month > todaysMonth)) {
            console.error("Cannot reset rollover for a future date.");
            return;
        }

        while ((year < todaysYear || (year === todaysYear && month <= todaysMonth)) && loopLimit > 0) {
            
            const key = `${month}-${year}-${expense}`;
            let currentRolloverData = rolloverMap.get(key);
            let rolloverEntry = currentRolloverData ? { ...currentRolloverData.entry } : null;

            // Get Budget
            const budget = getBudgetInMemory(expense, month, year);

            // Get Previous Rollover EOM
            let previousMonth = month - 1;
            let previousYear = year;
            if (previousMonth === 0) {
                previousMonth = 12;
                previousYear--;
            }
            const prevKey = `${previousMonth}-${previousYear}-${expense}`;
            const prevRolloverData = rolloverMap.get(prevKey);
            
            let prevEOM = 0;
            if (prevRolloverData) {
                prevEOM = prevRolloverData.entry.EOM;
            } else {
                // If previous doesn't exist, we assume it's the start, so we use Initial Amount
                // Matches logic: "And the expense's initial amount" from original getRollover fallback
                prevEOM = getInitialAmountInMemory(expense);
            }

            // Get Monthly Transactions
            const transactionEntry = transactionMap.get(`${month}-${year}-${expense}`);
            const totalAmount = transactionEntry?.total || 0;

            // Calculate
            if (!rolloverEntry) {
                // Create new
                rolloverEntry = {
                    Month: month,
                    Year: year,
                    Expense: expense,
                    Expenses: 0,
                    BOM: 0,
                    EOM: 0
                };
            }

            rolloverEntry.Expenses = totalAmount;
            rolloverEntry.BOM = prevEOM;
            rolloverEntry.EOM = rolloverEntry.BOM + budget + totalAmount;

            if (logDetails) {
                const transactionCount = transactionEntry?.count || 0;
                const prevSource = prevRolloverData ? `${previousMonth}/${previousYear}` : 'init';
                console.log(
                    `Rollover ${expense} ${month}/${year}: BOM ${prevEOM}, Budget ${budget}, ` +
                    `Spent ${totalAmount} (${transactionCount} tx), EOM ${rolloverEntry.EOM} (prev: ${prevSource})`
                );
            }

            // Store result
            if (currentRolloverData) {
                // It's an update
                // Check if we already have an update pending for this index?
                // Actually, just push to updates. If we process sequentially, we might update the same row?
                // No, we iterate time sequentially. 
                // BUT: rolloverMap.get(key).entry needs to be updated IN MEMORY so next iteration picks up new EOM!
                currentRolloverData.entry = rolloverEntry; // Update map reference
                updates.push({ rowIndex: currentRolloverData.index, data: rolloverEntry });
            } else {
                // It's a new entry
                newEntries.push(rolloverEntry);
                // Add to map so next iteration finds it
                // Note: Index is unknown for new entries until saved. 
                // But for calculation purposes (prevRollover), we only need the entry data (EOM).
                rolloverMap.set(key, { entry: rolloverEntry, index: -1 }); 
            }

            month++;
            if (month === 13) {
                month = 1;
                year++;
            }
            loopLimit--;
        }
    }

    // 4. Batch Write
    if (updates.length > 0) {
        await UpdateTableRows('Rollovers', updates);
    }

    if (newEntries.length > 0) {
        // AddToTable takes one object, WriteToTable takes array of arrays (rows)
        // We need to convert objects to arrays of values, ensuring order matches table columns?
        // AddToTable implementation: var row = Object.values(data); WriteToTable(tableName, [row]);
        // Caution: Object.values order is not guaranteed to match Excel column order if interface properties aren't ordered.
        // Ideally we should map to columns. 
        // Based on RolloverEntry class: Month, Year, Expense, Expenses, BOM, EOM.
        // Let's assume table columns match this order.
        const rows = newEntries.map(e => [e.Month, e.Year, e.Expense, e.Expenses, e.BOM, e.EOM]);
        await WriteToTable('Rollovers', rows);
    }

    await setLastUpdateDate(new Date());
}
