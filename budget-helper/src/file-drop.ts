import { WriteToTable } from "./excel-helpers";
import { ProcessTransactions } from "./transaction";
import { handleError } from "./error-handler";
import { logToConsole } from "./logger";

function getSuggestedStartMonthYear(rows: any[]): { month: number, year: number } | null {
    // Expected rows: [id, Month, Year, Date, Account, Expense, Amount, Description, Memo]
    let minYear = Number.POSITIVE_INFINITY;
    let minMonth = Number.POSITIVE_INFINITY;

    for (const row of rows) {
        const month = typeof row?.[1] === 'number' ? row[1] : parseInt(row?.[1], 10);
        const year = typeof row?.[2] === 'number' ? row[2] : parseInt(row?.[2], 10);
        if (!Number.isFinite(month) || !Number.isFinite(year)) continue;
        if (year < minYear || (year === minYear && month < minMonth)) {
            minYear = year;
            minMonth = month;
        }
    }

    if (!Number.isFinite(minYear) || !Number.isFinite(minMonth)) return null;
    if (minMonth < 1 || minMonth > 12) return null;

    return { month: minMonth, year: minYear };
}

function maybeApplySuggestedStartMonthYear(suggested: { month: number, year: number }) {
    if (typeof document === 'undefined') return;

    const monthInput = document.getElementById('month-input') as HTMLInputElement | null;
    const yearInput = document.getElementById('year-input') as HTMLInputElement | null;

    const userEdited = Boolean(monthInput?.dataset?.userEdited || yearInput?.dataset?.userEdited);
    if (userEdited) return;

    if (monthInput) monthInput.value = suggested.month.toString();
    if (yearInput) yearInput.value = suggested.year.toString();

    document.dispatchEvent(new Event('budgethelper:inputs-changed'));
}

export function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

export function handleDragOver(event) {
    event.stopPropagation();
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy'; // Explicitly show this is a copy.
}

export function handleFileDrop(event) {

    var files = event.dataTransfer.files;
    if (files.length > 0) {
        var file = files[0];
        logToConsole(`Reading file: ${file.name}...`, 'info');
        var reader = new FileReader();

        reader.onload = ProcessFileDrop;

        reader.readAsText(file);
    }
}

export async function ProcessFileDrop(event) {
    try {
        var contents = event.target.result;
        logToConsole('Processing transactions...', 'info');

        const transactions = await ProcessTransactions(contents);

        await WriteToTable("Transactions", transactions);
        
        logToConsole(`Successfully imported ${transactions.length} transactions.`, 'success');

        const suggested = getSuggestedStartMonthYear(transactions);
        if (suggested) {
            logToConsole(`Suggested recalc start: ${suggested.month}/${suggested.year} (earliest imported month).`, 'info');
            maybeApplySuggestedStartMonthYear(suggested);
        } else {
            logToConsole('Suggested recalc start: (could not infer from import)', 'warn');
        }
    } catch (error) {
        handleError(error, 'ProcessFileDrop');
        logToConsole('Failed to process file.', 'error');
    }
}
