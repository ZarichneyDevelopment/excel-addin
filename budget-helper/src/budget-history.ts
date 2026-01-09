export type BudgetHistoryEntry = {
    Expense: string;
    'Month Start': number | string;
    'Year Start': number | string;
    'Month End': number | string;
    'Year End': number | string;
    Amount: number | string;
};

type HistoryMatch = {
    entry: BudgetHistoryEntry;
    startIndex: number;
    endIndex: number;
};

function toNumber(value: unknown): number | null {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const parsed = parseFloat(String(value));
    return Number.isFinite(parsed) ? parsed : null;
}

function toMonthIndex(month: number, year: number): number | null {
    if (!Number.isFinite(month) || !Number.isFinite(year)) return null;
    if (month < 1 || month > 12) return null;
    return (year * 12) + (month - 1);
}

function getRangeIndices(row: BudgetHistoryEntry): { startIndex: number; endIndex: number } | null {
    const monthStart = toNumber(row['Month Start']);
    const yearStart = toNumber(row['Year Start']);
    const monthEnd = toNumber(row['Month End']);
    const yearEnd = toNumber(row['Year End']);
    if (monthStart === null || yearStart === null || monthEnd === null || yearEnd === null) return null;

    const startIndex = toMonthIndex(monthStart, yearStart);
    const endIndex = toMonthIndex(monthEnd, yearEnd);
    if (startIndex === null || endIndex === null) return null;

    return { startIndex, endIndex };
}

export function getBudgetHistoryMatches(
    history: BudgetHistoryEntry[],
    expense: string,
    month: number,
    year: number
): HistoryMatch[] {
    const targetIndex = toMonthIndex(month, year);
    if (targetIndex === null) return [];

    const matches: HistoryMatch[] = [];
    for (const row of history) {
        if (row['Expense'] !== expense) continue;
        const range = getRangeIndices(row);
        if (!range) continue;
        if (range.startIndex <= targetIndex && targetIndex <= range.endIndex) {
            matches.push({ entry: row, startIndex: range.startIndex, endIndex: range.endIndex });
        }
    }

    return matches;
}

export function selectBudgetHistoryEntry(
    history: BudgetHistoryEntry[],
    expense: string,
    month: number,
    year: number
): { entry: BudgetHistoryEntry | null; matches: BudgetHistoryEntry[] } {
    const matches = getBudgetHistoryMatches(history, expense, month, year);
    if (matches.length === 0) {
        return { entry: null, matches: [] };
    }

    matches.sort((a, b) => b.startIndex - a.startIndex);
    return { entry: matches[0].entry, matches: matches.map(match => match.entry) };
}

export function parseBudgetAmount(value: unknown): number {
    const parsed = toNumber(value);
    return parsed === null ? 0 : parsed;
}
