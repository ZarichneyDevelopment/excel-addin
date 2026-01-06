/** @jest-environment jsdom */

import { TextEncoder } from 'util';
global.TextEncoder = TextEncoder;

const mockCrypto = {
    subtle: {
        digest: jest.fn().mockResolvedValue(new Uint8Array([1, 2, 3, 4]).buffer)
    },
    getRandomValues: jest.fn((buffer) => new Uint8Array(buffer.length))
};
Object.defineProperty(global, 'crypto', {
    value: mockCrypto,
    writable: true
});

// Mock DOM
document.getElementById = jest.fn((id) => {
    if (id === 'month-input') return { value: '1' };
    if (id === 'year-input') return { value: '2023' };
    if (id === 'expense-dropdown') return { options: [{ text: 'All Expenses' }], selectedIndex: 0, add: jest.fn(), innerHTML: '' };
    if (id === 'drop-area') return { addEventListener: jest.fn(), classList: { add: jest.fn(), remove: jest.fn() } };
    if (id === 'reset') return { onclick: null };
    if (id === 'error-close-btn' || id === 'error-copy-btn' || id === 'clear-console') return { addEventListener: jest.fn() };
    if (id === 'last-updated') return { textContent: '' };
    if (id === 'console-output') return { appendChild: jest.fn(), scrollTop: 0, scrollHeight: 0, innerHTML: '' };
    return { value: '' }; // Fallback
}) as any;

// Mock Office Global BEFORE imports
global.Office = {
    onReady: jest.fn(async (callback) => {
        await callback({ host: 'Excel' });
    }),
    HostType: { Excel: 'Excel' }
} as any;

import { handleFileDrop } from './file-drop';
import { TriggerResetRollovers } from './taskpane/taskpane';
import { TableRows$, AddToTable, UpdateTableRow, UpdateTableRows, WriteToTable, NamedRangeValues$ } from './excel-helpers';
import * as lookups from './lookups';

// Mock Excel Global
const context = {
    sync: jest.fn(),
    workbook: {
        getSelectedRange: jest.fn(() => ({ values: [] })),
        tables: {
            getItem: jest.fn(),
            getItemOrNullObject: jest.fn()
        }
    }
};
global.Excel = {
    run: jest.fn(async (callback) => {
        await callback(context);
    }),
} as any;

// Mock modules but keep logic where possible
jest.mock('./excel-helpers', () => ({
    TableRows$: jest.fn(),
    AddToTable: jest.fn(),
    WriteToTable: jest.fn(),
    UpdateTableRow: jest.fn(),
    UpdateTableRows: jest.fn(),
    NamedRangeValues$: jest.fn(),
    EnsureTableExists: jest.fn(),
    SetNamedRangeValue: jest.fn()
}));

// In-Memory "Database"
let db = {
    Expenses: ['Dining', 'Groceries'],
    MatchingRules: [{ 'Match 1': 'Starbucks', 'Match 2': '', 'Expense Type': 'Dining' }],
    AmbiguousItems: [],
    Accounts: [{ Number: '1234', Name: 'Chequing' }],
    Transactions: [],
    Rollovers: [],
    ExpenseData: [
        { 'Expense Type': 'Dining', 'Budget': 500, 'Init': 0 },
        { 'Expense Type': 'Groceries', 'Budget': 800, 'Init': 0 }
    ],
    BudgetHistory: [],
    TransactionIds: []
};

// Setup Excel Helper Mocks to read/write from DB
import { of } from 'rxjs';

(TableRows$ as jest.Mock).mockImplementation((tableName) => {
    if (db[tableName]) {
        return of(...db[tableName]);
    }
    return of();
});

(NamedRangeValues$ as jest.Mock).mockImplementation((rangeName) => {
    if (rangeName === 'Expenses') {
        return of(...db.Expenses);
    }
    if (rangeName === 'TransactionIds') {
        return of(...db.TransactionIds);
    }
    return of();
});

// We need to capture writes to verify the flow
(AddToTable as jest.Mock).mockImplementation(async (tableName, data) => {
    db[tableName].push(data);
    return Promise.resolve();
});

(WriteToTable as jest.Mock).mockImplementation(async (tableName, rows) => {
    // Convert arrays back to objects for DB storage to support subsequent reads/updates
    if (tableName === 'Rollovers') {
        const objects = rows.map(row => ({
            Month: row[0],
            Year: row[1],
            Expense: row[2],
            Expenses: row[3],
            BOM: row[4],
            EOM: row[5]
        }));
        db.Rollovers.push(...objects);
    } else if (tableName === 'Transactions') {
        // [id, Month, Year, Date, Account, Expense, Amount, Description, Memo]
        const objects = rows.map(row => ({
            id: row[0],
            Month: row[1],
            Year: row[2],
            Date: row[3],
            Account: row[4],
            'Expense Type': row[5],
            Amount: row[6],
            Description: row[7],
            Memo: row[8]
        }));
        db.Transactions.push(...objects);
    } else {
        // Fallback
        db[tableName].push(...rows);
    }
    return Promise.resolve();
});

(UpdateTableRows as jest.Mock).mockImplementation(async (tableName, updates) => {
    updates.forEach(u => {
       if (db[tableName][u.rowIndex]) {
           db[tableName][u.rowIndex] = u.data;
       }
    });
    return Promise.resolve();
});

describe('Integration Test: Full Workflow', () => {
    beforeEach(() => {
        // Reset DB
        db.Transactions = [];
        db.Rollovers = [];
        // ... reset other tables if needed
        jest.clearAllMocks();
        
        // Ensure DOM mocks are clean/ready
        (document.getElementById as jest.Mock).mockImplementation((id) => {
            if (id === 'month-input') return { value: '1' };
            if (id === 'year-input') return { value: '2023' };
            if (id === 'expense-dropdown') return { options: [{ text: 'All Expenses' }], selectedIndex: 0, add: jest.fn(), innerHTML: '' };
            if (id === 'drop-area') return { addEventListener: jest.fn(), classList: { add: jest.fn(), remove: jest.fn() } };
            if (id === 'reset') return { onclick: null };
            if (id === 'error-close-btn' || id === 'error-copy-btn' || id === 'clear-console') return { addEventListener: jest.fn() };
            if (id === 'last-updated') return { textContent: '' };
            if (id === 'console-output') return { appendChild: jest.fn(), scrollTop: 0, scrollHeight: 0, innerHTML: '' };
            if (id === 'error-container') return { style: { display: 'none' } };
            if (id === 'error-content' || id === 'error-details') return { textContent: '' };
            return { value: '' }; // Fallback
        });
    });

    it('processes a file upload, categorizes transactions, and updates rollovers', async () => {
        // 1. Prepare "File Drop"
        const csvContent = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Starbucks,Coffee,1234,2023-01-15,5.50,
Walmart,Supercenter,1234,2023-01-16,100.00,
`;
        
        // Mock FileReader
        const mockReader = {
            readAsText: jest.fn(),
            result: csvContent,
            onload: null as any,
        };
        (global as any).FileReader = jest.fn(() => mockReader);

        // Simulate Drop
        const mockFile = new Blob([csvContent], { type: 'text/csv' });
        const mockEvent = {
            stopPropagation: jest.fn(),
            preventDefault: jest.fn(),
            dataTransfer: { files: [mockFile] }
        };

        // Trigger File Processing
        handleFileDrop(mockEvent as any);
        
        // Manually trigger onload since the mock doesn't do it
        mockReader.onload({ target: { result: csvContent } } as any);

        // Wait for async processing
        await new Promise(resolve => setTimeout(resolve, 100)); // Tick

        // ASSERT: Transactions should be in DB
        expect(db.Transactions.length).toBe(2);
        const starbucksTx = db.Transactions.find(t => t.Description.includes('Starbucks'));
        expect(starbucksTx).toBeDefined();
        expect(starbucksTx['Expense Type']).toBe('Dining'); // Auto-categorized

        // 2. Trigger Rollover Reset
        // We set input values in the mock above (Month 1, Year 2023)
        await TriggerResetRollovers();

        // ASSERT: Rollovers should be updated
        
        const diningRollover = db.Rollovers.find(r => r.Expense === 'Dining');
        expect(diningRollover).toBeDefined();
        expect(diningRollover.Month).toBe(1);
        expect(diningRollover.Year).toBe(2023);
        // Ensure values are numbers
        expect(Number(diningRollover.Expenses)).toBe(5.5);
        expect(Number(diningRollover.EOM)).toBe(505.5); // 0 + 500 + 5.5
    });
});