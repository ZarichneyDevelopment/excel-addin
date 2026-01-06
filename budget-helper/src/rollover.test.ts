import { RolloverEntry } from './rollover';
import { TableRows$, AddToTable, UpdateTableRow, UpdateTableRows, WriteToTable } from './excel-helpers';
import { getBudget, getExpenseList, getInitialAmount, getTransactions, getAllTransactions, getAllExpenseData, getAllBudgetHistory, getRollovers } from './lookups';
import { of, throwError } from 'rxjs';

// Get actual implementations for non-mocked exports
const actualRollover = jest.requireActual('./rollover');

// Mock the rollover module itself to replace getRollover, updateRollover, resetRollover with Jest mock functions
jest.mock('./rollover', () => ({
  // Keep original exports like RolloverEntry and actual resetRollover for testing its internal logic
  ...actualRollover,
  getRollover: jest.fn(),
  updateRollover: jest.fn(),
  // resetRollover will call the actual implementation which then calls the mocked getRollover/updateRollover
  resetRollover: jest.fn((month, year, expense) => actualRollover.resetRollover(month, year, expense)),
}));

// Now, import the functions which will be the mocked versions
import { getRollover, updateRollover, resetRollover } from './rollover';

// Mock dependencies for excel-helpers and lookups (these are external dependencies to rollover.ts)
jest.mock('./excel-helpers', () => ({
  TableRows$: jest.fn(),
  AddToTable: jest.fn(() => Promise.resolve(null)),
  UpdateTableRow: jest.fn(() => Promise.resolve(null)),
  UpdateTableRows: jest.fn(() => Promise.resolve(null)),
  WriteToTable: jest.fn(() => Promise.resolve(null)),
}));

jest.mock('./lookups', () => ({
  getBudget: jest.fn(() => Promise.resolve(0)),
  getInitialAmount: jest.fn(() => Promise.resolve(0)),
  getTransactions: jest.fn(() => Promise.resolve([])),
  getExpenseList: jest.fn(() => Promise.resolve(['Expense1', 'Expense2'])),
  getAllTransactions: jest.fn(() => Promise.resolve([])),
  getAllExpenseData: jest.fn(() => Promise.resolve([])),
  getAllBudgetHistory: jest.fn(() => Promise.resolve([])),
  getRollovers: jest.fn(() => Promise.resolve([])),
  setLastUpdateDate: jest.fn(() => Promise.resolve()),
}));

const mockedLookups = {
    getBudget: getBudget as jest.Mock,
    getInitialAmount: getInitialAmount as jest.Mock,
    getTransactions: getTransactions as jest.Mock,
    getExpenseList: getExpenseList as jest.Mock,
    getAllTransactions: getAllTransactions as jest.Mock,
    getAllExpenseData: getAllExpenseData as jest.Mock,
    getAllBudgetHistory: getAllBudgetHistory as jest.Mock,
    getRollovers: getRollovers as jest.Mock
};

describe('Rollover Module', () => {
  beforeEach(() => {
    // Reset mocks before each test
    // Clear mocks for functions from rollover.ts
    (getRollover as jest.Mock).mockClear();
    (updateRollover as jest.Mock).mockClear();
    (resetRollover as jest.Mock).mockClear();

    // Clear mocks for external dependencies
    (TableRows$ as jest.Mock).mockClear();
    (AddToTable as jest.Mock).mockClear();
    (UpdateTableRow as jest.Mock).mockClear();
    (UpdateTableRows as jest.Mock).mockClear();
    (WriteToTable as jest.Mock).mockClear();
    
    mockedLookups.getBudget.mockClear();
    mockedLookups.getInitialAmount.mockClear();
    mockedLookups.getTransactions.mockClear();
    mockedLookups.getExpenseList.mockClear();
    mockedLookups.getAllTransactions.mockClear();
    mockedLookups.getAllExpenseData.mockClear();
    mockedLookups.getAllBudgetHistory.mockClear();
    mockedLookups.getRollovers.mockClear();
  });

  describe('RolloverEntry', () => {
    it('should be able to create an instance', () => {
      const entry: RolloverEntry = {
        Month: 1,
        Year: 2023,
        Expense: 'Groceries',
        Expenses: 100,
        BOM: 50,
        EOM: 150,
      };
      expect(entry).toBeDefined();
      expect(entry.Expense).toBe('Groceries');
    });
  });

  describe('getRollover', () => {
    const mockRolloverEntry: RolloverEntry = {
      Month: 1,
      Year: 2023,
      Expense: 'Rent',
      Expenses: 0,
      BOM: 0,
      EOM: 0,
    };

    it('should return an existing rollover entry', async () => {
      // Mock the internal getRollover behavior for this specific test
      (TableRows$ as jest.Mock).mockReturnValue(of(mockRolloverEntry));

      const result = await actualRollover.getRollover(1, 2023, 'Rent'); // Call actual getRollover for testing
      expect(result).toEqual(mockRolloverEntry);
      expect(TableRows$).toHaveBeenCalledWith('Rollovers');
    });

    it('should calculate and add a new entry if none exists', async () => {
      // Mock TableRows$ to return empty for 'Rollovers' so the 'add new entry' path is taken
      (TableRows$ as jest.Mock).mockImplementation((tableName: string) => {
        if (tableName === 'Rollovers') {
          return of(); // Empty stream
        } else if (tableName === 'Transactions') {
          // Simulate no transactions for the given month/year/expense initially
          return of(); // Empty stream, so reduce will return 0
        }
        return of(); // Default to empty for other calls
      });

      // These mocks are now for the lookups functions themselves, not TableRows$
      mockedLookups.getTransactions.mockResolvedValueOnce([]); 
      mockedLookups.getInitialAmount.mockResolvedValueOnce(0); 
      mockedLookups.getBudget.mockResolvedValueOnce(0); 

      const newEntry = { Month: 1, Year: 2023, Expense: 'New Expense', Expenses: 0, BOM: 0, EOM: 0 };
      (AddToTable as jest.Mock).mockResolvedValueOnce(null); 

      const result = await actualRollover.getRollover(1, 2023, 'New Expense'); // Call actual getRollover for testing

      expect(TableRows$).toHaveBeenCalledWith('Rollovers'); // Called to check for existing entry
      expect(TableRows$).toHaveBeenCalledWith('Transactions'); // Called to get transactions for new entry calculation

      // The lookups functions are called within the Observable pipeline, so they are not directly called here
      // Instead, we assert that the final AddToTable is called with computed values based on the mocks
      expect(AddToTable).toHaveBeenCalledWith('Rollovers', expect.objectContaining(newEntry));
      expect(result).toEqual(expect.objectContaining(newEntry));
    });

    it('should warn and return the first entry if multiple exist', async () => {
      const entry1 = { ...mockRolloverEntry, EOM: 100 };
      const entry2 = { ...mockRolloverEntry, EOM: 200 };
      (TableRows$ as jest.Mock).mockReturnValue(of(entry1, entry2)); // Two existing entries

      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

      const result = await actualRollover.getRollover(1, 2023, 'Rent'); // Call actual getRollover for testing

      expect(consoleWarnSpy).toHaveBeenCalledTimes(1);
      expect(consoleWarnSpy).toHaveBeenCalledWith(
        "Unexpected multiple rollover entries found for the same month, year, and expense",
        [entry1, entry2]
      );
      expect(result).toEqual(entry1); // Should return the first one

      consoleWarnSpy.mockRestore();
    });

    it('should propagate errors from Excel read operations', async () => {
      const errorMessage = 'Excel read error';
      (TableRows$ as jest.Mock).mockReturnValue(throwError(() => new Error(errorMessage)));

      await expect(actualRollover.getRollover(1, 2023, 'Rent')).rejects.toThrow(errorMessage); // Call actual getRollover for testing
    });
  });

  describe('updateRollover', () => {
    const existingRolloverEntries = [
      { Month: 1, Year: 2023, Expense: 'Rent', Expenses: 500, BOM: 1000, EOM: 600 },
      { Month: 2, Year: 2023, Expense: 'Food', Expenses: 200, BOM: 300, EOM: 100 },
    ];

    it('should update an existing rollover entry', async () => {
      (TableRows$ as jest.Mock).mockReturnValue(of(...existingRolloverEntries));
      (UpdateTableRow as jest.Mock).mockResolvedValueOnce(null);

      const updatedEntry = { ...existingRolloverEntries[0], EOM: 700 };
      await actualRollover.updateRollover(updatedEntry); // Call actual updateRollover for testing

      expect(TableRows$).toHaveBeenCalledWith('Rollovers');
      expect(UpdateTableRow).toHaveBeenCalledWith('Rollovers', 0, updatedEntry);
    });

    it('should reject if no matching entry is found', async () => {
      (TableRows$ as jest.Mock).mockReturnValue(of(...existingRolloverEntries));
      const nonExistentEntry = { Month: 3, Year: 2023, Expense: 'Utilities', Expenses: 0, BOM: 0, EOM: 0 };

      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});

      await expect(actualRollover.updateRollover(nonExistentEntry)).rejects.toThrow("No matching row found."); // Call actual updateRollover for testing

      expect(TableRows$).toHaveBeenCalledWith('Rollovers');
      expect(UpdateTableRow).not.toHaveBeenCalled();
      expect(consoleErrorSpy).toHaveBeenCalledWith("No matching row found to update.");

      consoleErrorSpy.mockRestore();
    });

    it('should propagate errors from Excel write operations', async () => {
      (TableRows$ as jest.Mock).mockReturnValue(of(...existingRolloverEntries));
      const errorMessage = 'Excel write error';
      (UpdateTableRow as jest.Mock).mockRejectedValueOnce(new Error(errorMessage));
      // Define updatedEntry here since it's used in this specific test
      const updatedEntry = { ...existingRolloverEntries[0], EOM: 700 };

      await expect(actualRollover.updateRollover(updatedEntry)).rejects.toThrow(errorMessage); // Call actual updateRollover for testing
    });
  });

  describe('resetRollover', () => {
    const mockExpenseList = ['Expense1', 'Expense2'];
    const mockBudget = 100;
    const mockTransactions = [{ Amount: 50 }];
    const mockInitialRollover = { Month: 1, Year: 2023, Expense: 'Expense1', Expenses: 0, BOM: 0, EOM: 0 };
    const mockPreviousRollover = { Month: 12, Year: 2022, Expense: 'Expense1', Expenses: 0, BOM: 0, EOM: 0 };

    beforeEach(() => {
      // Mock common dependencies for resetRollover
      mockedLookups.getExpenseList.mockResolvedValue(mockExpenseList);
      mockedLookups.getBudget.mockResolvedValue(mockBudget);
      mockedLookups.getTransactions.mockResolvedValue(mockTransactions);
      (UpdateTableRow as jest.Mock).mockResolvedValue(null);
      (UpdateTableRows as jest.Mock).mockResolvedValue(null);
      (WriteToTable as jest.Mock).mockResolvedValue(null);

      // Mocks for bulk fetch
      mockedLookups.getRollovers.mockResolvedValue([]);
      mockedLookups.getAllTransactions.mockResolvedValue([]);
      mockedLookups.getAllExpenseData.mockResolvedValue([]);
      mockedLookups.getAllBudgetHistory.mockResolvedValue([]);
    });

    it('should reset rollovers for a single expense', async () => {
      const mockCurrentDate = new Date(2023, 0, 15); // Jan 15, 2023
      const dateSpy = jest.spyOn(global, 'Date').mockImplementation(() => mockCurrentDate as any);

      // Mock data for bulk fetch
      mockedLookups.getAllExpenseData.mockResolvedValue([
          { 'Expense Type': 'Expense1', 'Budget': 100, 'Init': 0 }
      ]);
      mockedLookups.getAllBudgetHistory.mockResolvedValue([]);
      mockedLookups.getRollovers.mockResolvedValue([
          { ...mockInitialRollover, Month: 12, Year: 2022, Expense: 'Expense1', EOM: 50 } // Previous EOM
      ]);
      mockedLookups.getAllTransactions.mockResolvedValue([
          { Month: 1, Year: 2023, 'Expense Type': 'Expense1', Amount: 50 }
      ]);

      await resetRollover(1, 2023, 'Expense1'); // Call the mocked resetRollover

      expect(mockedLookups.getRollovers).toHaveBeenCalled();
      expect(mockedLookups.getAllTransactions).toHaveBeenCalled();
      expect(mockedLookups.getAllExpenseData).toHaveBeenCalled();
      
      // Should perform one update for the existing previous rollover? 
      // Wait, resetRollover calculates current month. 
      // If start month is Jan 2023. Previous is Dec 2022.
      // Current month calc: Expenses=50, BOM=50 (prev EOM), EOM=50+100+50 = 200.
      // Since getRollovers returned Dec 2022, but NOT Jan 2023, this should be a NEW entry.
      // So AddToTable (via WriteToTable) should be called, UpdateTableRows should NOT be called for Jan 2023.
      
      expect(WriteToTable).toHaveBeenCalledWith('Rollovers', expect.arrayContaining([
          expect.arrayContaining([1, 2023, 'Expense1', 50, 50, 200])
      ]));

      dateSpy.mockRestore();
    });

    it('should reset rollovers for all expenses for multiple months', async () => {
        const mockCurrentDate = new Date(2023, 1 /* Feb */, 15); // Current date is Feb 15, 2023
        const dateSpy = jest.spyOn(global, 'Date').mockImplementation(() => mockCurrentDate as any);
  
        // Mock data
        mockedLookups.getAllExpenseData.mockResolvedValue([
            { 'Expense Type': 'Expense1', 'Budget': 100, 'Init': 0 },
            { 'Expense Type': 'Expense2', 'Budget': 100, 'Init': 0 }
        ]);
        
        // Dec 2022 rollovers exist
        mockedLookups.getRollovers.mockResolvedValue([
            { ...mockInitialRollover, Month: 12, Year: 2022, Expense: 'Expense1', EOM: 50 },
            { ...mockInitialRollover, Month: 12, Year: 2022, Expense: 'Expense2', EOM: 70 }
        ]);

        // Transactions for Jan and Feb 2023
        mockedLookups.getAllTransactions.mockResolvedValue([
            { Month: 1, Year: 2023, 'Expense Type': 'Expense1', Amount: 50 },
            { Month: 1, Year: 2023, 'Expense Type': 'Expense2', Amount: 50 },
            { Month: 2, Year: 2023, 'Expense Type': 'Expense1', Amount: 50 },
            { Month: 2, Year: 2023, 'Expense Type': 'Expense2', Amount: 50 }
        ]);
  
        await resetRollover(1, 2023); // Call the actual resetRollover
  
        // Jan 2023: Both are new entries (not in getRollovers).
        // Expense1: BOM=50, Budget=100, Exp=50 -> EOM=200.
        // Expense2: BOM=70, Budget=100, Exp=50 -> EOM=220.
        
        // Feb 2023: Both are new entries (since Jan 2023 were new and not yet in 'getRollovers' return).
        // BUT, my refactored logic adds new entries to the map in memory!
        // So Feb 2023 should find Jan 2023 in memory.
        
        // Expense1 Feb: BOM=200 (Jan EOM), Budget=100, Exp=50 -> EOM=350.
        // Expense2 Feb: BOM=220 (Jan EOM), Budget=100, Exp=50 -> EOM=370.

        // Expect 4 new entries in WriteToTable
        expect(WriteToTable).toHaveBeenCalledWith('Rollovers', expect.arrayContaining([
            expect.arrayContaining([1, 2023, 'Expense1', 50, 50, 200]),
            expect.arrayContaining([1, 2023, 'Expense2', 50, 70, 220]),
            expect.arrayContaining([2, 2023, 'Expense1', 50, 200, 350]),
            expect.arrayContaining([2, 2023, 'Expense2', 50, 220, 370])
        ]));
  
        expect(getExpenseList).toHaveBeenCalled();
  
        dateSpy.mockRestore();
      });

    it('should not reset rollover for a future date', async () => {
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});

      const mockCurrentDate = new Date(2023, 0, 15); // Jan 15, 2023
      const dateSpy = jest.spyOn(global, 'Date').mockImplementation(() => mockCurrentDate as any);

      await resetRollover(2, 2023, 'Expense1'); // Call the actual resetRollover

      expect(consoleErrorSpy).toHaveBeenCalledWith("Cannot reset rollover for a future date.");
      expect(UpdateTableRows).not.toHaveBeenCalled();
      expect(WriteToTable).not.toHaveBeenCalled();

      consoleErrorSpy.mockRestore();
      dateSpy.mockRestore();
    });

    it('should handle month and year rollovers correctly across year boundary', async () => {
      const mockCurrentDate = new Date(2023, 0 /* Jan */, 15); // Current date is Jan 15, 2023
      const dateSpy = jest.spyOn(global, 'Date').mockImplementation(() => mockCurrentDate as any);

      // Mock data
      mockedLookups.getAllExpenseData.mockResolvedValue([
        { 'Expense Type': 'Expense1', 'Budget': 100, 'Init': 0 }
      ]);
      mockedLookups.getRollovers.mockResolvedValue([
        { ...mockInitialRollover, Month: 11, Year: 2022, Expense: 'Expense1', EOM: 50 } // Nov 2022
      ]);
      mockedLookups.getAllTransactions.mockResolvedValue([
        { Month: 12, Year: 2022, 'Expense Type': 'Expense1', Amount: 50 }
      ]);

      await resetRollover(12, 2022, 'Expense1'); // Reset from Dec 2022

      // Dec 2022: New entry. BOM=50, Budget=100, Exp=50 -> EOM=200.
      expect(WriteToTable).toHaveBeenCalledWith('Rollovers', expect.arrayContaining([
          expect.arrayContaining([12, 2022, 'Expense1', 50, 50, 200])
      ]));

      dateSpy.mockRestore();
    });
  });
});