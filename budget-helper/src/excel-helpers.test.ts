import * as ExcelHelpers from './excel-helpers';
import { of } from 'rxjs';

describe('Excel Helpers', () => {
  let context: any;

  beforeEach(() => {
    context = {
      sync: jest.fn().mockResolvedValue(undefined),
      workbook: {
        names: {
          getItem: jest.fn(),
          getItemOrNullObject: jest.fn(),
          add: jest.fn(),
        },
        tables: {
          getItem: jest.fn(),
          getItemOrNullObject: jest.fn(),
        },
        worksheets: {
          getItemOrNullObject: jest.fn(),
          add: jest.fn(),
        },
      },
    };

    global.Excel = {
      run: jest.fn(async (callback) => {
        await callback(context);
      }),
    } as any;
  });

  afterEach(() => {
    jest.restoreAllMocks();
    delete global.Excel;
  });

  describe('WriteToTable', () => {
    it('adds rows when data is provided', async () => {
      const addSpy = jest.fn().mockReturnValue({});
      context.workbook.tables.getItemOrNullObject.mockReturnValue({
        isNullObject: false,
        rows: {
          add: addSpy,
        },
      });

      await ExcelHelpers.WriteToTable('Test', [['value']]);

      expect(context.workbook.tables.getItemOrNullObject).toHaveBeenCalledWith('Test');
      expect(addSpy).toHaveBeenCalledWith(null, [['value']]);
      expect(context.sync).toHaveBeenCalled();
    });

    it('warns when no data is provided', async () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

      await ExcelHelpers.WriteToTable('Test', []);

      expect(warnSpy).toHaveBeenCalledWith("No data to write to table 'Test'.");
      warnSpy.mockRestore();
    });
  });

  describe('NamedRangeValues$', () => {
    it('emits unique, non-empty values from the given range', async () => {
      const usedRange = {
        load: jest.fn().mockReturnThis(),
        values: [['Apples', ''], ['Bananas', 'Apples']],
      };
      const rangeAccessor = {
        getUsedRange: jest.fn(() => usedRange),
      };
      const namedItem = {
        getRange: jest.fn(() => rangeAccessor),
        load: jest.fn().mockReturnThis(),
        type: 'Range',
        isNullObject: false,
      };

      context.workbook.names.getItemOrNullObject.mockReturnValue(namedItem);

      const emitted: string[] = [];
      await new Promise<void>((resolve) => {
        ExcelHelpers.NamedRangeValues$('Expenses').subscribe({
          next: (value) => emitted.push(value),
          complete: resolve,
        });
      });

      expect(emitted).toEqual(['Apples', 'Bananas']);
      expect(namedItem.getRange).toHaveBeenCalled();
      expect(usedRange.load).toHaveBeenCalledWith('values');
    });
  });

  describe('TableRows$', () => {
    it('maps header/data rows into objects', async () => {
      const headerRow = { load: jest.fn().mockReturnThis(), values: [['ID', 'Name']] };
      const dataRows = { load: jest.fn().mockReturnThis(), values: [[1, 'Alpha'], [2, 'Beta']] };
      const table = {
        getHeaderRowRange: jest.fn(() => headerRow),
        getDataBodyRange: jest.fn(() => dataRows),
      };

      context.workbook.tables.getItem.mockReturnValue(table);

      const emitted: Record<string, unknown>[] = [];
      await new Promise<void>((resolve) => {
        ExcelHelpers.TableRows$('Atoms').subscribe({
          next: (row) => emitted.push(row),
          complete: resolve,
        });
      });

      expect(emitted).toEqual([
        { ID: 1, Name: 'Alpha' },
        { ID: 2, Name: 'Beta' },
      ]);
      expect(table.getHeaderRowRange).toHaveBeenCalled();
      expect(dataRows.load).toHaveBeenCalled();
    });
  });

  describe('AddToTable', () => {
    it('converts objects to rows and delegates to WriteToTable', async () => {
      const addSpy = jest.fn().mockReturnValue({});
      context.workbook.tables.getItemOrNullObject.mockReturnValue({
        isNullObject: false,
        rows: {
          add: addSpy,
        },
      });

      await ExcelHelpers.AddToTable('MatchingRules', { foo: 'bar', baz: 5 });

      expect(addSpy).toHaveBeenCalledWith(null, [['bar', 5]]);
    });
  });

  describe('UpdateTableRow', () => {
    it('updates the targeted row with values derived from headers', async () => {
      const rowRange = { values: [] };
      const dataBodyRange = { getRow: jest.fn(() => rowRange) };
      const headerRow = { load: jest.fn().mockReturnThis(), values: [['Expense', 'Amount']] };
      const table = {
        getDataBodyRange: jest.fn(() => dataBodyRange),
        getHeaderRowRange: jest.fn(() => headerRow),
      };

      context.workbook.tables.getItem.mockReturnValue(table);

      const entry = { Expense: 'Rent', Amount: 120 };
      await ExcelHelpers.UpdateTableRow('Rollovers', 1, entry);

      expect(headerRow.load).toHaveBeenCalledWith('values');
      expect(context.sync).toHaveBeenCalledTimes(2);
      expect(rowRange.values).toEqual([['Rent', 120]]);
    });
  });

  describe('EnsureTableExists', () => {
    it('creates sheet and table if they do not exist', async () => {
      // Mock table not found
      context.workbook.tables.getItemOrNullObject.mockReturnValue({ isNullObject: true });
      
      // Mock sheet not found
      context.workbook.worksheets.getItemOrNullObject.mockReturnValue({ isNullObject: true });
      
      // Mock creation
      const mockTable = { name: '', getHeaderRowRange: jest.fn().mockReturnValue({ values: [] }) };
      const mockSheet = { tables: { add: jest.fn().mockReturnValue(mockTable) } };
      context.workbook.worksheets.add.mockReturnValue(mockSheet);

      await ExcelHelpers.EnsureTableExists('NewTable', ['Col1', 'Col2']);

      expect(context.workbook.worksheets.add).toHaveBeenCalledWith('NewTable');
      expect(mockSheet.tables.add).toHaveBeenCalledWith('A1:B1', true);
      expect(mockTable.getHeaderRowRange().values).toEqual([['Col1', 'Col2']]);
    });

    it('does nothing if table already exists', async () => {
      context.workbook.tables.getItemOrNullObject.mockReturnValue({ isNullObject: false });

      await ExcelHelpers.EnsureTableExists('ExistingTable', ['Col1']);

      expect(context.workbook.worksheets.add).not.toHaveBeenCalled();
    });
  });

  describe('SetNamedRangeValue', () => {
    it('lazily creates named range if missing', async () => {
      const mockRange = { values: [] };
      context.workbook.worksheets.getItem = jest.fn().mockReturnValue({
        getRange: jest.fn().mockReturnValue(mockRange),
      });

      await ExcelHelpers.SetNamedRangeValue('LastRolloverUpdate', '2023-01-01');

      expect(mockRange.values).toEqual([['2023-01-01']]);
    });
  });
});
