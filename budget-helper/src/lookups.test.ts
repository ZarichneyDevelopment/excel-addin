import { of } from 'rxjs';
import * as rolloverModule from './rollover';
import {
  getAccounts,
  getBudget,
  getExpenseList,
  getInitialAmount,
  getMatchingRules,
  getAllTransactionIds,
  getRollovers,
  getTransactions,
  getRollover,
} from './lookups';
import { NamedRangeValues$, TableRows$ } from './excel-helpers';

jest.mock('./excel-helpers', () => ({
  NamedRangeValues$: jest.fn(),
  TableRows$: jest.fn(),
}));

jest.mock('./rollover', () => ({
  getRollover: jest.fn(),
}));

describe('Lookup helpers', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('builds account map from table rows', async () => {
    (TableRows$ as jest.Mock).mockReturnValue(of(
      { Number: '111', Name: 'Primary' },
      { Number: '222', Name: 'Savings' }
    ));

    const accounts = await getAccounts();

    expect(accounts).toEqual({
      '111': 'Primary',
      '222': 'Savings',
    });
    expect(TableRows$).toHaveBeenCalledWith('Accounts');
  });

  it('collects matching rules', async () => {
    const rows = [
      { 'Match 1': 'A', 'Match 2': '1' },
      { 'Match 1': 'B', 'Match 2': '2' },
    ];
    (TableRows$ as jest.Mock).mockReturnValue(of(...rows));

    const rules = await getMatchingRules();

    expect(rules).toEqual(rows);
  });

  it('reads the expense list from the ExpenseData table', async () => {
    (TableRows$ as jest.Mock).mockImplementation((tableName: string) => {
      if (tableName === 'ExpenseData') {
        return of(
          { 'Expense Type': 'Rent' },
          { 'Expense Type': 'Food' }
        );
      }
      return of();
    });

    const expenses = await getExpenseList();

    expect(expenses).toEqual(['Rent', 'Food']);
  });

  it('filters transactions by month, year, and expense', async () => {
    const data = [
      { Month: 1, Year: 2023, 'Expense Type': 'Food' },
      { Month: 1, Year: 2022, 'Expense Type': 'Food' },
      { Month: 1, Year: 2023, 'Expense Type': 'Gas' },
    ];
    (TableRows$ as jest.Mock).mockReturnValue(of(...data));

    const result = await getTransactions(1, 2023, 'Food');

    expect(result).toEqual([{ Month: 1, Year: 2023, 'Expense Type': 'Food' }]);
  });

  it('loads existing transaction ids', async () => {
    (NamedRangeValues$ as jest.Mock).mockReturnValue(of('ID1', 'ID2'));

    const ids = await getAllTransactionIds();

    expect(ids).toEqual(['ID1', 'ID2']);
  });

  it('reads rollovers from the Rollovers table', async () => {
    const rollovers = [
      { Expense: 'Rent' },
    ];
    (TableRows$ as jest.Mock).mockReturnValue(of(...rollovers));

    const result = await getRollovers();

    expect(result).toEqual(rollovers);
  });

  it('returns the initial amount for an expense', async () => {
    (TableRows$ as jest.Mock).mockReturnValue(of({ 'Expense Type': 'Groceries', Init: '42' }));

    const amount = await getInitialAmount('Groceries');

    expect(amount).toBe(42);
  });

  it('returns the current budget when no month/year are provided', async () => {
    (TableRows$ as jest.Mock).mockImplementation((tableName: string) => {
      if (tableName === 'ExpenseData') {
        return of({ 'Expense Type': 'Utilities', Budget: '180' });
      }

      return of();
    });

    const budget = await getBudget('Utilities');

    expect(budget).toBe(180);
  });

  it('prefers matching entries from change history', async () => {
    (TableRows$ as jest.Mock).mockImplementation((tableName: string) => {
      if (tableName === 'BudgetHistory') {
        return of({
          Expense: 'Phones',
          'Month Start': 1,
          'Month End': 3,
          'Year Start': 2023,
          'Year End': 2023,
          Amount: '250',
        });
      }

      if (tableName === 'ExpenseData') {
        return of({ 'Expense Type': 'Phones', Budget: '200' });
      }

      return of();
    });

    const budget = await getBudget('Phones', 2, 2023);

    expect(budget).toBe(250);
  });

  it('delegates rollover queries to the rollover module', async () => {
    const rolloverEntry = { Expense: 'Rent' };
    (rolloverModule.getRollover as jest.Mock).mockResolvedValue(rolloverEntry);

    const result = await getRollover(1, 2023, 'Rent');

    expect(result).toBe(rolloverEntry);
    expect(rolloverModule.getRollover).toHaveBeenCalledWith(1, 2023, 'Rent');
  });
});
