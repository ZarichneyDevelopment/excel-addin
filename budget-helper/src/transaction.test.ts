import { Transaction, ProcessTransactions } from './transaction';
import { AddToTable } from './excel-helpers';
import {
  getAccounts,
  getAllTransactionIds,
  getExpenseList,
  getMatchingRules,
  getAmbiguousItems,
} from './lookups';

const globalAny = global as any;
if (!globalAny.crypto) {
  globalAny.crypto = require('crypto').webcrypto;
}

jest.mock('./excel-helpers', () => ({
  AddToTable: jest.fn(),
}));

jest.mock('./lookups', () => ({
  getAccounts: jest.fn(),
  getAllTransactionIds: jest.fn(),
  getExpenseList: jest.fn(),
  getMatchingRules: jest.fn(),
  getAmbiguousItems: jest.fn(),
}));

describe('Transactions', () => {
  let digestSpy: jest.SpyInstance;

  beforeEach(() => {
    jest.clearAllMocks();
    digestSpy = jest.spyOn(globalAny.crypto.subtle, 'digest').mockResolvedValue(
      new Uint8Array([1, 2, 3, 4]).buffer
    );
    // Default mocks
    (getExpenseList as jest.Mock).mockResolvedValue(['Dining', 'Groceries']);
    (getMatchingRules as jest.Mock).mockResolvedValue([]);
    (getAmbiguousItems as jest.Mock).mockResolvedValue([]);
    (getAccounts as jest.Mock).mockResolvedValue({ '12345678': 'Primary Account' });
    (getAllTransactionIds as jest.Mock).mockResolvedValue([]);
  });

  afterEach(() => {
    digestSpy.mockRestore();
  });

  it('parses CSV rows and drops entries without a description', async () => {
    const csv = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Coffee,Shop,12345678,2023-01-15,10.5,
,NoDesc,12345678,2023-01-15,5,
`;

    const transactions = Transaction.fromCsv(csv);
    await new Promise((resolve) => setImmediate(resolve));

    const meaningful = transactions.filter(tx => tx['Description 1']);
    expect(meaningful).toHaveLength(1);
    expect(meaningful[0].Description).toBe('Coffee Shop');
    expect(meaningful[0]['id']).toBe('01020304');
    expect(digestSpy).toHaveBeenCalled();
  });

  it('processes transactions and applies exact matching rules', async () => {
    (getMatchingRules as jest.Mock).mockResolvedValue([
      {
        'Match 1': 'Coffee',
        'Match 2': 'Shop',
        'Expense Type': 'Dining',
      },
    ]);

    const csv = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Coffee,Shop,12345678,2023-01-15,10.5,
`;

    const rows = await ProcessTransactions(csv);

    expect(rows[0][4]).toBe('Primary Account (5678)');
    expect(rows[0][5]).toBe('Dining');
    expect(AddToTable).not.toHaveBeenCalled();
  });

  it('applies aggressive substring matching rules', async () => {
    (getMatchingRules as jest.Mock).mockResolvedValue([
      {
        'Match 1': 'Uber',
        'Match 2': '', // Wildcard/Aggressive
        'Expense Type': 'Transport',
      },
    ]);

    // "Uber Trip" should match "Uber" rule via includes()
    const csv = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Uber Trip,Help,12345678,2023-01-15,15.0,
`;

    const rows = await ProcessTransactions(csv);

    expect(rows[0][5]).toBe('Transport');
  });

  it('flags ambiguous transactions', async () => {
    (getAmbiguousItems as jest.Mock).mockResolvedValue([
      { Item: 'Walmart' },
    ]);
    (getMatchingRules as jest.Mock).mockResolvedValue([
        {
          'Match 1': 'Walmart',
          'Match 2': '', 
          'Expense Type': 'Groceries',
        },
      ]);

    const csv = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Walmart Supercentre,,12345678,2023-01-15,50.0,
`;

    const rows = await ProcessTransactions(csv);

    // Should be 'Ambiguous' even if there is a matching rule, because it hit the ambiguous check first
    expect(rows[0][5]).toBe('Ambiguous');
  });

  it('does NOT write a new matching rule when categorization fails (auto-add removed)', async () => {
    const csv = `Description 1,Description 2,Account Number,Transaction Date,CAD$,USD$
Unknown,Vendor,12345678,2023-01-15,10.5,
`;

    await ProcessTransactions(csv);

    expect(AddToTable).not.toHaveBeenCalled();
  });
});