/** @jest-environment jsdom */

jest.mock('../logger', () => ({
  logToConsole: jest.fn(),
  logToTaskpane: jest.fn(),
  clearConsole: jest.fn(),
}));

const mockLookups = {
  initializeSchema: jest.fn().mockResolvedValue(undefined),
  getExpenseList: jest.fn().mockResolvedValue(['Rent', 'Food']),
  getLastUpdateDate: jest.fn().mockResolvedValue(null),
  getAllBudgetHistory: jest.fn().mockResolvedValue([]),
};
jest.mock('../lookups', () => mockLookups);

jest.mock('../rollover', () => ({
  resetRollover: jest.fn().mockResolvedValue(undefined),
}));

const mockExcelHelpers = {
  WriteToTable: jest.fn(),
};
jest.mock('../excel-helpers', () => mockExcelHelpers);

jest.mock('../file-drop', () => ({
  preventDefaults: jest.fn(),
  handleFileDrop: jest.fn(),
}));

jest.mock('../error-handler', () => ({
  closeErrorConsole: jest.fn(),
  copyErrorToClipboard: jest.fn(),
  handleError: jest.fn(),
}));

describe('Taskpane', () => {
  const originalConsole = {
    log: console.log,
    warn: console.warn,
    error: console.error,
  };

  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();
    Object.defineProperty(document, 'readyState', { value: 'complete', configurable: true });

    document.body.innerHTML = `
      <button id="error-close-btn"></button>
      <button id="error-copy-btn"></button>
      <button id="clear-console"></button>
      <nav class="feature-tabs">
        <button class="feature-tab is-active" data-feature="ingestion"></button>
        <button class="feature-tab" data-feature="recalc"></button>
        <button class="feature-tab" data-feature="transfer"></button>
        <button class="feature-tab" data-feature="budget-update"></button>
      </nav>
      <section class="feature-panel" data-feature="ingestion"></section>
      <section class="feature-panel" data-feature="recalc"></section>
      <section class="feature-panel" data-feature="transfer"></section>
      <section class="feature-panel" data-feature="budget-update"></section>
      <div id="drop-area"></div>
      <button id="reset"></button>
      <input id="month-input" type="number" />
      <input id="year-input" type="number" />
      <span id="app-version"></span>
      <span id="last-updated"></span>
      <select id="expense-dropdown"><option value="">(stale)</option></select>
      <select id="transfer-from"><option value="">(stale)</option></select>
      <select id="transfer-to"><option value="">(stale)</option></select>
      <input id="transfer-amount" type="number" />
      <button id="transfer-btn"></button>
      <select id="budget-update-expense"><option value="">(stale)</option></select>
      <input id="budget-update-amount" type="number" />
      <button id="budget-update-btn"></button>
      <div id="console-output"></div>
    `;

    // Minimal Office shim to trigger taskpane setup on import.
    (global as any).Office = {
      HostType: { Excel: 'Excel' },
      onReady: jest.fn(async (cb: any) => cb({ host: 'Excel' })),
      context: {
        document: {
          settings: {
            get: jest.fn(),
            set: jest.fn(),
            saveAsync: jest.fn((cb: any) => cb({ status: 'succeeded' })),
          },
        },
      },
      AsyncResultStatus: { Failed: 'failed' },
    };

    // Excel.run is not used by window.onload, but some imports reference it.
    (global as any).Excel = {
      run: jest.fn(async (cb: any) => cb({ sync: jest.fn(), workbook: {} })),
    };
  });

  afterEach(() => {
    jest.useRealTimers();
    console.log = originalConsole.log;
    console.warn = originalConsole.warn;
    console.error = originalConsole.error;
  });

  it('populates the expense dropdown with all Expense Types', async () => {
    await import('./taskpane');

    // Initialization runs automatically on Office.onReady; wait for async setup.
    await new Promise((resolve) => setTimeout(resolve, 0));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const select = document.getElementById('expense-dropdown') as HTMLSelectElement;
    expect(select).toBeTruthy();
    expect(select.options.length).toBe(3);
    expect(select.options[0].text).toBe('All Expenses');
    expect(select.options[1].text).toBe('Rent');
    expect(select.options[2].text).toBe('Food');
    const transferFrom = document.getElementById('transfer-from') as HTMLSelectElement;
    const transferTo = document.getElementById('transfer-to') as HTMLSelectElement;
    expect(transferFrom.options.length).toBe(3);
    expect(transferFrom.options[0].text).toBe('From');
    expect(transferFrom.options[1].text).toBe('Rent');
    expect(transferFrom.options[2].text).toBe('Food');
    expect(transferTo.options.length).toBe(3);
    expect(transferTo.options[0].text).toBe('To');
    expect(transferTo.options[1].text).toBe('Rent');
    expect(transferTo.options[2].text).toBe('Food');
    const budgetUpdate = document.getElementById('budget-update-expense') as HTMLSelectElement;
    expect(budgetUpdate.options.length).toBe(3);
    expect(budgetUpdate.options[0].text).toBe('Select');
    expect(budgetUpdate.options[1].text).toBe('Rent');
    expect(budgetUpdate.options[2].text).toBe('Food');
    expect(mockLookups.getExpenseList).toHaveBeenCalled();
  });

  it('adds paired transfer rows for bucket transfers', async () => {
    const { TriggerBudgetTransfer } = await import('./taskpane');

    await new Promise((resolve) => setTimeout(resolve, 0));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const fromSelect = document.getElementById('transfer-from') as HTMLSelectElement;
    const toSelect = document.getElementById('transfer-to') as HTMLSelectElement;
    const amountInput = document.getElementById('transfer-amount') as HTMLInputElement;

    fromSelect.value = 'Rent';
    toSelect.value = 'Food';
    amountInput.value = '25';

    await TriggerBudgetTransfer();

    expect(mockExcelHelpers.WriteToTable).toHaveBeenCalledTimes(1);
    const [tableName, rows] = mockExcelHelpers.WriteToTable.mock.calls[0];
    expect(tableName).toBe('Transactions');
    expect(rows).toHaveLength(2);
    expect(rows[0][5]).toBe('Rent');
    expect(rows[0][6]).toBe(-25);
    expect(rows[0][8]).toBe('Transfer to Food');
    expect(rows[1][5]).toBe('Food');
    expect(rows[1][6]).toBe(25);
    expect(rows[1][8]).toBe('Transfer from Rent');
  });

  it('updates budget and logs ChangeHistory using fallback start date', async () => {
    const headerRange = { values: [['Expense Type', 'Budget']], load: jest.fn().mockReturnThis() };
    const cell = { values: null as any };
    const dataRange = {
      values: [
        ['Rent', 500],
        ['Food', 300],
      ],
      load: jest.fn().mockReturnThis(),
      getCell: jest.fn(() => cell),
    };
    const table = {
      getHeaderRowRange: jest.fn(() => headerRange),
      getDataBodyRange: jest.fn(() => dataRange),
    };

    (global as any).Excel.run = jest.fn(async (cb: any) =>
      cb({
        workbook: { tables: { getItem: jest.fn(() => table) } },
        sync: jest.fn(),
      })
    );

    const { TriggerBudgetUpdate } = await import('./taskpane');

    await new Promise((resolve) => setTimeout(resolve, 0));
    await new Promise((resolve) => setTimeout(resolve, 0));

    jest.useFakeTimers().setSystemTime(new Date('2024-06-15T12:00:00Z'));

    const expenseSelect = document.getElementById('budget-update-expense') as HTMLSelectElement;
    const amountInput = document.getElementById('budget-update-amount') as HTMLInputElement;
    expenseSelect.value = 'Rent';
    amountInput.value = '650';

    await TriggerBudgetUpdate();

    expect(cell.values).toEqual([[650]]);
    expect(mockExcelHelpers.WriteToTable).toHaveBeenCalledWith('BudgetHistory', [[
      'Rent',
      2,
      2024,
      null,
      6,
      2024,
      null,
      500,
    ]]);

  });
});
