/** @jest-environment jsdom */

import { of } from 'rxjs';

jest.mock('../logger', () => ({
  logToConsole: jest.fn(),
  logToTaskpane: jest.fn(),
  clearConsole: jest.fn(),
}));

const mockLookups = {
  initializeSchema: jest.fn().mockResolvedValue(undefined),
  getExpenseList: jest.fn().mockResolvedValue(['Rent', 'Food']),
  getLastUpdateDate: jest.fn().mockResolvedValue(null),
};
jest.mock('../lookups', () => mockLookups);

jest.mock('../rollover', () => ({
  resetRollover: jest.fn().mockResolvedValue(undefined),
}));

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
    jest.clearAllMocks();

    document.body.innerHTML = `
      <button id="error-close-btn"></button>
      <button id="error-copy-btn"></button>
      <button id="clear-console"></button>
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
    expect(mockLookups.getExpenseList).toHaveBeenCalled();
  });
});
