import { getAllBudgetHistory, getExpenseList, getLastUpdateDate, initializeSchema } from '../lookups';
import { preventDefaults, handleFileDrop } from '../file-drop';
import { resetRollover } from '../rollover';
import { closeErrorConsole, copyErrorToClipboard, handleError } from '../error-handler';
import { logToConsole, logToTaskpane, clearConsole } from '../logger';
import { WriteToTable } from '../excel-helpers';

declare const __ADDIN_VERSION__: string | undefined;

const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

let consoleTeeInstalled = false;
const defaultFeature = 'ingestion';
const budgetUpdateFallback = { month: 2, year: 2024 };

function safeFormatConsoleArgs(args: unknown[]): string {
  const formatted = args.map(arg => {
    if (typeof arg === 'string') return arg;
    try {
      return JSON.stringify(arg);
    } catch {
      return String(arg);
    }
  }).join(' ');

  const maxLen = 2000;
  if (formatted.length <= maxLen) return formatted;
  return formatted.slice(0, maxLen) + '…';
}

function installConsoleTeeToTaskpane() {
  if (consoleTeeInstalled) return;
  consoleTeeInstalled = true;

  const original = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  };

  console.log = (...args: unknown[]) => {
    original.log(...args);
    logToTaskpane(safeFormatConsoleArgs(args), 'info');
  };
  console.warn = (...args: unknown[]) => {
    original.warn(...args);
    logToTaskpane(safeFormatConsoleArgs(args), 'warn');
  };
  console.error = (...args: unknown[]) => {
    original.error(...args);
    logToTaskpane(safeFormatConsoleArgs(args), 'error');
  };
}

async function ensureDomReady(): Promise<void> {
  if (document.readyState !== 'loading') return;
  await new Promise<void>((resolve) => {
    document.addEventListener('DOMContentLoaded', () => resolve(), { once: true });
  });
}

async function initializeTaskpane() {
  try {
    await ensureDomReady();

    installConsoleTeeToTaskpane();
    initializeFeatureSwitcher();
    const versionLabel = document.getElementById('app-version');
    if (versionLabel) {
      const version = typeof __ADDIN_VERSION__ !== 'undefined' ? __ADDIN_VERSION__ : 'dev';
      versionLabel.textContent = `v${version}`;
    }
    logToConsole('Verbose logging enabled (capturing console output).', 'info');
    logToConsole('Initializing...', 'info');

    await initializeSchema();
    await populateExpenseDropdown();
    await updateLastSyncInfo();
    updateRecalcButtonLabel();

    logToConsole('Ready.', 'success');
  } catch (error) {
    handleError(error, 'initializeTaskpane');
    logToConsole('Initialization failed.', 'error');
  }
}

function getNumberInput(id: string): HTMLInputElement | null {
  const element = document.getElementById(id);
  if (!element) return null;
  return element as HTMLInputElement;
}

function getSelect(id: string): HTMLSelectElement | null {
  const element = document.getElementById(id);
  if (!element) return null;
  return element as HTMLSelectElement;
}

function setActiveFeature(feature: string) {
  const panels = Array.from(document.querySelectorAll<HTMLElement>('.feature-panel'));
  if (panels.length === 0) return;

  let matched = false;
  panels.forEach(panel => {
    const panelFeature = panel.dataset.feature;
    const isActive = panelFeature === feature;
    panel.classList.toggle('is-hidden', !isActive);
    if (isActive) matched = true;
  });

  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>('.feature-tab'));
  tabs.forEach(tab => {
    const tabFeature = tab.dataset.feature;
    const isActive = tabFeature === feature;
    tab.classList.toggle('is-active', isActive);
  });

  if (!matched && feature !== defaultFeature) {
    setActiveFeature(defaultFeature);
  }
}

function initializeFeatureSwitcher() {
  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>('.feature-tab'));
  if (tabs.length === 0) return;

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const feature = tab.dataset.feature || defaultFeature;
      setActiveFeature(feature);
    });
  });

  setActiveFeature(defaultFeature);
}


function setSelectOptions(select: HTMLSelectElement, options: string[], defaultLabel: string) {
  select.replaceChildren();
  const defaultOption = document.createElement('option');
  defaultOption.value = '';
  defaultOption.text = defaultLabel;
  select.appendChild(defaultOption);

  for (const optionLabel of options) {
    const option = document.createElement('option');
    option.value = option.text = optionLabel;
    select.appendChild(option);
  }
}

function parseAmountInput(value: string): number | null {
  const cleaned = value.replace(/,/g, '').trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed) || parsed <= 0) return null;
  return parsed;
}

function parseNonNegativeAmountInput(value: string): number | null {
  const cleaned = value.replace(/,/g, '').trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed) || parsed < 0) return null;
  return parsed;
}

function formatAmount(amount: number): string {
  return amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function createTransferRef(): string {
  try {
    if (typeof crypto !== 'undefined' && typeof (crypto as Crypto).randomUUID === 'function') {
      return (crypto as Crypto).randomUUID();
    }
  } catch {
    // Fall through to deterministic fallback.
  }
  return `xfer-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function parseHistoryNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const parsed = parseFloat(String(value));
  return Number.isFinite(parsed) ? parsed : null;
}

function getLatestHistoryEndIndex(historyRows: any[], expense: string): number | null {
  let latest: number | null = null;
  for (const row of historyRows) {
    if (row['Expense'] !== expense) continue;
    const monthEnd = parseHistoryNumber(row['Month End']);
    const yearEnd = parseHistoryNumber(row['Year End']);
    if (monthEnd === null || yearEnd === null) continue;
    const endIndex = monthYearToIndex(monthEnd, yearEnd);
    if (latest === null || endIndex > latest) {
      latest = endIndex;
    }
  }
  return latest;
}

async function updateExpenseBudgetValue(expense: string, newAmount: number): Promise<number> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem('ExpenseData');
    const headerRange = table.getHeaderRowRange().load('values');
    const dataRange = table.getDataBodyRange().load('values');
    await context.sync();

    const headers = headerRange.values[0] as string[];
    const expenseIndex = headers.indexOf('Expense Type');
    const budgetIndex = headers.indexOf('Budget');
    if (expenseIndex === -1 || budgetIndex === -1) {
      throw new Error('ExpenseData table is missing expected columns.');
    }

    const rows = dataRange.values;
    const rowIndex = rows.findIndex(row => String(row[expenseIndex]).trim() === expense);
    if (rowIndex === -1) {
      throw new Error(`Expense "${expense}" was not found in ExpenseData.`);
    }

    const oldRaw = rows[rowIndex][budgetIndex];
    const oldBudget = parseHistoryNumber(oldRaw) ?? 0;
    dataRange.getCell(rowIndex, budgetIndex).values = [[newAmount]];
    await context.sync();
    return oldBudget;
  });
}

function getStartMonthYearFromInputs(): { month: number | null, year: number | null } {
  const monthInput = getNumberInput('month-input');
  const yearInput = getNumberInput('year-input');
  const month = monthInput ? parseInt(monthInput.value, 10) : NaN;
  const year = yearInput ? parseInt(yearInput.value, 10) : NaN;
  return {
    month: Number.isFinite(month) ? month : null,
    year: Number.isFinite(year) ? year : null,
  };
}

function monthYearToIndex(month: number, year: number): number {
  // 0-based month index
  return year * 12 + (month - 1);
}

function monthYearFromIndex(index: number): { month: number; year: number } {
  const year = Math.floor(index / 12);
  const month = (index % 12) + 1;
  return { month, year };
}

function formatMonthYear(month: number, year: number): string {
  const name = monthNames[month - 1] ?? `M${month}`;
  return `${name} ${year}`;
}

function computeRecalcRangeLabel(startMonth: number, startYear: number): string | null {
  if (startMonth < 1 || startMonth > 12 || startYear < 1900 || startYear > 2500) return null;

  const today = new Date();
  const endMonth = today.getMonth() + 1;
  const endYear = today.getFullYear();

  const startIndex = monthYearToIndex(startMonth, startYear);
  const endIndex = monthYearToIndex(endMonth, endYear);
  if (startIndex > endIndex) return null;

  const months = (endIndex - startIndex) + 1;
  return `Recalc ${formatMonthYear(startMonth, startYear)} → ${formatMonthYear(endMonth, endYear)} (${months} mo)`;
}

function updateRecalcButtonLabel() {
  const button = document.getElementById('reset');
  if (!button) return;

  const { month, year } = getStartMonthYearFromInputs();
  if (!month || !year) {
    button.textContent = 'Recalc';
    return;
  }

  const label = computeRecalcRangeLabel(month, year);
  button.textContent = label ?? 'Recalc';
}

function markStartDateUserEdited() {
  const monthInput = getNumberInput('month-input');
  const yearInput = getNumberInput('year-input');
  if (monthInput?.dataset) monthInput.dataset.userEdited = '1';
  if (yearInput?.dataset) yearInput.dataset.userEdited = '1';
}

function startDateWasUserEdited(): boolean {
  const monthInput = getNumberInput('month-input');
  const yearInput = getNumberInput('year-input');
  return Boolean(monthInput?.dataset?.userEdited || yearInput?.dataset?.userEdited);
}

async function updateLastSyncInfo() {
    try {
        const lastDate = await getLastUpdateDate();
        const display = document.getElementById('last-updated');
        
        if (lastDate) {
            if (display) display.textContent = `Synced: ${lastDate.toISOString().split('T')[0]}`;
            // Auto-populate inputs to continue from where we left off, but don't stomp user edits.
            if (!startDateWasUserEdited()) {
              const monthInput = getNumberInput('month-input');
              const yearInput = getNumberInput('year-input');
              if (monthInput) monthInput.value = (lastDate.getMonth() + 1).toString();
              if (yearInput) yearInput.value = lastDate.getFullYear().toString();
              updateRecalcButtonLabel();
            }
            logToConsole(`Last sync found: ${lastDate.toLocaleDateString()}`, 'info');
        } else {
            if (display) display.textContent = 'Synced: Never';
            // Default to today was already set, but good to know
            logToConsole('No previous sync date found.', 'warn');
        }
    } catch (error) {
        console.error("Error fetching last update:", error);
    }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    installConsoleTeeToTaskpane();
    logToConsole('Office ready (Excel).', 'info');

    // Error Console Bindings
    document.getElementById('error-close-btn').addEventListener('click', closeErrorConsole);
    document.getElementById('error-copy-btn').addEventListener('click', copyErrorToClipboard);

    // Console Bindings
    document.getElementById('clear-console').addEventListener('click', clearConsole);

    // Default Date Initialization (Fallback)
    const today = new Date();
    const monthInput = getNumberInput('month-input');
    const yearInput = getNumberInput('year-input');
    if (monthInput) monthInput.value = (today.getMonth() + 1).toString();
    if (yearInput) yearInput.value = today.getFullYear().toString();

    // Drop Zone Setup
    let dropArea = document.getElementById('drop-area');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
      dropArea.addEventListener(eventName, preventDefaults, false)
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropArea.addEventListener(eventName, (e) => {
        dropArea.classList.add('highlight');
      }, false)
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropArea.addEventListener(eventName, (e) => {
        dropArea.classList.remove('highlight');
      }, false)
    });

    dropArea.addEventListener('drop', handleFileDrop, false);

    document.getElementById("reset").onclick = TriggerResetRollovers;

    // Keep UX stable: if user edits start date, never overwrite it automatically.
    monthInput?.addEventListener?.('input', () => {
      markStartDateUserEdited();
      updateRecalcButtonLabel();
    });
    yearInput?.addEventListener?.('input', () => {
      markStartDateUserEdited();
      updateRecalcButtonLabel();
    });

    const expenseSelect = getSelect('expense-dropdown');
    expenseSelect?.addEventListener?.('change', () => updateRecalcButtonLabel());

    document.addEventListener('budgethelper:inputs-changed', () => updateRecalcButtonLabel());

    const transferButton = document.getElementById('transfer-btn');
    transferButton?.addEventListener?.('click', TriggerBudgetTransfer);
    const budgetUpdateButton = document.getElementById('budget-update-btn');
    budgetUpdateButton?.addEventListener?.('click', TriggerBudgetUpdate);

    // `window.onload` can fire before this handler is assigned in Office taskpanes.
    // Initialize immediately once Office is ready and the DOM is present.
    void initializeTaskpane();
  }
});

export async function TriggerResetRollovers() {
  try {
    logToConsole('Starting rollover recalculation...', 'info');

    var month = parseInt((<HTMLInputElement>document.getElementById("month-input")).value);
    var year = parseInt((<HTMLInputElement>document.getElementById("year-input")).value);

    let selectElement = document.getElementById('expense-dropdown') as HTMLSelectElement;
    let selectedOption = selectElement.options[selectElement.selectedIndex];
    let selectedExpense = selectedOption.text;

    if (selectedExpense === "All Expenses") {
      selectedExpense = null;
    }

    await Excel.run(async (context) => {
      await resetRollover(month, year, selectedExpense);
      await context.sync();
    });

    logToConsole('Recalculation complete.', 'success');
    await updateLastSyncInfo(); // Refresh header (doesn't overwrite user-edited inputs)

  } catch (error) {
    handleError(error, 'TriggerResetRollovers');
    logToConsole('Error during recalculation.', 'error');
  }
}

export async function TriggerBudgetTransfer() {
  try {
    const fromSelect = getSelect('transfer-from');
    const toSelect = getSelect('transfer-to');
    const amountInput = getNumberInput('transfer-amount');

    if (!fromSelect || !toSelect || !amountInput) {
      logToConsole('Transfer controls are missing.', 'error');
      return;
    }

    const from = fromSelect.value.trim();
    const to = toSelect.value.trim();
    const amount = parseAmountInput(amountInput.value);

    if (!from) {
      logToConsole('Select a source bucket for the transfer.', 'warn');
      return;
    }

    if (!to) {
      logToConsole('Select a destination bucket for the transfer.', 'warn');
      return;
    }

    if (from === to) {
      logToConsole('Source and destination buckets must be different.', 'warn');
      return;
    }

    if (amount === null) {
      logToConsole('Enter a transfer amount greater than zero.', 'warn');
      return;
    }

    const now = new Date();
    const month = now.getMonth() + 1;
    const year = now.getFullYear();
    const date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const account = 'Budget Transfer';
    const description = 'Bucket Transfer';

    const transferRef = createTransferRef();
    const memoOut = `Transfer to ${to}`;
    const memoIn = `Transfer from ${from}`;

    const rows = [
      [`${transferRef}-out`, month, year, date, account, from, -amount, description, memoOut],
      [`${transferRef}-in`, month, year, date, account, to, amount, description, memoIn],
    ];

    await WriteToTable('Transactions', rows);

    logToConsole(`Transfer saved: ${formatAmount(amount)} from ${from} → ${to}.`, 'success');
    amountInput.value = '';
  } catch (error) {
    handleError(error, 'TriggerBudgetTransfer');
    logToConsole('Error adding transfer.', 'error');
  }
}

export async function TriggerBudgetUpdate() {
  try {
    const expenseSelect = getSelect('budget-update-expense');
    const amountInput = getNumberInput('budget-update-amount');

    if (!expenseSelect || !amountInput) {
      logToConsole('Budget update controls are missing.', 'error');
      return;
    }

    const expense = expenseSelect.value.trim();
    const newAmount = parseNonNegativeAmountInput(amountInput.value);

    if (!expense) {
      logToConsole('Select an expense to update.', 'warn');
      return;
    }

    if (newAmount === null) {
      logToConsole('Enter a new budget amount (zero or greater).', 'warn');
      return;
    }

    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentYear = now.getFullYear();
    const currentIndex = monthYearToIndex(currentMonth, currentYear);

    const historyRows = await getAllBudgetHistory();
    const lastEndIndex = getLatestHistoryEndIndex(historyRows, expense);

    let startIndex: number;
    let shouldWriteHistory = true;

    if (lastEndIndex !== null) {
      const lastEnd = monthYearFromIndex(lastEndIndex);
      logToConsole(`Latest history for ${expense} ends ${formatMonthYear(lastEnd.month, lastEnd.year)}.`, 'info');
      if (lastEndIndex >= currentIndex) {
        logToConsole('Existing history already covers the current month; skipping history insert.', 'warn');
        shouldWriteHistory = false;
        startIndex = currentIndex;
      } else {
        startIndex = lastEndIndex + 1;
      }
    } else {
      const fallbackIndex = monthYearToIndex(budgetUpdateFallback.month, budgetUpdateFallback.year);
      logToConsole(
        `No history for ${expense}; using fallback start (${formatMonthYear(budgetUpdateFallback.month, budgetUpdateFallback.year)}).`,
        'info'
      );
      startIndex = fallbackIndex;
    }

    if (startIndex > currentIndex) {
      logToConsole('History range start is after the current month; using current month instead.', 'warn');
      startIndex = currentIndex;
    }

    const start = monthYearFromIndex(startIndex);

    logToConsole(`Updating ${expense} budget to ${formatAmount(newAmount)} (effective next month).`, 'info');

    const oldAmount = await updateExpenseBudgetValue(expense, newAmount);
    const oldAmountLabel = formatAmount(oldAmount);

    if (shouldWriteHistory) {
      const historyRow = [
        expense,
        start.month,
        start.year,
        null,
        currentMonth,
        currentYear,
        null,
        oldAmount,
      ];

      await WriteToTable('BudgetHistory', [historyRow]);

      logToConsole(
        `Logged ${expense} history: ${oldAmountLabel} from ${formatMonthYear(start.month, start.year)} → ${formatMonthYear(currentMonth, currentYear)}.`,
        'success'
      );
    }

    logToConsole(`Updated ${expense} budget to ${formatAmount(newAmount)}.`, 'success');

    amountInput.value = '';
  } catch (error) {
    handleError(error, 'TriggerBudgetUpdate');
    logToConsole('Error applying budget update.', 'error');
  }
}

async function populateExpenseDropdown() {
  try {
    logToConsole('Loading expense categories...', 'info');
    const expenseList = await getExpenseList();
    const expenseDropdown = document.getElementById('expense-dropdown') as HTMLSelectElement;
    if (!expenseDropdown) {
      logToConsole('Expense dropdown element not found.', 'error');
      return;
    }

    setSelectOptions(expenseDropdown, expenseList, 'All Expenses');
    const transferFrom = getSelect('transfer-from');
    if (transferFrom) {
      setSelectOptions(transferFrom, expenseList, 'From');
    }
    const transferTo = getSelect('transfer-to');
    if (transferTo) {
      setSelectOptions(transferTo, expenseList, 'To');
    }
    const budgetUpdateSelect = getSelect('budget-update-expense');
    if (budgetUpdateSelect) {
      setSelectOptions(budgetUpdateSelect, expenseList, 'Select');
    }

    const sample = expenseList.slice(0, 3).join(', ');
    logToConsole(`Loaded ${expenseList.length} expense categories.`, 'info');
    logToConsole(`Dropdown now has ${expenseDropdown.options.length} options (sample: ${sample || 'n/a'}).`, 'info');
  } catch (error) {
    handleError(error, 'populateExpenseDropdown');
  }
}
