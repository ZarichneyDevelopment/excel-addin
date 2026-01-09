import { getExpenseList, getLastUpdateDate, initializeSchema } from '../lookups';
import { preventDefaults, handleFileDrop } from '../file-drop';
import { resetRollover } from '../rollover';
import { closeErrorConsole, copyErrorToClipboard, handleError } from '../error-handler';
import { logToConsole, logToTaskpane, clearConsole } from '../logger';

const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

let consoleTeeInstalled = false;

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

async function initializeTaskpane() {
  try {
    installConsoleTeeToTaskpane();
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

	async function populateExpenseDropdown() {
	  try {
	    logToConsole('Loading expense categories...', 'info');
	    const expenseList = await getExpenseList();
	    const expenseDropdown = document.getElementById('expense-dropdown') as HTMLSelectElement;

	    // Ensure the dropdown is clear before adding new options
	    expenseDropdown.innerHTML = '<option value="">All Expenses</option>';

    for (const expense of expenseList) {
      const option = document.createElement('option');
      option.value = option.text = expense;
      expenseDropdown.add(option);
	    }
	    logToConsole(`Loaded ${expenseList.length} expense categories.`, 'info');
	  } catch (error) {
	    handleError(error, 'populateExpenseDropdown');
	  }
	}
