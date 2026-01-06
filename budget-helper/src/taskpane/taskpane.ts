import { getExpenseList, getLastUpdateDate, initializeSchema } from '../lookups';
import { preventDefaults, handleFileDrop } from '../file-drop';
import { resetRollover } from '../rollover';
import { closeErrorConsole, copyErrorToClipboard, handleError } from '../error-handler';
import { logToConsole, clearConsole } from '../logger';

async function updateLastSyncInfo() {
    try {
        const lastDate = await getLastUpdateDate();
        const display = document.getElementById('last-updated');
        
        if (lastDate) {
            display.textContent = `Synced: ${lastDate.toISOString().split('T')[0]}`;
            // Auto-populate inputs to continue from where we left off
            (document.getElementById('month-input') as HTMLInputElement).value = (lastDate.getMonth() + 1).toString();
            (document.getElementById('year-input') as HTMLInputElement).value = lastDate.getFullYear().toString();
            logToConsole(`Last sync found: ${lastDate.toLocaleDateString()}`, 'info');
        } else {
            display.textContent = 'Synced: Never';
            // Default to today was already set, but good to know
            logToConsole('No previous sync date found.', 'warn');
        }
    } catch (error) {
        console.error("Error fetching last update:", error);
    }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    // Error Console Bindings
    document.getElementById('error-close-btn').addEventListener('click', closeErrorConsole);
    document.getElementById('error-copy-btn').addEventListener('click', copyErrorToClipboard);

    // Console Bindings
    document.getElementById('clear-console').addEventListener('click', clearConsole);

    // Default Date Initialization (Fallback)
    const today = new Date();
    (document.getElementById('month-input') as HTMLInputElement).value = (today.getMonth() + 1).toString();
    (document.getElementById('year-input') as HTMLInputElement).value = today.getFullYear().toString();

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

    window.onload = async () => {
        logToConsole('Initializing...');
        await initializeSchema();
        await populateExpenseDropdown();
        await updateLastSyncInfo();
        logToConsole('Ready.');
    };
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
    await updateLastSyncInfo(); // Refresh UI

  } catch (error) {
    handleError(error, 'TriggerResetRollovers');
    logToConsole('Error during recalculation.', 'error');
  }
}

async function populateExpenseDropdown() {
  try {
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