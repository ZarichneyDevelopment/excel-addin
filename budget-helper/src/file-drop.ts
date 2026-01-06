import { WriteToTable } from "./excel-helpers";
import { ProcessTransactions } from "./transaction";
import { handleError } from "./error-handler";
import { logToConsole } from "./logger";

export function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

export function handleDragOver(event) {
    event.stopPropagation();
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy'; // Explicitly show this is a copy.
}

export function handleFileDrop(event) {

    var files = event.dataTransfer.files;
    if (files.length > 0) {
        var file = files[0];
        logToConsole(`Reading file: ${file.name}...`, 'info');
        var reader = new FileReader();

        reader.onload = ProcessFileDrop;

        reader.readAsText(file);
    }
}

export async function ProcessFileDrop(event) {
    try {
        var contents = event.target.result;
        logToConsole('Processing transactions...', 'info');

        const transactions = await ProcessTransactions(contents);

        await WriteToTable("Transactions", transactions);
        
        logToConsole(`Successfully imported ${transactions.length} transactions.`, 'success');
    } catch (error) {
        handleError(error, 'ProcessFileDrop');
        logToConsole('Failed to process file.', 'error');
    }
}