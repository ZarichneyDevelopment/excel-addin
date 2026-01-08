
import { Observable } from 'rxjs';

export function NamedRangeValues$(rangeName: string): Observable<string> {
    return new Observable(observer => {
        Excel.run(async (context) => {
            const namedItem = context.workbook.names.getItem(rangeName);
            const range = namedItem.getRange().getUsedRange();
            range.load('values');

            try {
                await context.sync();
            } catch (error) {
                console.error("Error getting range named '" + rangeName + "':", error);
                throw error; // Rethrow to ensure calling code is aware of the failure
            }

            if (range.values && range.values.length > 0 && range.values[0].length > 0) {
                // Flatten the 2D array to 1D and filter out empty values
                const values = range.values.flat().filter((value, index, self) => value !== "" && self.indexOf(value) === index);

                values.forEach(value => observer.next(value));
            } else {
                observer.next('Named range is empty or does not exist');
            }

            observer.complete();
        }).catch(error => observer.error(error));
    });
}

export function TableRows$(tableName: string): Observable<any> {
    return new Observable(observer => {
        Excel.run(async (context) => {
            const table = context.workbook.tables.getItem(tableName);
            const headerRow = table.getHeaderRowRange().load('values');
            const dataRows = table.getDataBodyRange().load('values');

            try {
                await context.sync();
            } catch (error) {
                console.error("Error getting rows for table '" + tableName + "':", error);
                throw error; // Rethrow to ensure calling code is aware of the failure
            }

            const headers = headerRow.values[0];
            const rows = dataRows.values;

            rows.forEach(row => {
                const rowObject = headers.reduce((obj, header, index) => {
                    obj[header] = row[index];
                    return obj;
                }, {});

                observer.next(rowObject);
            });

            observer.complete();
        }).catch(error => observer.error(error));
    });
}

export async function AddToTable(tableName: string, data: any) {
    // Turn any object into an array of values
    var row = Object.values(data);
    WriteToTable(tableName, [row]);
}

export async function WriteToTable(tableName: string, data: any[]) {
    try {
        console.log(`WriteToTable called for '${tableName}' with ${data.length} rows.`);

        if (data.length === 0) {
            console.warn("No data to write to table '" + tableName + "'.");
            return;
        }

        await Excel.run(async (context) => {
            console.log(`WriteToTable: Getting table '${tableName}'...`);
            const table = context.workbook.tables.getItemOrNullObject(tableName);
            await context.sync(); // Ensure table is loaded or null if not found

            if (table.isNullObject) {
                throw new Error(`Table "${tableName}" not found.`);
            }

            console.log(`WriteToTable: Adding rows to '${tableName}'...`);
            const addedRows = table.rows.add(null, data);

            // Load the address of the added rows to access them later for formatting
            // addedRows.load("address");

            await context.sync();
            console.log(`WriteToTable: Successfully added rows to '${tableName}'.`);

            // Example: Set number format for the first column of the added rows
            // const range = context.workbook.worksheets.getActiveWorksheet().getRange(addedRows.address);
            // range.numberFormat = [[null, "mm-dd-yyyy", null]]; // Assume the second column needs date formatting
            // Adjust the numberFormat array to match the formatting requirements of your table columns
            // await context.sync();
        });
    } catch (error) {
        console.error(`Error writing to table '${tableName}':`, error);
        console.error("Could not add data to table '" + tableName + "':", data);
        throw error; // Rethrow to ensure calling code is aware of the failure
    }
}

export async function UpdateTableRow(tableName: string, rowIndex: number, data: any) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(tableName);
        // Assuming row index is based on the data body range (not including the header)
        const rowRange = table.getDataBodyRange().getRow(rowIndex);
        // Convert the entry object to an array of values based on table headers
        const headers = table.getHeaderRowRange().load('values');
        await context.sync(); // Load headers

        const updatedValues = headers.values[0].map(header => data[header] ?? null);
        rowRange.values = [updatedValues];
        await context.sync();
    }).catch(error => {
        console.error("Error updating row in table:", error);
        throw error; // Rethrow to ensure calling code is aware of the failure
    });
}

export async function UpdateTableRows(tableName: string, updates: { rowIndex: number, data: any }[]) {
    if (updates.length === 0) return;
    console.log(`UpdateTableRows called for '${tableName}' with ${updates.length} updates.`);

    await Excel.run(async (context) => {
        console.log(`UpdateTableRows: Getting table '${tableName}'...`);
        const table = context.workbook.tables.getItem(tableName);
        const headers = table.getHeaderRowRange().load('values');
        await context.sync();

        const headerValues = headers.values[0];
        console.log(`UpdateTableRows: Processing updates for '${tableName}'...`);

        updates.forEach(update => {
            const rowRange = table.getDataBodyRange().getRow(update.rowIndex);
            const updatedValues = headerValues.map(header => update.data[header] ?? null);
            rowRange.values = [updatedValues];
        });

        console.log(`UpdateTableRows: Syncing updates for '${tableName}'...`);
        await context.sync();
        console.log(`UpdateTableRows: Successfully updated rows in '${tableName}'.`);
    }).catch(error => {
        console.error(`Error updating rows in table '${tableName}':`, error);
        throw error;
    });
}

export async function EnsureTableExists(tableName: string, columns: string[], sheetName: string = tableName) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItemOrNullObject(tableName);
        await context.sync();

        if (!table.isNullObject) {
            return; // Table exists
        }

        console.log(`Table '${tableName}' not found. Creating it...`);

        // Check if sheet exists
        let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        await context.sync();

        if (sheet.isNullObject) {
            console.log(`Sheet '${sheetName}' not found. Creating it...`);
            sheet = context.workbook.worksheets.add(sheetName);
        }

        // Create Table
        // Determine range based on columns count. e.g., A1:C1 for 3 columns.
        const endChar = String.fromCharCode('A'.charCodeAt(0) + columns.length - 1);
        const rangeAddress = `A1:${endChar}1`;
        
        const newTable = sheet.tables.add(rangeAddress, true);
        newTable.name = tableName;
        newTable.getHeaderRowRange().values = [columns];

        await context.sync();
        console.log(`Table '${tableName}' created successfully.`);
    }).catch(error => {
        console.error(`Error ensuring table '${tableName}' exists:`, error);
        throw error;
    });
}

export async function SetNamedRangeValue(rangeName: string, value: any) {
    console.log(`SetNamedRangeValue called for '${rangeName}' with value '${value}'.`);
    await Excel.run(async (context) => {
        try {
            console.log(`SetNamedRangeValue: Getting named item '${rangeName}'...`);
            let namedItem = context.workbook.names.getItemOrNullObject(rangeName);
            await context.sync();

            if (namedItem.isNullObject) {
                // Lazy creation logic for known ranges
                if (rangeName === 'LastRolloverUpdate') {
                    // Default to Rollovers!H1 if missing
                    console.warn(`Named range '${rangeName}' not found. Attempting to create it at Rollovers!H1.`);
                    namedItem = context.workbook.names.add(rangeName, 'Rollovers!H1');
                } else {
                    console.warn(`Named range '${rangeName}' not found and no default creation logic exists. Skipping.`);
                    return;
                }
            }

            console.log(`SetNamedRangeValue: Setting value for '${rangeName}'...`);
            const range = namedItem.getRange();
            range.values = [[value]]; // Range values must be 2D array
            await context.sync();
            console.log(`SetNamedRangeValue: Successfully set value for '${rangeName}'.`);
        } catch (innerError) {
            console.warn(`Warning: Could not set named range '${rangeName}'. This is non-critical. Error:`, innerError);
            // Do NOT re-throw. Allow the main process to continue.
        }
    }).catch(error => {
        // This catch block handles errors in Excel.run setup itself, which we still log but don't rethrow to avoid crashing caller
        console.warn(`Excel.run failed in SetNamedRangeValue for '${rangeName}':`, error);
    });
}

