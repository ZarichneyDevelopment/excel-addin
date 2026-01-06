# GEMINI Maintenance Document

This document provides a comprehensive overview of the "Budget Helper" Excel add-in for future maintenance and development.

## Project Overview

"Budget Helper" is an Excel add-in designed for personal finance management. Its primary features are:

1.  **Transaction Importing:** Users can import transaction data from CSV files (specifically tailored for RBC bank format) into an Excel workbook using a drag-and-drop interface.
2.  **Auto-Categorization:** The add-in features a self-learning categorization engine. It uses a "MatchingRules" table in the workbook to automatically categorize expenses. New, uncategorized transactions are added to this table for the user to classify, which improves future imports.
3.  **Budget Rollover:** The add-in calculates and manages budget amounts that roll over from one month to the next.

The entire system is architected to work with a specific Excel template containing numerous tables (e.g., `Transactions`, `Accounts`, `MatchingRules`, `Rollovers`) and named ranges.

## Technologies

The project is built using the following technologies:

*   **TypeScript:** The primary language for the add-in's logic.
*   **Office.js:** The JavaScript API for interacting with Microsoft Office applications.
*   **RxJS:** Used for managing asynchronous data streams from the workbook.
*   **papaparse:** Used for parsing CSV files.
*   **Webpack:** Used for building and bundling the application.

## Project Structure

The project is organized into two main directories:

*   **`/` (root):** Contains stale build artifacts. This directory can be ignored for development purposes.
*   **`/budget-helper`:** The main directory containing the source code and configuration for the add-in.

### Key Files in `/budget-helper`

*   `manifest.xml`: The add-in's manifest file, which defines its settings and capabilities.
*   `package.json`: Lists the project's dependencies and build scripts.
*   `webpack.config.js`: The configuration file for Webpack, which defines the build process.
*   `src/`: The directory containing the add-in's source code.
    *   `taskpane/`: Contains the UI logic for the add-in's task pane.
        *   `taskpane.ts`: The main entry point for the user interface.
    *   `commands/`: Contains the logic for the add-in's commands.
    *   `transaction.ts`: Contains the core business logic for processing transactions.
    *   `rollover.ts`: Contains the logic for calculating budget rollovers.
    *   `lookups.ts`: The data access layer for reading from the workbook.
    *   `excel-helpers.ts`: A low-level wrapper for the Office.js API.

## Getting Started

To set up the development environment, you will need to have Node.js and npm installed.

1.  **Install dependencies:** Navigate to the `budget-helper` directory and run `npm install`.
2.  **Build the project:** Run `npm run build` to build the project.
3.  **Start the development server:** Run `npm start` to start the development server.
4.  **Sideload the add-in:** Follow the instructions in the [Office Add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) to sideload the add-in in Excel.

### LibreOffice and `recalc.py`

This project uses `recalc.py` to recalculate Excel formulas and check for errors, which is crucial for development and testing. `recalc.py` depends on LibreOffice, which is expected to be installed via Flatpak.

**Current Issue:** The `flatpak` command is not accessible in the current environment's PATH, which prevents the use of `recalc.py`. Before using `recalc.py`, please ensure that `flatpak` is correctly installed and accessible in the system's PATH. If `flatpak` is installed but not found, you may need to specify its absolute path or adjust the environment's PATH variable.

## Future Maintenance

Here are some notes and suggestions for future maintenance and development:

*   **Stale Build Artifacts:** The root directory contains stale build artifacts. These should be cleaned up to avoid confusion.
*   **RBC-Specific Code:** The CSV parsing logic is currently tailored for RBC bank format. This could be generalized to support other bank formats.
*   **Error Handling:** The error handling in the application could be improved to provide more informative feedback to the user.
*   **Testing:** Jest suites now cover `rollover`, `excel-helpers`, `lookups`, `transaction`, and `file-drop`. Run `npm test -- --coverage` from `budget-helper` to execute everything and generate coverage reports, and keep the generated `coverage` directory out of version control unless you intend to publish it.
*   **Documentation:** The code is well-structured, but adding more comments and documentation would be beneficial for future maintainers.

## Known Issues

### `tmp` Vulnerability

There is a known low-severity vulnerability in the `tmp` package, which is a deep dependency of the `office-addin-debugging` package. The vulnerability is tracked as [GHSA-52f5-9888-hmc6](https://github.com/advisories/GHSA-52f5-9888-hmc6).

Numerous attempts were made to fix this vulnerability, including:
*   Running `npm audit fix` and `npm audit fix --force`.
*   Manually updating all `office-addin-*` packages to their latest versions.
*   Deleting `node_modules` and `package-lock.json` and performing a clean install.

None of these steps have resolved the issue. Since this is a low-severity vulnerability in a development dependency, it is not expected to have an impact on the production build of the add-in. This issue is documented here for future reference.

## Budget.xlsx Structure

The `Budget.xlsx` file is the heart of the "Budget Helper" add-in. It is a comprehensive, zero-based budget spreadsheet for Steven and Mel's personal finance management.

### Overview

This is Steven and Mel's way to organize their life finances. It is intended as a zero-based budget, easy with known fixed income.

### Sheets

The workbook contains the following sheets, each with a specific purpose:

*   **Income:** One row for Steve, one for Mel, one for total. The add-in primarily uses the monthly/yearly total from here.
*   **Expenses:** A list of monthly expenses, also known as "budgeting buckets". Each has rollovers, so for variable expenses, the EOM balance is reviewed to alter real-life spending according to the health of each bucket.
*   **Monthly Overview:** A dashboard that takes data from the `Income` and `Expenses` sheets for analytical purposes.
*   **Yearly Overview:** Same as the `Monthly Overview`, but for a 12-month span.
*   **Budget:** The main view of the spreadsheet, controlled by a month and year input which generates the calculations.
    *   Columns C, I, and O are for chequing and savings accounts. These are drop-downs that feed from the `Expenses` tab.
    *   The `Rollover` columns look at the `Rollovers` tab for the given entry value based on month/year/category.
    *   The `Budgeted` columns are first fed from the `ChangeHistory` tab, and if no matching entry exists, the value from the `Expenses` tab is used.
    *   The `Spent` columns are a summation of the matching category/date of rows from the `Transactions` tab.
    *   The `Remaining` columns are a simple calculation of the prior three numbers, showing the EOM amount, which is then stamped into the `Rollovers` tab to be pulled in the future.
*   **Transactions:** Holds all RBC transactions. Once a month, CSVs from three accounts (shared visa, chequing, and savings) are dropped into the add-in, which adds the entries to this tab.
*   **CategoryRules:** Used by the add-in for auto-categorization of transactions.
*   **Accruals:** A special sub-budgeting system for large, one-time purchases. The finish date is when the accumulated money is ready to be spent. This allows for flexibility with leftover money after all monthly expenses are taken care of.
*   **Rollovers:** Stores the end-of-month (EOM) rollover amounts for each budget bucket.
*   **ChangeHistory:** Tracks changes to budgeted amounts.
*   **Accounts:** A list of bank accounts.
*   **Feature Idea:** A place to note ideas for new features.

### Named Ranges

The workbook uses the following named ranges to easily access data from the add-in:

*   **Accruals:** `Accruals!$A$1:$L$13`
*   **AccrualsMonthly:** `Accruals!$G$14`
*   **AccruedTotal:** `Accruals!$F$14`
*   **Categories:** `Expenses!$C:$C`
*   **CurrentMonth:** `Budget!$A$2`
*   **CurrentYear:** `Budget!$B$2`
*   **Deductions:** `Income!$I:$I`
*   **ExpenseAccounts:** `Expenses!$E:$E`
*   **ExpenseBudgets:** `Expenses!$B:$B`
*   **Expenses:** `Expenses!$A:$A`
*   **Gross:** `Income!$D:$D`
*   **Monthly:** `Income!$J:$J`
*   **MonthlyLeftover:** `'Monthly Overview'!$B$12`
*   **Requiredness:** `Expenses!$D:$D`
*   **Source:** `Income!$B:$B`
*   **TotalAccrued:** `Accruals!$F$14`
*   **TransactionDates:** `Transactions[Date]`
*   **TransactionExpenseType:** `Transactions[Expense Type]`
*   **TransactionIds:** `Transactions[ID]`
*   **TransactionsAmount:** `Transactions[Amount]`
*   **Yearly:** `Income!$K:$K`

### Tables

The workbook also uses structured tables to organize data:

*   **ExpenseData:** Referenced by the `TotalMonthlyExpenses` named range.
*   **IncomeData:** Referenced by the `TotalMonthlyIncome` and `TotalSalary` named ranges.
*   **Transactions:** The `Transactions` sheet is a table, as indicated by the structured references in the `Transaction*` named ranges (e.g., `Transactions[Date]`).
