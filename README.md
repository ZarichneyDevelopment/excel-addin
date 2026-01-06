# Budget Helper Excel Add-in

**Product Owner:** Steven & Mel  
**Current Version:** 1.0 (Refactoring to 2.0 in progress)

## 1. Project Purpose
The "Budget Helper" is a custom Excel add-in designed to automate and streamline personal finance management for Steven and Mel. It interfaces directly with a specific zero-based budgeting spreadsheet (`Budget.xlsx`) to eliminate manual data entry and complex categorization tasks.

## 2. Core Requirements
- **Zero-Based Budgeting:** Every dollar of income is allocated to expenses, savings, or accruals.
- **Automated Transaction Import:** Import CSV files from RBC accounts (Shared Visa, Chequing, Savings) directly into the workbook.
- **Smart Categorization:** Automatically categorize transactions based on a self-learning "MatchingRules" table. Uncategorized items are flagged for user review.
- **Rollover Management:** Track "budgeting buckets" where unused funds roll over to the next month (e.g., Groceries, Entertainment).
- **Accrual Tracking:** Manage savings for large, one-time purchases with specific target dates.

## 3. Current State & Roadmap
The project is currently in a functional but "quick-and-dirty" state (v1.0). 

**Current Pain Points:**
- **UI/UX:** Confusing control behaviors and implementation rules.
- **Performance:** Rollover reprocessing is becoming slow over time.
- **Reliability:** Auto-categorization is inconsistent.

**v2.0 Refactoring Goals:**
- **Testing:** Implement a comprehensive test suite (Jest) to prevent regressions (In Progress).
- **Performance:** Optimize the `rollover.ts` logic to handle growing data efficiently.
- **Robustness:** Refactor `transaction.ts` to improve categorization accuracy.
- **Usability:** Redesign the task pane UI for better clarity and workflow.

## 4. How It Works (The "Secret Sauce")

The add-in acts as the bridge between raw bank data and the structured `Budget.xlsx` spreadsheet.

### The Spreadsheet (`Budget.xlsx`)
This file is the "database" and "dashboard" of the system.
- **`Budget` Tab:** The main interface. Controlled by Month/Year inputs. It pulls "Budgeted" amounts from `ChangeHistory` (or defaults to `Expenses`), sums actuals from `Transactions`, and calculates "Remaining" values which are stamped into `Rollovers`.
- **`Transactions` Tab:** The central ledger. All imported bank data lands here.
- **`Rollovers` Tab:** A historical record of EOM balances for every category, allowing the budget to "remember" past performance.
- **`CategoryRules` Tab:** The brain of the auto-categorizer. Maps text patterns (e.g., "Starbucks") to Categories (e.g., "Dining").
- **`Accruals` Tab:** A sub-system for sinking funds. Calculates monthly savings required to hit a future target amount.

### The Code
- **`transaction.ts`:** Parses CSVs, deduplicates entries, and applies regex-based rules from `CategoryRules` to assign categories.
- **`rollover.ts`:** The heavy lifter. It reads monthly totals, budget limits, and previous rollover states to compute the new End-of-Month (EOM) balance for every bucket. This is currently the performance bottleneck.
- **`taskpane.ts`:** The UI layer that triggers these actions.

## 5. Maintenance & Setup
For detailed developer setup, dependencies (Node.js, LibreOffice, Flatpak), and known issues, please refer to [GEMINI.md](GEMINI.md).
