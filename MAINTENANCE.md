# Budget Helper Maintenance Guide

This document is the "Owner's Manual" for the Budget Helper codebase. It is designed to help you (or an AI agent) quickly understand the system's architecture, run tests, and safely implement changes.

## 1. System Architecture

The add-in follows a unidirectional data flow pattern where possible, heavily relying on RxJS for async data handling.

### Core Modules
*   **`src/taskpane/taskpane.ts`**: The UI Controller. It handles DOM events (clicks, drops) and orchestrates calls to the business logic. It does *not* contain business logic itself.
*   **`src/transaction.ts`**: The Parsing Engine. Responsibilities:
    *   Parses raw CSV strings using `papaparse`.
    *   Generates unique IDs (SHA-256 hash of transaction fields) to prevent duplicates.
    *   **Auto-Categorization:** Uses a 3-step process:
        1.  **Ambiguity Check:** Checks `AmbiguousItems` table. If matched, stops.
        2.  **Exact Match:** Checks `MatchingRules` for exact description matches.
        3.  **Aggressive Match:** Checks if any rule keyword is *contained* in the description.
*   **`src/rollover.ts`**: The Calculation Engine.
    *   **`resetRollover`**: The most complex function. It performs a "bulk fetch, in-memory calc, batch write" loop to ensure performance.
    *   **Optimization:** It loads all transactions and budget history into `Map` structures to avoid O(N) Excel reads inside loops.
*   **`src/lookups.ts`**: The Data Access Layer (DAL).
    *   Abstracts all "Read" operations from Excel tables.
    *   Returns strongly-typed Promises (e.g., `Promise<Transaction[]>`).
*   **`src/excel-helpers.ts`**: The Low-Level Interface.
    *   Directly interacts with the `Excel.run` context.
    *   Handles "Write" operations (`WriteToTable`, `UpdateTableRows`).
    *   **Lazy Creation:** `SetNamedRangeValue` automatically creates missing ranges like `LastRolloverUpdate`.

## 2. Development Workflow

### Prerequisites
*   **Node.js**: Required for build and test.
*   **LibreOffice (Flatpak)**: Required for the `skills/xlsx/recalc.py` script if you need to validate spreadsheet formulas outside of Excel.

### Running Tests
We use **Jest** for testing. The suite mocks the global `Excel` object and the `lookups` module to run fast and isolated tests.

```bash
cd budget-helper
npm test
```

**Key Test Files:**
*   `rollover.test.ts`: Covers the complex recalculation logic.
*   `transaction.test.ts`: Verifies categorization rules (including the "Walmart" ambiguity test).
*   `integration.test.ts`: Simulates a full end-to-end run (File Drop -> Process -> Recalc).

### Making Changes

#### How to Add a New Transaction Column
1.  **Excel:** Add the column to the `Transactions` table in `Budget.xlsx`.
2.  **Code (`transaction.ts`):** Update the `Transaction` class property list.
3.  **Code (`transaction.ts`):** Update the `ProcessTransactions` return array mapping to include the new field.
4.  **Test:** Update `integration.test.ts` to expect the new field in the mocked DB.

#### How to Modify Categorization Logic
1.  **Edit `transaction.ts`:** Look for the `ProcessTransactions` function.
2.  **Logic:** The current priority is: `Ambiguous Check` > `Exact Match` > `Aggressive (Includes) Match`.
3.  **Verify:** Run `npm test` to ensure `transaction.test.ts` still passes (especially the "Ambiguous" test case).

#### How to Debug Performance
If `resetRollover` becomes slow:
1.  Check `rollover.ts`.
2.  Ensure you are **NOT** calling `await` inside the `while` loop for Excel reads.
3.  All data should be pre-fetched into `Map` objects before the loop starts.

## 3. Known "Gotchas"

*   **Date Handling:** Excel dates are tricky. We generally use 1-based months (1=January) to match Excel's expected input for formulas.
*   **Mocking:** `rollover.ts` is tricky to test because `resetRollover` calls internal functions (`getRollover`). The tests use a specific `jest.mock` pattern to intercept these internal calls. **Do not remove the `jest.requireActual` logic in `rollover.test.ts`.**
*   **Environment:** The integration test uses `jsdom` to mock the browser environment (FileReader, DOM).

## 4. Deployment Strategy (Pre-Flight)

Before deploying to production (or handing off to the user):
1.  **Run Tests:** `npm test` (Must be green).
2.  **Build:** `npm run build` (Must complete without webpack errors).
3.  **Clean:** Remove `dist/` artifacts if needed.
