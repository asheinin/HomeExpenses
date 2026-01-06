# Home Expenses

A comprehensive Google Apps Script project for managing and tracking home expenses shared between two individuals (Spouses/Partners). It provides a structured way to record monthly expenses, handle shared payments, and generate year-to-date (YTD) summaries and tax-ready documents.

## Key Features

- **Custom Sidebar & Dialogs**: Uses a custom UI for adding expenses, choosing between recurrent or one-time entries.
- **AI-Assisted Entry**: Includes logic for AI-powered expense addition (integration with form-based entries).
- **Monthly Balance Tracking**: Specifically designed for two users (Spouse 1 & Spouse 2) with automated calculations of who owes what.
- **Flexible Month Settlement**: Interactive options to mark months as fully paid, carry over balances, or make partial payments.
- **YTD Summary & Historical Analytics**: Automatically aggregates expenses by category across months and years, providing cross-file historical spending trends.
- **Automated Notifications**: Sends monthly balance summaries and document readiness alerts via email.
- **End-of-Year (EOY) Tools**: Generates next year's spreadsheet and tax receipt documents.
- **Pre-Authorized Payment (PAP) Support**: Tracks and flags automated payments.

## Project Structure

- `main.js`: Contains global configurations, constants for spreadsheet layout, and constructor functions for static numbers.
- `open.js`: Handles the `onOpen` trigger, creates custom menus (`Payments`, `Settle Month`, `Expenses`, `General Actions`), and manages real-time cell validation and formula updates.
- `AddNewExpence.js`: Original logic for adding recurrent or one-time expenses with basic dialog prompts.
- `ExpenseForm.html`: HTML form for adding expenses with fields for name, type, amount, PAP status, billing period, split, paid status, and who paid.
- `SummaryAI.js`: Logic for generating the "Summary" sheet, aggregating data by type, and creating charts.
- `Analytics.js`: Cross-file historical spending data aggregation and multi-year chart generation.
- `CloseMonthDialog.js`: Interactive dialogs and logic for settling monthly balances (Paid, Carry Over, Partial).
- `GetMonthlyBalance.js`: Calculates the final balance between users and sends automated email summaries.
- `CreateEOYDocument.js` & `CreateNewFile.js`: Administrative tools for wrapping up the year and preparing for the next.
- `NotifyNewFile.js`: System for notifying users when documents are ready via email and UI toasts.
- `SendMail.js`: Internal utility for sending formatted emails.
- `Utilities.js`: Helper functions for custom UI dialogs (Yes/No, Input prompts).
- `BarCharts.html` & `Chart.js`: Components for visualizing expense data.

## Setup Instructions

1.  **Google Sheet Setup**: Create a new Google Sheet.
2.  **Open Apps Script**: Go to `Extensions` > `Apps Script`.
3.  **Copy Files**: Copy all `.js` and `.html` files from this repository into the Apps Script editor.
4.  **Rename Files**: Change `.js` extensions to `.gs` within the Apps Script editor (if copying manually).
5.  **Initialize**: Run the `open` function once from the script editor to set up triggers and custom menus.
6.  **Dashboard Configuration**: Ensure your first sheet is your "Dashboard" where User 1 and User 2 names and emails are set (see `main.js` for layout constants).

## Usage

Once installed, you will see a **Payments** and **Expenses** menu in your Google Sheet:
- **Expenses**: Use "Create/Update Expense" to add new items. It will automatically populate the relevant months.
- **Settle Month**: Use this to close the previous month with options for full payment, carry over, or partial amounts.
- **General Actions**: 
    - **Summarize Expenses**: Refresh your Summary sheet and annual chart.
    - **Run Analytics**: Aggregate data from all historical files and generate a multi-year spending analysis.

## License

MIT
