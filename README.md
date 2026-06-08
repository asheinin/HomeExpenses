# Home Expenses

A comprehensive Google Apps Script project for managing and tracking home expenses shared between two individuals (Spouses/Partners). It provides a structured way to record monthly expenses, handle shared payments, and generate year-to-date (YTD) summaries and tax-ready documents.

## Key Features

- **Custom Sidebar & Dialogs**: Uses a custom UI for adding expenses, choosing between recurrent or one-time entries.
- **AI-Assisted Entry**: Includes logic for AI-powered expense addition (integration with form-based entries).
- **Monthly Balance Tracking**: Specifically designed for two users (Spouse 1 & Spouse 2) with automated calculations of who owes what.
- **Flexible Month Settlement**: Interactive options to mark months as fully paid, carry over balances, or make partial payments.
- **YTD Summary & Historical Analytics**: Automatically aggregates expenses by category across months and years, providing cross-file historical spending trends with sparkline visualizations.
- **Automated Notifications**: Sends monthly balance summaries and document readiness alerts via email.
- **End-of-Year (EOY) Tools**: Generates next year's spreadsheet and tax receipt documents.
- **Pre-Authorized Payment (PAP) Support**: Tracks and flags automated payments.

---

## Project Structure

```
HomeExpenses/
├── appsscript.json          # Apps Script manifest (timezone, runtime settings)
├── .clasp.json              # CLASP configuration for local development
├── README.md                # This file
├── src/                     # Core application logic
│   ├── static/
│   │   └── main.js          # Global configs, constants, spreadsheet layout definitions
│   ├── utils/
│   │   └── Utilities.js     # Helper functions for custom UI dialogs (Yes/No, Input prompts)
│   ├── open.js              # onOpen trigger, custom menus, cell validation, formula updates
│   ├── AddNewExpence.js     # Logic for adding recurrent or one-time expenses
│   ├── SummaryAI.js         # Summary sheet generation, data aggregation, chart creation
│   ├── Analytics.js         # Cross-file historical spending aggregation, multi-year charts
│   ├── CloseMonthDialog.js  # Dialogs and logic for settling monthly balances
│   ├── GetMonthlyBalance.js # Balance calculations and automated email summaries
│   ├── CreateEOYDocument.js # End-of-year document generation
│   ├── CreateNewFile.js     # New year spreadsheet creation
│   ├── NotifyNewFile.js     # Email and toast notifications for document readiness
│   ├── SendMail.js          # Internal email utility
│   ├── Chart.js             # Chart visualization components
│   ├── CleanMonths.js       # Month data cleanup utilities
│   └── CopyMonth.js         # Month data duplication logic
└── ui/                      # HTML templates for user interface
    ├── ExpenseForm.html     # Form for adding expenses (name, type, amount, PAP, split, etc.)
    ├── DeleteExpenseForm.html # Form for removing expenses
    └── BarCharts.html       # Chart visualization template
```

### Core Files Explained

| File | Purpose |
|------|---------|
| `main.js` | Defines all spreadsheet layout constants (row/column positions), threshold limits, and color schemes |
| `open.js` | Entry point that creates custom menus (`Payments`, `Settle Month`, `Expenses`, `General Actions`) and handles cell edit triggers |
| `Analytics.js` | Aggregates data across multiple yearly files to generate historical spending trends with stacked column charts |
| `SummaryAI.js` | Creates the "Summary" sheet with category aggregations and annual charts |
| `CloseMonthDialog.js` | Provides interactive dialogs for month settlement (Paid, Carry Over, Partial) |

---

## Spreadsheet Layout

The project expects a specific spreadsheet structure:

### Dashboard Sheet
| Row | Content |
|-----|---------|
| 1 | Address |
| 2 | Names (Spouse 1 in Col B, Spouse 2 in Col C) |
| 3 | Email addresses |
| 4 | Split percentages |
| 5 | Current balances |
| 6 | Title row |
| 7+ | Monthly expense summaries |

### Monthly Expense Sheets
| Column | Content |
|--------|---------|
| A | Expense Type |
| B | Description |
| C | Date |
| D | Amount |
| E-F | Split amounts (Sp2/Sp1) |
| G | Split indicator |
| H-I | Payment columns |
| J | Billing Period |
| K | Paid status |
| L | PAP (Pre-Authorized Payment) |
| M-N | Transfers between spouses |

---

## Prerequisites

- Google Account with access to Google Sheets
- [CLASP](https://github.com/google/clasp) (for local development, optional)

---

## Setup Instructions

### Option 1: Manual Setup (Copy/Paste)

1. **Create a Google Sheet**: Start a new spreadsheet in Google Drive.
2. **Open Apps Script**: Navigate to `Extensions` > `Apps Script`.
3. **Copy Files**: Copy all `.js` and `.html` files from this repository into the Apps Script editor.
4. **Rename Extensions**: Change `.js` extensions to `.gs` within the Apps Script editor.
5. **Initialize**: Run the `open` function once from the script editor to set up triggers and custom menus.
6. **Configure Dashboard**: Set up your first sheet as "Dashboard" with:
   - Row 1: Your home address
   - Row 2: Spouse/Partner names (Col B & C)
   - Row 3: Email addresses (Col B & C)
   - Row 4: Split percentages (e.g., 50/50)

### Option 2: CLASP Deployment (Recommended for Development)

1. **Install CLASP globally**:
   ```bash
   npm install -g @google/clasp
   ```

2. **Login to CLASP**:
   ```bash
   clasp login
   ```

3. **Clone this repository and push**:
   ```bash
   git clone <repository-url>
   cd HomeExpenses
   clasp push
   ```

4. **Configure Dashboard** as described above.

---

## Usage

Once installed, you will see custom menus in your Google Sheet:

### Expenses Menu
- **Create/Update Expense**: Opens a form to add new expenses with fields for:
  - Name & Type
  - Amount
  - PAP status
  - Billing period
  - Split (Y/N)
  - Paid status
  - Who Paid

### Payments Menu
- View and manage payment records between spouses

### Settle Month Menu
- **Full Payment**: Mark month as completely settled
- **Carry Over**: Move unpaid balance to next month
- **Partial Payment**: Record partial settlement amounts

### General Actions Menu
- **Summarize Expenses**: Refresh the Summary sheet and generate annual charts
- **Run Analytics**: Aggregate data from all historical files and generate multi-year spending analysis with:
  - Historical spending data tables
  - Sparkline visualizations
  - Stacked column charts showing top categories per year

---

## Configuration

Key configuration values are defined in `src/static/main.js`:

```javascript
// Global settings
MAILER = 'HomePayments';        // Email sender name
FILENAME = 'Home payments';     // Base filename for spreadsheets

// Threshold for closing month
thresholdLimitForClosingMonth = 1;

// Color scheme for balance indicators
dashBalanceNegativeBgColor = "red";
dashBalancePositiveBgColor = "green";
```

---

## Runtime Environment

| Setting | Value |
|---------|-------|
| Runtime | V8 |
| Timezone | America/New_York |
| Exception Logging | Stackdriver |
| Web App Access | MYSELF |
| Execute As | USER_DEPLOYING |

---

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## License

MIT
