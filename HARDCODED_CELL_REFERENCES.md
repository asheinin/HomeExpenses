# Hardcoded Cell References Audit

This document lists every place in the **HomeExpenses** repo where Google Sheets cells/rows/columns are addressed by literal numbers or A1 notation instead of by the constants defined in `staticNumbers` (in [src/static/main.js](src/static/main.js)).

## Audit rules

A line is flagged when it uses a literal cell reference that **could** be expressed with a `staticNumbers` field. The only allowed exceptions are:

- A literal **row** of `1`, `2`, or `3` (the very top of a sheet — header / metadata rows).
- A relative offset of `+1`, `+2`, or `+3` from a constant.

Hardcoded **columns** (even `1`, `2`, `3`) are **not** exceptions — `expenseTypeColumn`, `expenseDescrColumn`, `expenseInitialBalanceCol`, `summaryAnalyticsYearColumn`, etc. exist and should be used.

Constants reference (from [src/static/main.js](src/static/main.js)):

| Field | Value |
|---|---|
| `expenseFirstRow` | 3 |
| `expenseLastRow` | 50 |
| `expenseTypeColumn` | 1 |
| `expenseDescrColumn` | 2 |
| `expenseDateColumn` | 3 |
| `expenseAmountColumn` | 4 |
| `expenceSplitColumn` | 7 |
| `expenseFirstPayColumn` | 8 |
| `expenseSecondPayColumn` | 9 |
| `expencePeriodColumn` | 10 |
| `expensePaidColumn` | 11 |
| `expensePAPColumn` | 12 |
| `expenseSp1ToSp2Column` | 14 |
| `expenseInitialBalanceCol` | 2 |
| `summarySumRow` | 2 |
| `summaryAmountColumn` | 3 |
| `summaryAnalyticsYearColumn` | 1 |
| `summaryAnalyticsDataStartColumn` | 4 |
| `dashColumns` | 13 |
| `dashMonthNameColumn` | 1 |

---

## Violations

### A1-notation literals (`'A2:D50'`) — **FIXED**

| File | Line | Status |
|---|---|---|
| [src/CreateEOYDocument.js:95](src/CreateEOYDocument.js#L95) | 95 | Fixed — now `sheet.getRange(myNumbers.expenseCarryOverRow, myNumbers.expenseTypeColumn, myNumbers.expenseLastRow - myNumbers.expenseCarryOverRow + 1, myNumbers.expenseAmountColumn)` |
| [src/SummaryExpenses.js:55](src/SummaryExpenses.js#L55) | 55 | Fixed — same replacement |

---

### Hardcoded column `1` where `expenseTypeColumn` should be used

| File | Line | Code |
|---|---|---|
| [src/AddNewExpence.js:98](src/AddNewExpence.js#L98) | 98 | `sheet.getRange(row, 1, 1, numColsToClear).clearContent();` |
| [src/AddNewExpence.js:138](src/AddNewExpence.js#L138) | 138 | `sheet.getRange(myNumbers.expenseFirstRow, 1, numRows, myNumbers.expensePAPColumn).getValues();` |
| [src/AddNewExpence.js:257](src/AddNewExpence.js#L257) | 257 | `sheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseAmountColumn).getValues();` |
| [src/CleanMonths.js:40](src/CleanMonths.js#L40) | 40 | `targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, 4).clearContent();` (also `4` → `expenseAmountColumn`) |
| [src/CopyMonth.js:70](src/CopyMonth.js#L70) | 70 | `targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol).getA1Notation();` |
| [src/CopyMonth.js:94](src/CopyMonth.js#L94) | 94 | `sourceSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol).getValues();` |
| [src/CopyMonth.js:97](src/CopyMonth.js#L97) | 97 | `targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol);` |
| [src/CopyMonth.js:134](src/CopyMonth.js#L134) | 134 | `sourceSheet.getRange(sourceRow, 1, 1, lastCol)` |
| [src/CopyMonth.js:135](src/CopyMonth.js#L135) | 135 | `.copyTo(targetSheet.getRange(targetRow, 1));` |
| [src/CreateNewFile.js:62](src/CreateNewFile.js#L62) | 62 | `oldDecSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseSp1ToSp2Column).getValues();` |
| [src/CreateNewFile.js:71](src/CreateNewFile.js#L71) | 71 | `targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, 4).clearContent();` (also `4` → `expenseAmountColumn`) |
| [src/CreateNewFile.js:111](src/CreateNewFile.js#L111) | 111 | `summarySheetNext.getRange(myNumbers.summarySumRow + 1, 1, lastRowSummary - myNumbers.summarySumRow, summarySheetNext.getMaxColumns()).clearContent();` |
| [src/GetMonthlyBalance.js:47](src/GetMonthlyBalance.js#L47) | 47 | `sheet.getRange(myNumbers.dashFirstMonthRow, 1, 12, myNumbers.dashColumns);` (col `1` → `dashMonthNameColumn`) |
| [src/open.js:420](src/open.js#L420) | 420 | `var sourceRange = sheet.getRange(myNumbers.expenseFirstRow, 1, 1, sheet.getLastColumn());` |
| [src/open.js:421](src/open.js#L421) | 421 | `var targetRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());` |

---

### Hardcoded column `2` where a constant exists

| File | Line | Code | Suggested constant |
|---|---|---|---|
| [src/CleanMonths.js:49](src/CleanMonths.js#L49) | 49 | `targetSheet.getRange(myNumbers.expenseCarryOverRow, 2, 1, targetSheet.getMaxColumns() - 1).clearContent();` | `expenseInitialBalanceCol` (or `expenseDescrColumn`) |
| [src/CreateNewFile.js:80](src/CreateNewFile.js#L80) | 80 | `targetSheet.getRange(myNumbers.expenseCarryOverRow, 2, 1, targetSheet.getMaxColumns() - 1).clearContent();` | `expenseInitialBalanceCol` |

---

### Hardcoded column `7` where `expenceSplitColumn` should be used (and `6` for span) — **FIXED**

The span `6` corresponds to `expensePAPColumn (12) - expenceSplitColumn (7) + 1`.

| File | Line | Status |
|---|---|---|
| [src/CleanMonths.js:43](src/CleanMonths.js#L43) | 43 | Fixed — now `getRange(myNumbers.expenseFirstRow, myNumbers.expenceSplitColumn, numOfRows, myNumbers.expensePAPColumn - myNumbers.expenceSplitColumn + 1)` |
| [src/CreateNewFile.js:74](src/CreateNewFile.js#L74) | 74 | Fixed — same replacement |

---

### Hardcoded literal rows beyond the `1/2/3` exception

| File | Line | Code | Suggested constant |
|---|---|---|---|
| [src/Analytics.js:317](src/Analytics.js#L317) | 317 | `const dataRange = summarySheet.getRange(3, 1, lastRow - 2, 3).getValues();` | row `3` is borderline allowed (`expenseFirstRow`/header band), but the `lastRow - 2` and col `1`, span `3` should be `summaryAnalyticsYearColumn`, `summaryAmountColumn`, etc. The inline comment on lines 314-316 already calls this out. |

---

### Hardcoded column-span `25` with no constant

`Analytics.js` reads/clears 25 columns of the Summary sheet but no `staticNumbers` field describes that width. Either define one (e.g. `summaryAnalyticsTotalColumns`) or compute from `summaryAnalyticsDataStartColumn` + a tier count.

| File | Line | Code |
|---|---|---|
| [src/Analytics.js:126](src/Analytics.js#L126) | 126 | `summarySheet.getRange(1, 1, Math.max(lastRowOfData, 1), 25).getValues();` (also col `1` → `summaryAnalyticsYearColumn`; row `1` is allowed exception) |
| [src/Analytics.js:135](src/Analytics.js#L135) | 135 | `summarySheet.getRange(startRow, 1, Math.max(lastRowOfData - startRow + 30, 1), 25).clearContent();` (col `1`) |

> The literal `16` on [src/Analytics.js:132](src/Analytics.js#L132) (`existingHeaders[i][16]`) is also a magic index into the same 25-column row.

---

### Borderline (row `1` exception, but col is hardcoded)

These pass the row-`1` allowed exception, but the column is still a literal:

| File | Line | Code | Note |
|---|---|---|---|
| [src/SummaryExpenses.js:99](src/SummaryExpenses.js#L99) | 99 | `summarySheet.getRange(1, i + 1, lastRow, 1);` | row `1` OK; trailing `1` is column-span (acceptable). |
| [src/SummaryExpenses.js:111](src/SummaryExpenses.js#L111) | 111 | `summarySheet.getRange(1, myNumbers.expenseTypeColumn, lastRow, 1);` | OK (row `1` exception, span `1`). |
| [src/SummaryExpenses.js:112](src/SummaryExpenses.js#L112) | 112 | `summarySheet.getRange(1, myNumbers.expenseDescrColumn, lastRow, 1);` | OK. |
| [src/SummaryExpenses.js:113](src/SummaryExpenses.js#L113) | 113 | `summarySheet.getRange(1, myNumbers.summaryAmountColumn, lastRow, 1);` | OK. |

---

## Files audited — no violations found

- [src/Charts.js](src/Charts.js) — every `getRange` uses `myNumbers.*`.
- [src/CloseMonthDialog.js](src/CloseMonthDialog.js) — `spfrow`, `cocol`, `corow` are local variables, not literals.
- [src/Summary.js](src/Summary.js) — no `getRange` calls.
- [src/SendMail.js](src/SendMail.js) — no `getRange` calls.
- [src/NotifyNewFile.js](src/NotifyNewFile.js) — uses `myNumbers.*`.
- [src/utils/Utilities.js](src/utils/Utilities.js) — no `getRange` calls.
- [src/ai/ExpenseAnalysisAgent.js](src/ai/ExpenseAnalysisAgent.js) — uses `myNumbers.*`.
- [src/ai/MonthlyAnalyticsEmail.js](src/ai/MonthlyAnalyticsEmail.js) — uses `myNumbers.*`.
- [src/ai/Gemini.js](src/ai/Gemini.js) — no sheet access.
- [src/CreateEOYDocument.js:176](src/CreateEOYDocument.js#L176) and below — uses `myNumbers.*`.
- [ui/](ui/) HTML & JS — no sheet access.

---

## Summary

| Category | Count |
|---|---|
| A1 notation (`'A2:D50'`) | **2** |
| Hardcoded col `1` | **15** |
| Hardcoded col `2` | **2** |
| Hardcoded col `7` (with span `6`) | **2** |
| Hardcoded row beyond exception | **1** |
| Hardcoded col-span `25` (no constant exists) | **2** |
| **Total flagged** | **24** |

The cleanest single fix is `Analytics.js` (define `summaryAnalyticsTotalColumns` in `staticNumbers`). The most repetitive fix is replacing `, 1,` with `, myNumbers.expenseTypeColumn,` in the seven `CopyMonth.js` / `CreateNewFile.js` / `CleanMonths.js` / `AddNewExpence.js` / `open.js` callsites that scan the expense block.
