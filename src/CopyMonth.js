function copyMonthOT() { // One time: current to next month
    copyMonthAI('ot');
}

function copyMonthRM() { // Remaining: current to all remaining
    copyMonthAI('rm');
}

//Bulk Copy group of functions *************************************************************************/

function copyMonthAI(mode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var date = new Date();
    var currentMonth = date.getMonth() + 1; // 1-12

    if (currentMonth == 12) {
        SpreadsheetApp.getUi().alert("Cannot copy from December to next year/month via this function.");
        return;
    }

    var sourceSheet = ss.getSheets()[currentMonth]; // index 1 is Jan (currentMonth 1)
    var targetRange;
    if (mode == 'ot') {
        targetRange = [currentMonth + 1];
    } else {
        targetRange = [];
        for (var i = currentMonth + 1; i <= 12; i++) {
            targetRange.push(i);
        }
    }

    var copiedCount = 0;
    var skippedCount = 0;
    var myNumbers = new staticNumbers();
    var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;

    for (var i = 0; i < targetRange.length; i++) {
        var targetSheetIndex = targetRange[i];
        var targetSheet = ss.getSheets()[targetSheetIndex];

        // Check if target sheet has any expenses (no types and no names)
        var expenseRange = targetSheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, numOfRows, 2).getValues();
        var hasExpenses = expenseRange.some(function (row) {
            var type = row[0] ? row[0].toString().trim() : "";
            var name = row[1] ? row[1].toString().trim() : "";
            return type !== "" || name !== "";
        });

        if (!hasExpenses) {
            copiedCount += copyExpensesBetweenSheets(sourceSheet, targetSheet);
        } else {
            skippedCount++;
        }
    }

    var message = "Finished copying expenses. Total operations: " + copiedCount;
    if (skippedCount > 0) {
        message += ". Skipped " + skippedCount + " sheet(s) because they already have expenses.";
    }
    ss.toast(message, "Success", 5);
}

function copyExpensesBetweenSheets(sourceSheet, targetSheet) {
    var myNumbers = new staticNumbers();
    var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var lastCol = myNumbers.expenseSp1ToSp2Column; // Column 14

    // Copy conditional formatting rules that intersect with the expense range
    var rules = sourceSheet.getConditionalFormatRules();
    var expenseRangeA1 = targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol).getA1Notation();
    var newRules = [];

    rules.forEach(function (rule) {
        var ranges = rule.getRanges();
        var relevantRanges = ranges.filter(function (r) {
            // Check if rule range intersects with expense range (3-50, 1-14)
            return r.getRow() <= myNumbers.expenseLastRow &&
                (r.getRow() + r.getNumRows() - 1) >= myNumbers.expenseFirstRow &&
                r.getColumn() <= lastCol;
        });

        if (relevantRanges.length > 0) {
            var builder = rule.copy();
            var newRanges = relevantRanges.map(function (r) {
                return targetSheet.getRange(r.getA1Notation());
            });
            builder.setRanges(newRanges);
            newRules.push(builder.build());
        }
    });
    targetSheet.setConditionalFormatRules(newRules);

    // Get all data from source sheet to check Period and for names
    var sourceData = sourceSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol).getValues();

    // Get current target state to find empty rows or avoid duplicates
    var targetRange = targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, lastCol);
    var targetData = targetRange.getValues();
    var targetNames = targetData.map(function (row) { return row[myNumbers.expenseDescrColumn - 1]; });

    var count = 0;

    for (var i = 0; i < numOfRows; i++) {
        var period = sourceData[i][myNumbers.expencePeriodColumn - 1];
        var expenseName = sourceData[i][myNumbers.expenseDescrColumn - 1];

        // Only copy if Period is not empty and Name is not empty
        if (period && expenseName) {
            var existingIndex = targetNames.indexOf(expenseName);
            var targetRowIndex = -1;

            if (existingIndex !== -1) {
                // Update existing expense in target
                targetRowIndex = existingIndex;
            } else {
                // Find first empty row in target sheet
                for (var j = 0; j < numOfRows; j++) {
                    // Check if row is empty (no description and no type)
                    if (!targetData[j][myNumbers.expenseDescrColumn - 1] && !targetData[j][myNumbers.expenseTypeColumn - 1]) {
                        targetRowIndex = j;
                        // Mark as taken in our local cache
                        targetData[j][myNumbers.expenseDescrColumn - 1] = expenseName;
                        targetNames[j] = expenseName;
                        break;
                    }
                }
            }

            if (targetRowIndex !== -1) {
                var sourceRow = i + myNumbers.expenseFirstRow;
                var targetRow = targetRowIndex + myNumbers.expenseFirstRow;

                // Copy the whole range (1 to 14) to preserve formulas and formatting
                sourceSheet.getRange(sourceRow, 1, 1, lastCol)
                    .copyTo(targetSheet.getRange(targetRow, 1));

                count++;
            }
        }
    }
    validateType(targetSheet);
    return count;
}


