function open() {

  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var ui = SpreadsheetApp.getUi();

  var date = new Date();
  var currentMonth = date.getMonth();

  var currYear = date.getFullYear();

  var fileName = ss.getName();

  var fileYear = fileName.split(" ").slice(-1).pop();

  Logger.log(currYear + " " + " " + fileYear);

  //var currentMonth = 1;

  // Check if this is current year file
  if (currYear > fileYear) {
    currentMonth = 12;
  }

  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var formattedMonthCurrent = currentMonth < 12 ? months[currentMonth] : "";
  var formattedMonthPast = currentMonth < 12 ? months[(currentMonth - 1 + 12) % 12] : months[11];

  var specFundValueSp1 = sheet.getRange(myNumbers.dashBalancesRow, myNumbers.dashSpouse1NameColumn).getValue();
  var specFundValueSp2 = sheet.getRange(myNumbers.dashBalancesRow, myNumbers.dashSpouse2NameColumn).getValue();

  var specFund = ((specFundValueSp1 == 0) && (specFundValueSp2 == 0)) ? false : true;

  //var addNewExpenseRY = [{name: 'addNewExpenseRY', functionName: 'addNewExpense("ry")'}]; 
  //var addNewExpenseRM = [{name: 'addNewExpenseRM', functionName: 'addNewExpense("rm")'}];
  //var addNewExpenseOT = [{name: 'addNewExpenseOT', functionName: 'addNewExpense("ot")'}];


  //Check if this is current year file

  console.log("month: " + currentMonth);

  switch (currentMonth) {
    case 0:

      if (specFund) {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent")
            .addItem("Post from Initial Balance", "payMonthFromBalanceCurrent"))
          .addToUi();
      } else {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent"))
          .addToUi();
      }

      ui.createMenu('Expenses')
        .addItem("Create/Update One Time Expense in " + formattedMonthCurrent, "addNewExpenseOT")
        .addSeparator()
        .addSubMenu(ui.createMenu('Create/Update Recurrent Expense')
          .addItem("Create/Update Expense from January", "addNewExpenseRY"))
        .addSeparator()
        .addSubMenu(ui.createMenu('Delete Recurrent Expense')
          .addItem("Delete Expense from January", "deleteExpenseRY"))
        .addToUi();

      ui.createMenu('Bulk Actions')
        .addItem("Copy " + formattedMonthCurrent + " to Next Month", "copyMonthOT")
        .addItem("Copy " + formattedMonthCurrent + " to remanining months", "copyMonthRM")
        .addSeparator()
        .addItem("Clean after " + formattedMonthCurrent + " all remaining months", "cleanMonthsRM")
        .addToUi();

      break;

    case 11:

      if (specFund) {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly")
            .addItem("Post from Initial Balance", "payMonthFromBalance"))
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent")
            .addItem("Post from Initial Balance", "payMonthFromBalanceCurrent"))
          .addToUi();
      } else {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly"))
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent"))
          .addToUi();
      }

      ui.createMenu('Settle ' + formattedMonthPast)
        .addItem("Paid In Full", "closeMonthPaid")
        .addItem("Balance Carry Over", "closeMonthCarryOver")
        .addToUi();

      ui.createMenu('Expenses')
        .addItem("Create/Update One Time Expense in " + formattedMonthCurrent, "addNewExpenseOT")
        .addSeparator()
        .addSubMenu(ui.createMenu('Create/Update Recurrent Expense')
          .addItem("Create/Update Expense from January", "addNewExpenseRY"))
        .addSeparator()
        .addSubMenu(ui.createMenu('Delete Recurrent Expense')
          .addItem("Delete Expense from January", "deleteExpenseRY"))
        .addToUi();

      break;

    case 12:

      if (specFund) {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly")
            .addItem("Post from Initial Balance", "payMonthFromBalance"))
          .addToUi();
      } else {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly"))
          .addToUi();
      }

      ui.createMenu('Settle ' + formattedMonthPast)
        .addItem("Paid In Full", "closeMonthPaid")
        .addToUi();

      break;

    default:

      if (specFund) {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly")
            .addItem("Post from Initial Balance", "payMonthFromBalance"))
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent")
            .addItem("Post from Initial Balance", "payMonthFromBalanceCurrent"))
          .addToUi();
      } else {
        ui.createMenu('Payments')
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthPast)
            .addItem("Post Any Amount", "payMonthPartly"))
          .addSubMenu(ui.createMenu('Pay in ' + formattedMonthCurrent)
            .addItem("Post Any Amount", "payMonthPartlyCurrent"))
          .addToUi();
      }

      ui.createMenu('Settle ' + formattedMonthPast)
        .addItem("Paid In Full", "closeMonthPaid")
        .addItem("Balance Carry Over", "closeMonthCarryOver")
        .addToUi();

      ui.createMenu('Expenses')
        .addItem("Create/Update One Time Expense in " + formattedMonthCurrent, "addNewExpenseOT")
        .addSeparator()
        .addSubMenu(ui.createMenu('Create/Update Recurrent Expense')
          .addItem("Create/Update Expense from " + formattedMonthCurrent, "addNewExpenseRM")
          .addItem("Create/Update Expense from January", "addNewExpenseRY"))
        .addSeparator()
        .addSubMenu(ui.createMenu('Delete Recurrent Expense')
          .addItem("Delete Expense from " + formattedMonthCurrent, "deleteExpenseRM")
          .addItem("Delete Expense from January", "deleteExpenseRY"))
        .addToUi();

      ui.createMenu('Bulk Actions')
        .addSeparator()
        .addItem("Copy " + formattedMonthCurrent + " to Next Month", "copyMonthOT")
        .addItem("Copy " + formattedMonthCurrent + " to remanining months", "copyMonthRM")
        .addSeparator()
        .addItem("Clean after " + formattedMonthCurrent + " all remaining months", "cleanMonthsRM")
        .addToUi();


  }

  ui.createMenu('AI playground')
    .addItem("Run Expense Analysis Agent", "runExpenseAnalysisAgent")
    .addSeparator()
    .addItem("Set Gemini API Key", "setGeminiApiKey")
    .addToUi();


  ui.createMenu('General Actions')
    .addItem("Rebalance YTD", "rebalanceExpenses")
    .addSeparator()
    .addItem("Calculate YTD Totals", "Summary")
    .addSeparator()
    .addItem("Create Next Year File", "createNewFile")
    .addItem("Tax Receipt", "createEOYDocument")
    .addSeparator()
    .addItem("Change Split From Month...", "showSplitDialog")
    .addItem("Initialize Split Config", "migrateSplitConfig")
    .addToUi();


  // Set the data validation to require text in the form of an email address.
  var emailSp1Cell = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse1NameColumn);
  var emailSp2Cell = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse2NameColumn);
  var rule = SpreadsheetApp.newDataValidation().requireTextIsEmail().build();
  var a = emailSp1Cell.setDataValidation(rule);
  var b = emailSp2Cell.setDataValidation(rule);


  //set triggers if not set
  if (ScriptApp.getProjectTriggers().length < 5) {
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }

    ScriptApp.newTrigger('findAmount')
      .timeBased()
      .onMonthDay(5)
      .atHour(8)
      .create();

    ScriptApp.newTrigger('sendMonthlySummaryEmail')
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();

    ScriptApp.newTrigger('open')
      .forSpreadsheet(ss)
      .onOpen()
      .create();

    ScriptApp.newTrigger('edit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    ScriptApp.newTrigger('Summary')
      .timeBased()
      .everyDays(1) // Specifies a daily recurrence
      .atHour(0)    // Sets the trigger to run within the 12 AM to 1 AM hour range (0 in 24-hour time)
      .create();

  }
}


function edit() {
  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var col = activeCell.getColumn();
  var row = activeCell.getRow();

  var rs = ss.getSheets()[0];
  var rsName = rs.getName();
  var sheetName = sheet.getName();

  activeCell.setFontFamily('arial').setFontSize('10');


  if (sheetName == 'Summary') return;

  if (rsName == sheetName) {
    if (row == myNumbers.dashBalancesRow) {
      open();
    }
  } else {
    if ((row > myNumbers.expenseCarryOverRow) && (row <= myNumbers.expenseLastRow)) {
      if (col !== myNumbers.expenseDateColumn) {
        var typeVal = sheet.getRange(row, myNumbers.expenseTypeColumn).getValue();
        var descVal = sheet.getRange(row, myNumbers.expenseDescrColumn).getValue();
        var amountVal = sheet.getRange(row, myNumbers.expenseAmountColumn).getValue();
        var hasRecord = (typeVal != null && typeVal.toString().trim() !== "") ||
                        (descVal != null && descVal.toString().trim() !== "") ||
                        (amountVal != null && amountVal.toString().trim() !== "");

        if (hasRecord) {
          if (!isFutureMonthOrYear(sheet)) {
            sheet.getRange(row, myNumbers.expenseDateColumn).setValue(new Date());
          }
        } else {
          sheet.getRange(row, myNumbers.expenseDateColumn).clearContent();
        }
      }

      var splitRange = sheet.getRange(row, myNumbers.expenceSplitColumn);
      var splitRange1 = sheet.getRange(row, myNumbers.expenceSplit1Column);
      var splitRange2 = sheet.getRange(row, myNumbers.expenceSplit2Column);
      var amountRange = sheet.getRange(row, myNumbers.expenseAmountColumn);

      var sp1Col = sheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit1Column).getA1Notation().slice(0, 1);
      var sp2Col = sheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit2Column).getA1Notation().slice(0, 1);

      var formulaSp1 = '=IF (' + splitRange.getA1Notation() + '<> "N", IF(ISBLANK(' + amountRange.getA1Notation() + '),"", ROUND(';
      formulaSp1 += amountRange.getA1Notation() + '*$' + sp1Col + '$' + myNumbers.monthSplitConfigRow + ',2)),"")';

      var formulaSp2 = '=IF (' + splitRange.getA1Notation() + '<> "N", IF(ISBLANK(' + amountRange.getA1Notation() + '),"", ROUND(';
      formulaSp2 += amountRange.getA1Notation() + '*$' + sp2Col + '$' + myNumbers.monthSplitConfigRow + ',2)),"")';

      splitRange1.setValue(formulaSp1);
      splitRange2.setValue(formulaSp2);

      copyFormatting(row);
      validateType(row, col);
      validatePeriod(row, col);


      if (row < myNumbers.expenseLastRow) {
        copyFormatting(row + 1);
        validateType(row + 1, col);
        validatePeriod(row + 1, col);

      }

    }
  }

}



function validateType(row, col) {
  // Create data validation rule for monthly sheet

  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dashSheet = ss.getSheets()[0];
  var summarySheet = ss.getSheetByName("Summary")
  var lastFilledRow = new myUtil().getLastRowBeforeEmpty(summarySheet);

  if ((sheet == dashSheet) || (sheet == summarySheet)) return;

  if (col <= Math.max(myNumbers.expenseTypeColumn, myNumbers.expenseDescrColumn)) {

    //console.log(row + " " + myNumbers.expenseLastRow);

    var range = sheet.getRange(row, myNumbers.expenseTypeColumn);
    //console.log(myNumbers.expenseTypeColumn + " " + col);
    var summaryValues = summarySheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, lastFilledRow - myNumbers.expenseFirstRow + 1, 1).getValues().flat();
    console.log(lastFilledRow);
    console.log("drop down list: " + myNumbers.expenseFirstRow + " " + myNumbers.expenseTypeColumn + " " + (lastFilledRow - myNumbers.expenseFirstRow) + " " + 1);
    console.log("Summary values: " + summaryValues);

    /*if (firstEmptyIndex !== -1) {
      summaryValues = summaryValues.slice(0, firstEmptyIndex);
    }
    */

    var uniqueSummaryTypeValues = Array.from(new Set(summaryValues));
    var uniqueMonthlyTypeValues = Array.from(new Set(sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, myNumbers.expenseLastRow - myNumbers.expenseFirstRow, 1).getValues().flat()));
    console.log("Monthly values: " + uniqueMonthlyTypeValues);


    var uniqueTypeValues = uniqueSummaryTypeValues.concat(uniqueMonthlyTypeValues);

    console.log("Combined values: " + uniqueTypeValues);

    //var uniqueTypeValues = uniqueSummaryTypeValues.concat(uniqueMonthlyTypeValues.filter(value => !uniqueSummaryTypeValues.includes(value)));

    range.clearDataValidations();

    var validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueTypeValues)
      .build();
    range.setDataValidation(validationRule);

  }

}


function validatePeriod(row, col) {
  // Create data validation rule for expense period column
  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheets = ss.getSheets();
  var dashSheet = sheets[0];
  var summarySheet = ss.getSheetByName("Summary");

  if ((sheet == dashSheet) || (sheet == summarySheet)) return;

  // Source the validation rule (color template/settings) from January tab (index 1)
  var januarySheet = sheets[1];
  var sourceRange = januarySheet.getRange(myNumbers.expenseFirstRow, myNumbers.expencePeriodColumn);
  var templateRule = sourceRange.getDataValidation();

  if (!templateRule) return;

  // Get unique periods from the active sheet's period column
  var periods = new Set();

  var values = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expencePeriodColumn, myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1).getValues().flat();
  values.forEach(function (val) {
    if (val) periods.add(val.toString().trim());
  });

  var uniquePeriods = Array.from(periods).sort();

  if (uniquePeriods.length > 0) {
    // Build a new rule based on the template but with active sheet values
    var newRule = templateRule.copy().requireValueInList(uniquePeriods).build();
    var targetRange = sheet.getRange(row, myNumbers.expencePeriodColumn);
    targetRange.setDataValidation(newRule);
  }
}




function copyFormatting(row) {
  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dashSheet = ss.getSheets()[0];
  var summarySheet = ss.getSheetByName("Summary");

  if ((sheet == dashSheet) || (sheet == summarySheet)) return;

  var sourceRange = sheet.getRange(myNumbers.expenseFirstRow, 1, 1, sheet.getLastColumn());
  var targetRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());

  // Copy formatting (including conditional formatting)
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  // Copy data validation (including dropdown color schemes/chips)
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}


function showSplitDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ui/ChangeSplitDialog')
    .setWidth(380)
    .setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Change Split From Month');
}

function setSplitFromMonth(startMonthIndex, sp1Pct, sp2Pct) {
  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Update Dashboard as a display indicator
  sheets[0].getRange(myNumbers.dashSplitRow, myNumbers.dashSp1SplitColumn).setValue(sp1Pct);
  sheets[0].getRange(myNumbers.dashSplitRow, myNumbers.dashSp2SplitColumn).setValue(sp2Pct);

  // Update monthly config rows from startMonthIndex through December
  for (var i = startMonthIndex; i <= 12; i++) {
    var monthSheet = sheets[i];
    if (monthSheet) {
      monthSheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit1Column).setValue(sp1Pct);
      monthSheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit2Column).setValue(sp2Pct);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Split updated from month ' + startMonthIndex + ' onward.', 'Done', 3);
}

function migrateSplitConfig() {
  var myNumbers = new staticNumbers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var dash = sheets[0];

  var sp1Pct = dash.getRange(myNumbers.dashSplitRow, myNumbers.dashSp1SplitColumn).getValue();
  var sp2Pct = dash.getRange(myNumbers.dashSplitRow, myNumbers.dashSp2SplitColumn).getValue();

  for (var i = 1; i <= 12; i++) {
    var monthSheet = sheets[i];
    if (!monthSheet) continue;
    var sp1Cell = monthSheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit1Column);
    var sp2Cell = monthSheet.getRange(myNumbers.monthSplitConfigRow, myNumbers.expenceSplit2Column);
    if (sp1Cell.getValue() === '' && sp2Cell.getValue() === '') {
      sp1Cell.setValue(sp1Pct);
      sp2Cell.setValue(sp2Pct);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Split config initialized for all months.', 'Done', 3);
}

function isFutureMonthOrYear(sheet) {
  if (!sheet) return false;
  var now = new Date();
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth() + 1; // 1-12

  var name = sheet.getName();
  var parts = name.split(" ");
  if (parts.length < 2) return false;

  var monthsFull = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var monthsShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  var sheetMonth = monthsFull.indexOf(parts[0]) + 1;
  if (sheetMonth === 0) {
    sheetMonth = monthsShort.indexOf(parts[0]) + 1;
  }
  var sheetYear = parseInt(parts[1]);

  if (isNaN(sheetYear) || sheetMonth === 0) return false;

  if (sheetYear > currentYear) return true;
  if (sheetYear === currentYear && sheetMonth > currentMonth) return true;

  return false;
}
