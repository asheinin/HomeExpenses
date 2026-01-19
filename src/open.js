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

  var formattedMonthCurrent = Utilities.formatDate(date, "GMT", "MMMM");

  date.setMonth(currentMonth - 1, 1)
  var m = date.getMonth();
  var y = date.getYear();

  var formattedMonthPast = Utilities.formatDate(date, "GMT", "MMMM");

  var specFundValueSp1 = sheet.getRange(myNumbers.dashBalancesRow, myNumbers.dashSpouse1NameColumn).getValue();
  var specFundValueSp2 = sheet.getRange(myNumbers.dashBalancesRow, myNumbers.dashSpouse2NameColumn).getValue();

  var specFund = ((specFundValueSp1 == 0) && (specFundValueSp2 == 0)) ? false : true;

  //var addNewExpenseRY = [{name: 'addNewExpenseRY', functionName: 'addNewExpense("ry")'}]; 
  //var addNewExpenseRM = [{name: 'addNewExpenseRM', functionName: 'addNewExpense("rm")'}];
  //var addNewExpenseOT = [{name: 'addNewExpenseOT', functionName: 'addNewExpense("ot")'}];


  //Check if this is current year file

  console.log("month: " + currentMonth);

  if (currYear > fileYear) currentMonth = 12;

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



  ui.createMenu('General Actions')
    .addItem("Rebalance YTD", "rebalanceExpenses")
    .addSeparator()
    .addItem("Calculate YTD Totals", "Summary")
    .addSeparator()
    .addItem("Create Next Year File", "createNewFile")
    .addItem("Tax Receipt", "createEOYDocument")
    .addToUi();


  // Set the data validation to require text in the form of an email address.
  var emailSp1Cell = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse1NameColumn);
  var emailSp2Cell = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse2NameColumn);
  var rule = SpreadsheetApp.newDataValidation().requireTextIsEmail().build();
  var a = emailSp1Cell.setDataValidation(rule);
  var b = emailSp2Cell.setDataValidation(rule);


  //set triggers if not set
  if (ScriptApp.getProjectTriggers().length < 3) {
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

    ScriptApp.newTrigger('open')
      .forSpreadsheet(ss)
      .onOpen()
      .create();

    ScriptApp.newTrigger('edit')
      .forSpreadsheet(ss)
      .onEdit()
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
      var splitRange = sheet.getRange(row, myNumbers.expenceSplitColumn);
      var splitRange1 = sheet.getRange(row, myNumbers.expenceSplit1Column); //5 Goldy
      var splitRange2 = sheet.getRange(row, myNumbers.expenceSplit2Column); //6 Alex 
      var amountRange = sheet.getRange(row, myNumbers.expenseAmountColumn);
      var splitDashRange1 = rs.getRange(myNumbers.dashSplitRow, myNumbers.dashSp1SplitColumn); //2 Alex
      var splitDashRange2 = rs.getRange(myNumbers.dashSplitRow, myNumbers.dashSp2SplitColumn); //3 Goldy 

      var formulaSp1 = '=IF (' + splitRange.getA1Notation() + '<> "N", IF(ISBLANK(' + amountRange.getA1Notation() + '),"", ROUND(';
      formulaSp1 += amountRange.getA1Notation() + '*' + rsName + '!$' + splitDashRange1.getA1Notation().slice(0, 1);
      formulaSp1 += '$' + myNumbers.dashSplitRow + ',2)),"")';

      var formulaSp2 = '=IF (' + splitRange.getA1Notation() + '<> "N", IF(ISBLANK(' + amountRange.getA1Notation() + '),"", ROUND(';
      formulaSp2 += amountRange.getA1Notation() + '*' + rsName + '!$' + splitDashRange2.getA1Notation().slice(0, 1);
      formulaSp2 += '$' + myNumbers.dashSplitRow + ',2)),"")';

      splitRange1.setValue(formulaSp1);
      splitRange2.setValue(formulaSp2);

      validateType(row, col);
      validatePeriod(row, col);
      copyFormatting(row);

      if (row < myNumbers.expenseLastRow) {
        validateType(row + 1, col);
        validatePeriod(row + 1, col);
        copyFormatting(row + 1);
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

  if ((sheet == dashSheet) || (sheet == summarySheet)) return;

  if (col <= Math.max(myNumbers.expenseTypeColumn, myNumbers.expenseDescrColumn)) {

    console.log(row + " " + myNumbers.expenseLastRow);

    var range = sheet.getRange(row, myNumbers.expenseTypeColumn);
    console.log(myNumbers.expenseTypeColumn + " " + col);
    var uniqueSummaryTypeValues = Array.from(new Set(summarySheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, summarySheet.getLastRow(), myNumbers.expenseTypeColumn).getValues().flat()));
    var uniqueMonthlyTypeValues = Array.from(new Set(sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, myNumbers.expenseLastRow - myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn).getValues().flat()));

    var uniqueTypeValues = uniqueSummaryTypeValues.concat(uniqueMonthlyTypeValues);

    //var uniqueTypeValues = uniqueSummaryTypeValues.concat(uniqueMonthlyTypeValues.filter(value => !uniqueSummaryTypeValues.includes(value)));
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


function fncOpenMyDialog() {
  //Open a dialog
  var htmlDlg = HtmlService.createHtmlOutputFromFile('HTML_myHtml')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(200)
    .setHeight(150);
  SpreadsheetApp.getUi()
    .showModalDialog(htmlDlg, 'A Title Goes Here');
};


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
