function createNewFile() {

  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var sheet = ss.getSheets()[0];

  var editors = ss.getEditors();

  Logger.log(editors);

  try {


    var currentTime = new Date()
    //var month = currentTime.getMonth() + 1
    //var day = currentTime.getDate()

    var currentFileYearString = ssName.split(" ")[2];

    console.log("current file year - ", currentFileYearString);

    var currYear = currentTime.getFullYear()
    var currYearString = currYear.toString();

    if (currYearString == currentFileYearString) {
      var nextYear = currYear + 1;
    } else {
      var nextYear = currYear;
    }

    console.log("current year - ", currYearString);
    console.log("next year file - ", nextYear);

    //var ssNewDate = month + "/" + day + "/" + year;
    var newFileName = FILENAME + " " + nextYear;

    //check if file already exists
    if (DriveApp.getFilesByName(newFileName).hasNext() === true) throw "FileAlreadyExists";

    //copy sourceSheet from one spreadsheet to another
    var copyDoc = myUtils.saveFile(newFileName);
    if (copyDoc == -1) throw "CopyFailed";

    var ssNext = SpreadsheetApp.open(copyDoc);
    var sheetNextDash = ssNext.getSheets()[0];

    //For next year spreadsheet dashboard, advance 1 year for Month titles
    for (var i = 0; i < 12; i++) {
      var monthDate = sheet.getRange(myNumbers.dashFirstMonthRow + i, myNumbers.dashMonthNameColumn).getValue();
      var monthDateNextYear = formatAdvancedDate(monthDate);
      sheetNextDash.getRange(myNumbers.dashFirstMonthRow + i, myNumbers.dashMonthNameColumn).setValue(monthDateNextYear);
    }

    //For next year spreadsheet dashboard, clean initial balance
    sheetNextDash.getRange(myNumbers.dashBalancesRow, myNumbers.dashSp1BalanceUsedColumn, 1, 2).clearContent();

    var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var oldDecSheet = ss.getSheets()[12]; // December is the 13th sheet (index 12)
    var decExpensesData = oldDecSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseSp1ToSp2Column).getValues();
    // 1. Clean all 12 monthly tabs in the new spreadsheet
    for (var j = 1; j <= 12; j++) {
      var targetSheet = ssNext.getSheets()[j];
      var name = targetSheet.getSheetName().replace(currentFileYearString, nextYear.toString());
      targetSheet.setName(name);

      // Completely clean the expense range while preserving formulas:
      // Clear definition block 1 (Cols 1-4: Type, Name, Date, Amount)
      targetSheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, 4).clearContent();
      // Clear definition block 2 (Cols 7-12: SplitYN, Pay1, Pay2, Period, Paid, PAP)
      // Note: Columns 5, 6, 13, 14 contain formulas and are preserved.
      targetSheet.getRange(myNumbers.expenseFirstRow, 7, numOfRows, 6).clearContent();

      // Reset background to white for the Amount column (Col 4)
      targetSheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseAmountColumn, numOfRows, 1).setBackground("white");

      // Clear other metadata rows
      targetSheet.getRange(myNumbers.expenseCarryOverRow, 2, 1, targetSheet.getMaxColumns() - 1).clearContent();
      targetSheet.getRange(myNumbers.expenseSp1InitBalancePaidRow, myNumbers.expenseInitialBalanceCol, 1, 1).clearContent();
      targetSheet.getRange(myNumbers.expenseSp2InitBalancePaidRow, myNumbers.expenseInitialBalanceCol, 1, 1).clearContent();
    }
  

    // 2. Load recurring expenses from December (old) into January (new)
    var janSheet = ssNext.getSheets()[1];
    var janTargetRow = myNumbers.expenseFirstRow;
    for (var k = 0; k < decExpensesData.length; k++) {
      var periodValue = decExpensesData[k][myNumbers.expencePeriodColumn - 1]; // expencePeriodColumn is 10
      if (periodValue && periodValue.toString().trim() !== "") {
        // Carry over only the "definition" columns
        janSheet.getRange(janTargetRow, myNumbers.expenseTypeColumn).setValue(decExpensesData[k][0]);
        janSheet.getRange(janTargetRow, myNumbers.expenseDescrColumn).setValue(decExpensesData[k][1]);
        janSheet.getRange(janTargetRow, myNumbers.expenseAmountColumn).setValue(decExpensesData[k][3]);
        janSheet.getRange(janTargetRow, myNumbers.expenceSplitColumn).setValue(decExpensesData[k][6]);
        janSheet.getRange(janTargetRow, myNumbers.expencePeriodColumn).setValue(decExpensesData[k][9]);
        janSheet.getRange(janTargetRow, myNumbers.expensePAPColumn).setValue(decExpensesData[k][11]);

        // Ensure Paid status is reset to "N"
        janSheet.getRange(janTargetRow, myNumbers.expensePaidColumn).setValue("N");
        janTargetRow++;
      }
    }

    // 3. Clean the Summary sheet in the new spreadsheet
    var summarySheetNext = ssNext.getSheetByName("Summary");
    if (summarySheetNext) {
      var lastRowSummary = summarySheetNext.getLastRow();
      if (lastRowSummary > myNumbers.summarySumRow) {
        summarySheetNext.getRange(myNumbers.summarySumRow + 1, 1, lastRowSummary - myNumbers.summarySumRow, summarySheetNext.getMaxColumns()).clearContent();
      }

      // Remove all charts from the Summary tab
      var charts = summarySheetNext.getCharts();
      for (var c = 0; c < charts.length; c++) {
        summarySheetNext.removeChart(charts[c]);
      }
    }

    var fileName = copyDoc.getName();
    var fileURL = copyDoc.getUrl();
    var notes = "<br>Prior to first use, click on Authorize button at the bottom of your Dashboard<br>";

    notifyNewFile(fileName, fileURL, notes);


  }

  catch (err) {
    if (err == "CopyFailed") {
      ss.toast("No file created", "Error", 5);
      return (-1);
    }
    if (err == "FileAlreadyExists") {
      ss.toast("Next Year File Already Exists", "Error", 5);
      return (-2);
    } else {
      Logger.log(err);
    }
  }

  return;

}



function formatAdvancedDate(date) {
  var monthNames = [
    "January", "February", "March",
    "April", "May", "June", "July",
    "August", "September", "October",
    "November", "December"
  ];

  var day = date.getDate();
  var monthIndex = date.getMonth();
  var year = date.getFullYear() + 1;

  return monthNames[monthIndex] + ' ' + year;
}