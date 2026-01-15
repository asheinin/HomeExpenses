function cleanMonthsRM() { //clean current month +1 onward

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var currentMonth = date.getMonth() + 1; // 1-12
  if (currentMonth == 12) return;
  cleanMonths(currentMonth+1,ss); //from next month onward)
}


function cleanMonths(tab,ss) {

  if (!tab) var tab = 1;

    if (tab == 1) {
      var warning = "This action will clean all tabs in this spreadsheet. Do you want to continue?";
    } else {
      var date = new Date();
      var formattedMonthCurrent = Utilities.formatDate(date,"GMT", "MMMM");
      var warning = "This action will clean all tabs after " + formattedMonthCurrent + " in this spreadsheet. Do you want to continue?";
    }  

  var response = SpreadsheetApp.getUi().alert(warning, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (response == SpreadsheetApp.getUi().Button.NO) return;
    

  if (!ss) var ss = SpreadsheetApp.getActiveSpreadsheet();

  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();
  var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;

  var cleanedCount = 0;

  for (var j = tab; j <= 12; j++) {
      var targetSheet = ss.getSheets()[j];

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
      cleanedCount++;
  } 

  var message = "Finished cleaning tabs. Total months cleaned: " + cleanedCount;
  (cleanedCount == 0) ? ss.toast(message, "No Tabs Cleaned", -1) : ss.toast(message, "Success", 5);
    
}     