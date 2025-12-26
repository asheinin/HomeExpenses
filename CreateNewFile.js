function createNewFile() {
  
  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var sheet = ss.getSheets()[0];
  
  var editors =ss.getEditors();
  
  Logger.log(editors);
  
  try {
    
    
    var currentTime = new Date()
    //var month = currentTime.getMonth() + 1
    //var day = currentTime.getDate()

    var currentFileYearString = ssName.split(" ")[2];

    console.log("current file year - ", currentFileYearString);

    var currYear = currentTime.getFullYear()
    var currYearString = currYear.toString();

    if (currYearString==currentFileYearString) {
      var nextYear = currYear + 1;
    } else {
      var nextYear = currYear;
    }

    console.log("current year - ", currYearString);
    console.log("next year file - ",nextYear);

    //var ssNewDate = month + "/" + day + "/" + year;
    var newFileName = FILENAME + " " + nextYear;
    
    //check if file already exists
    
    if (DriveApp.getFilesByName(newFileName).hasNext() === true) throw "FileAlreadyExists";  
    
    //copy sourceSheet from one spreadsheet to another
    
    var copyDoc = myUtils.saveFile(newFileName); 
    if (copyDoc == -1) throw "CopyFailed";
    
    var expenseItems = myUtils.dialogYN('Do you want to copy over expense items from current year?');
    
    if (expenseItems == 'YES') var expenseAmounts = myUtils.dialogYN('Do you want to copy over expense amounts?');
    
    var ssNext = SpreadsheetApp.open(copyDoc);
    var sheetNext = ssNext.getSheets()[0];
    
    //For next year spreadsheet dashboard, advance 1 year for Month titles
    
    for (var i = 0; i < 12; i++) {
      var monthDate = sheet.getRange(myNumbers.dashFirstMonthRow+i,myNumbers.dashMonthNameColumn).getValue();
      var monthDateNextYear = formatAdvancedDate(monthDate);
      sheetNext.getRange(myNumbers.dashFirstMonthRow+i,myNumbers.dashMonthNameColumn).setValue(monthDateNextYear);
    } 
    
    //For next year spreadsheet dashboard, clean initial balance
    
    sheetNext.getRange(myNumbers.dashBalancesRow, myNumbers.dashSp1BalanceUsedColumn,1,2).clearContent();
    
    //For next year spreadsheet expense sheets, advance 1 year for sheet names and clean data
    
    var startCol;
    var numCol;
    
    for (var j = 1; j <=12; j++) {
      var sheetNext = ssNext.getSheets()[j];
      var numOfCols = sheetNext.getMaxColumns();
      var name = sheetNext.getSheetName().replace(currentFileYearString, nextYear.toString());
      sheetNext.setName(name);
      
      if (expenseItems == 'YES') {
        startCol = myNumbers.expenseDateColumn;
        numCol = myNumbers.expenseAmountColumn - myNumbers.expenseDateColumn +1;
        
        if (expenseAmounts == 'YES') {
          //startCol = myNumbers.expenseFirstPayColumn;
          startCol = myNumbers.expenseDateColumn;
          numCol = 1;
        }  
      } else {
        startCol = myNumbers.expenseTypeColumn;
        numCol = myNumbers.expenseAmountColumn - myNumbers.expenseTypeColumn +1;
      }  
      var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow;
      //clean expense sheet expenses, amounts data
      sheetNext.getRange(myNumbers.expenseFirstRow,startCol, numOfRows, numCol).clearContent();  
      //clean expense sheet paid data
      sheetNext.getRange(myNumbers.expenseFirstRow, myNumbers.expenseFirstPayColumn, numOfRows,numOfCols).clearContent();
      //clean expense sheet carryover data
      sheetNext.getRange(myNumbers.expenseCarryOverRow, 2, 1, numOfCols).clearContent();
      //clean expense sheet initial balance data
      sheetNext.getRange(myNumbers.expenseSp1InitBalancePaidRow,myNumbers.expenseInitialBalanceCol,1,1).clearContent();
      sheetNext.getRange(myNumbers.expenseSp2InitBalancePaidRow,myNumbers.expenseInitialBalanceCol,1,1).clearContent();
      
    }    
    
    var fileName = copyDoc.getName();
    var fileURL = copyDoc.getUrl();
    var notes = "<br>Prior to first use, click on Authorize button at the bottom of your Dashboard<br>";
    
    notifyNewFile(fileName,fileURL,notes);
    
    
  }
    
  catch(err)  
  {           
      if (err == "CopyFailed") {   
         ss.toast("No file created","Error",5);
         return(-1);
      }  
      if (err == "FileAlreadyExists") {   
         ss.toast("Next Year File Already Exists","Error",5);
         return(-2);  
      }else{
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
  var year = date.getFullYear()+1;

  return monthNames[monthIndex] + ' ' + year;
}