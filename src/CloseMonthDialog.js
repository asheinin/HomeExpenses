function closeMonthPaid() {
  var mode = "f";
  payMonth(mode);
  return;
}  
  
  
function closeMonthCarryOver() {
  var mode = "c";
  payMonth(mode);
  return;
}  


function payMonthPartly() {
  var mode = "p";
  payMonth(mode);
  return;
} 

function payMonthFromBalance() {
  var mode = "s";
  payMonth(mode);
  return;
} 

function payMonthPartlyCurrent() {
  var mode = "pc";
  payMonth(mode);
  return;
} 

function payMonthFromBalanceCurrent() {
  var mode = "sc";
  payMonth(mode);
  return;
} 

function rebalanceExpenses() {
  
  var date = new Date();
  var currentMonth = date.getMonth();
  //Loop monthly sheets
  for (var i = 1; i <= currentMonth; i++) payMonth('c', i);
}  


function payMonth(mode, month) {
  if (!mode) return;

  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();
  
  var date = new Date();
  var currentMonth = date.getMonth();
  var formattedMonth = Utilities.formatDate(date,"GMT", "MMMM");
  
  //date.setMonth(currentMonth - 1,1)
  //var m = date.getMonth();
  
  if ((mode == "pc")||(mode == "sc")) {
    //current month
    var m = date.getMonth();
    mode = mode.slice(0, -1);
  } else {
    //past month
    var m = date.getMonth() -1;
  } 
  
  if (month) var m = month - 1;
  if (m == -1) m = 0;
  
  //To review. If month is Jan (prev year), to month is Dec. If month is Jan current year, month is 0
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var currYear = date.getFullYear();
  
  var fileName = ss.getName();
  
  var fileYear = fileName.split(" ").slice(-1).pop(); 

  Logger.log("1 " + currYear + " " + " " + fileYear);
  
  if (currYear > fileYear) m = 11;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[m+1];
  var sheetName = sheet.getName();
  var dash = ss.getSheets()[0];
  
  var aSp1 = sheet.getRange(myNumbers.expenseSp1MonthlyBalanceRow, myNumbers.expenseInitialBalanceCol).getValue();
  var aSp2 = sheet.getRange(myNumbers.expenseSp2MonthlyBalanceRow, myNumbers.expenseInitialBalanceCol).getValue();
  
  Logger.log("2 " + myNumbers.expenseSp1InitBalanceLeftRow + " " + myNumbers.expenseSp2InitBalanceLeftRow);
  
  var aFundSp1 = sheet.getRange(myNumbers.expenseSp1InitBalanceLeftRow, myNumbers.expenseInitialBalanceCol).getValue();
  var aFundSp2 = sheet.getRange(myNumbers.expenseSp2InitBalanceLeftRow, myNumbers.expenseInitialBalanceCol).getValue();
  
  //Logger.log(aSp1 + " " + aSp2 + " " + m + " " + currentMonth);
  //Logger.log(date);
  
  if ((Math.abs(aSp1) < myNumbers.thresholdLimitForClosingMonth) && (Math.abs(aSp2) < myNumbers.thresholdLimitForClosingMonth)) {
    var response = myUtils.dialogOK(sheetName + ': Month has already been settled or no payments are needed');
    return;
  }  
  
  var totalAmount = myNumbers.expenseTotalAmountRow;
  var totalPaid = myNumbers.expenseTotalPaidSp2Row + myNumbers.expenseTotalPaidSp2Row;
  
  if ((Math.abs(aSp1 + aSp2) > myNumbers.thresholdLimitForClosingMonth)||(totalAmount - totalPaid > myNumbers.thresholdLimitForClosingMonth)) {    
    var response = myUtils.dialogYN('Not all payments are marked as paid in this month, do you want to fix it first?');
    if (response == 'YES') return;
  }  
  
  if (aSp1 < 0) {
    //Spouse 1 owes
    var spouse = dash.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
    var a = Math.abs(aSp1).toFixed(2);
    var aF = Math.abs(aFundSp1).toFixed(2);
    var spFRow = myNumbers.expenseSp1InitBalancePaidRow;
    var spFCol = myNumbers.expenseInitialBalanceCol;
    var spRow = myNumbers.expenseSpToSpRow;
    var spColumn = myNumbers.expenseSp1ToSp2Column;
    var spCarryOverRow = myNumbers.expenseCarryOverRow;
    var spCarryOverColumn = myNumbers.expenseCarryOverSp1OwesColumn;
    Logger.log("Sp1 owes" + spouse + " " + a + " " + aSp1);
  } else {
    //Spouse 2 owes
    var spouse = dash.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();
    var a = Math.abs(aSp2).toFixed(2);
    var aF = Math.abs(aFundSp2).toFixed(2);
    var spFRow = myNumbers.expenseSp2InitBalancePaidRow;
    var spFCol = myNumbers.expenseInitialBalanceCol;
    var spRow = myNumbers.expenseSpToSpRow;
    var spColumn = myNumbers.expenseSp2ToSp1Column;
    var spCarryOverRow = myNumbers.expenseCarryOverRow;
    var spCarryOverColumn = myNumbers.expenseCarryOverSp2OwesColumn;
    Logger.log("Sp1 owes" + spouse + " " + a + " " + aSp2);
  }  
 
  switch(mode) {
    case "f":
        var response = myUtils.dialogYN(spouse + ' owes $' + a + '. Do you want to mark it as fully paid?');
        if (response == 'YES') updatePayments(m,a,spRow,spColumn,0,0,0);
        break;
    case "c":
        var response = 'YES';
        if (!month) response = myUtils.dialogYN(spouse + ' owes $' + a + '. Do you want to carry it over to '+ formattedMonth + '?');
        if (response == 'YES') updatePayments(m,a,spRow,spColumn,a,spCarryOverRow,spCarryOverColumn);
        Logger.log("UpdatePayments " + m+" " +a+" " +spRow+" " +spColumn+" " +a+" " +spCarryOverRow+" " +spCarryOverColumn);
        break;
    case "p":
        var response = myUtils.dialogQA(spouse + ' owes $' + a + '.','Enter paid amount by '+spouse); 
        if (response != -1) updatePayments(m,response.getResponseText(),spRow,spColumn,0,0,0);
        break;
    case "s":
        if (aF == 0) {
           var response = myUtils.dialogOK(spouse + ' does not have special fund left');
           break;
        }     
        var response = myUtils.dialogQA(spouse + ' owes $' + a + '.','Enter paid amount from ' + spouse + ' initial balance, $' + aF + ' left'); 
        if (response != -1) {
          var amount = Math.abs(response.getResponseText()).toFixed(2);
          Logger.log(amount + "<" + aF);
          if (aF - amount < 0) {
            Logger.log(amount + " " + aF + " " + (aF - amount));
            var response = myUtils.dialogOK(spouse + 'initial balance has $' + aF + ' left, cannot pay $' + amount);
            break;
          }
          if (a - amount < 0) {
            var response = myUtils.dialogYN(spouse + ' owes $' + a + '.' + spouse +', Do you want overpay?');
            if (response == 'NO') amount = a; 
          }
          Logger.log(amount);
          updatePayments(m,amount,spRow,spColumn,0,0,0,spFRow, spFCol);
        }  
        break;    
    default:
        return;
  }
  
  
} 



function updatePayments(month, amount,row, col, carryover, corow, cocol, spfrow, spfcol) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[month+1];
  var currentSheet = ss.getSheets()[month+2];
  var currentSheetName = currentSheet.getName();
  Logger.log(currentSheet.getName());
  
  //Add payment to closing month
  
  if ((spfrow)&&(spfcol)) {
    sheet.getRange(spfrow, spfcol).setValue(amount);
    return;
  }  
  
  for (var i=0; i<20; i++) {
    if (!sheet.getRange(row+i, col).getValue()) {
      sheet.getRange(row+i, col).setValue(amount);
      break;
    }  
  }
  
  Logger.log(carryover);
  
  if (carryover != 0) {  
    //Add carryover to current month
    const noteText = "Carryover amount";
    var currValue = currentSheet.getRange(corow, cocol).getValue()+0;
    Logger.log("currValue " + currValue + ":" + parseFloat(currValue) + ":" + parseFloat(carryover));
    var updatedValue = parseFloat(currValue) + parseFloat(carryover);
    var cellRange = currentSheet.getRange(corow, cocol);
    cellRange.setValue(updatedValue);
    cellRange.setNote(noteText);
  
    Logger.log(currentSheetName + " " + corow + " " + cocol + " " + updatedValue);
  }  
}  
  