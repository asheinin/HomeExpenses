function addNewExpenseRM() { // Recurrent from current month
  addNewExpense('rm');
  return;
}  

function addNewExpenseRY() { // Recurrent from begining of year
  addNewExpense('ry');
  return;
}  

function addNewExpenseOT() { // One time in current month 
  addNewExpense('ot');
  return;
} 

function addNewExpense(mode) {
  
 if (!mode) return;

 var myNumbers = new staticNumbers();
 var myUtils = new myUtil(); 
  
 var date = new Date();
 var currentMonth = date.getMonth() + 1;
 var formattedMonth = Utilities.formatDate(date,"GMT", "MMMM");
  
 if (mode == 'rm') {
   var m = currentMonth;
   var n = 12;
 }
 if (mode == 'ry') {
   var m = 1;
   var n = 12;
 }
 if (mode == 'ot') {
   var m = currentMonth;
   var n = m;
 }

  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0]; 
  
 var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
 var amountColumnArrayIndex = myNumbers.expenseAmountColumn-1;
  
 /*
 var response = myUtils.dialogQA('Create/Update Expense','Name'); 
 if (response == -1) throw "No New Expense"; 
 var newExpenseItem = response.getResponseText(); 
 */


// Get data from spreadsheet

 var data = sheet.getRange(myNumbers.expenseFirstRow, 1, myNumbers.expenseLastRow - myNumbers.expenseFirstRow, 1).getValues().flat();

 var response = myUtils.dialogList(data);
 if (response == -1) throw "No New Expense"; 
 var newExpenseItem = response.getResponseText(); 



  
 var response2 = myUtils.dialogQA('Enter expense amount or press <Cancel> to keep it unpopulated','Amount:'); 
  
 if (response2 != -1) {
    var expenseAmount = isNaN(parseFloat(response2.getResponseText())) ? -1 : parseFloat(response2.getResponseText()).toFixed(2); 
    if (expenseAmount < 0) throw "Amount Must Be Positive Number"; 
 }

 var pap = false;
 var expensePeriod = ""; 
  
 if (mode != 'ot') {
   var response3 = myUtils.dialogYN('Is this pre-authorized payment?');
   if (response3 == 'YES') pap = true;  
   
   var response4 = myUtils.dialogQA('Enter expense period e.g. "monthly" or press <Cancel> to keep it unpopulated','Period:'); 
  
   if (response4 != -1) {
     var expensePeriod = response4.getResponseText();
   }
 } 
  
 var inserted = false;
 var exists = false;
  
 for (var i = m; i <= n; i++) {
    
   var expenseItems = ss.getSheets()[i].getRange(myNumbers.expenseFirstRow,1, numOfRows).getValues(); 
   
   for (var j = 0; j <= numOfRows; j++) {
     if (expenseItems[j] == newExpenseItem) {
      if (!exists) var response5 = myUtils.dialogYN('This expense already exists. Do you want to update it?');
      if (response3 == 'YES') {
        pap = true; 
        exists = true;
      }  
      if (response2 != -1) ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseAmountColumn).setValue(expenseAmount);
      if (pap) {
        ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expensePAPColumn).setValue('PAP');
      } else {
        ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expensePAPColumn).setValue('');
      }  
      ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expencePeriodColumn).setValue(expensePeriod);
      ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenceSplitColumn).setValue("Y");


      Logger.log("3 " + expenseItems[j] + " " + newExpenseItem); 
   }
   
  }
 }
  
 if (exists) throw "Expense Already Exists";
   
 for (var i = m; i <= n; i++) {
   
   var expenseItems = ss.getSheets()[i].getRange(myNumbers.expenseFirstRow,1, numOfRows, myNumbers.expenseAmountColumn).getValues(); 
   Logger.log(expenseItems);
    
   for (var j = 0; j < numOfRows; j++) {  
      if ((expenseItems[j][0].length == 0) && (expenseItems[j][amountColumnArrayIndex].length == 0)) {
        Logger.log("check " + expenseItems[j][0] + " " + expenseItems[j][0].length + " " + expenseItems[j][amountColumnArrayIndex] + " " + expenseItems[j][amountColumnArrayIndex].length);
        Logger.log("insert to line " + j+1);
        ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseTypeColumn).setValue(newExpenseItem);
        if (response2 != -1) ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseAmountColumn).setValue(expenseAmount);

if (ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseFirstPayColumn).getValue() != "") {
            ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseAmountColumn).setValue(expenseAmount);
          } else {
            if (ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseFirstPayColumn+1).getValue() !="") {
              ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenseAmountColumn+1).setValue(expenseAmount);
            }  

          }


        if (pap) ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expensePAPColumn).setValue('PAP');
        ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expencePeriodColumn).setValue(expensePeriod);
        ss.getSheets()[i].getRange(j+myNumbers.expenseFirstRow,myNumbers.expenceSplitColumn).setValue("Y");
        inserted = true;
        break;
      } 
      Logger.log("check non 0 " + expenseItems[j][0] + " " + expenseItems[j][0].length + " " + expenseItems[j][amountColumnArrayIndex] + " " + expenseItems[j][amountColumnArrayIndex].length);
   }
   
   if (!inserted) throw "No More Space To Add Expense";
   Logger.log(i);
 
 }  
            
 try {
    
  }
    
 catch(err)  
  {           
      if (err == "No New Expense") {   
         return(-1);
      }  
      if (err == "No More Space To Add Expense") {
         ss.toast("No More Space To Add Expense","Error",5); 
         return(-2);
      } 
      if (err == "Amount Must Be Positive Number") {
         ss.toast("Amount Must Be Positive Number","Error",5); 
         return(-3);
      } 
      if (err == "Expense Already Exists") {   
         ss.toast("Expense already exists","Notice",5);
         return(0);  
      }else{
         Logger.log(err);
      }   
  }  
   
return;

} 