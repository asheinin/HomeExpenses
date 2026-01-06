function addNewExpenseRM() { // Recurrent from current month
  addNewExpenseAI('rm');
}

function addNewExpenseRY() { // Recurrent from beginning of year
  addNewExpenseAI('ry');
}

function addNewExpenseOT() { // One time in current month 
  addNewExpenseAI('ot');
}

function addNewExpenseAI(mode) {
  if (!mode) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var types = new Set();
  var myNumbers = new staticNumbers();

  // Get spouse names from Dashboard (B2 and C2)
  var dashboard = ss.getSheets()[0];
  var spouse1Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
  var spouse2Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();
  var spouseNames = [spouse1Name, spouse2Name];

  // Get distinct values from column 1 in sheets 2 to 12 (index 1 to 11), restricted to rows 2 to 50
  for (var i = 1; i < Math.min(sheets.length, 12); i++) {
    var sheet = sheets[i];
    var numRows = myNumbers.expenseLastRow - myNumbers.expenseCarryOverRow + 1;
    var values = sheet.getRange(myNumbers.expenseCarryOverRow, 1, numRows, 1).getValues();
    values.forEach(function (row) {
      if (row[0]) types.add(row[0]);
    });
  }

  var template = HtmlService.createTemplateFromFile('ExpenseForm');
  template.mode = mode;
  template.expenseTypes = Array.from(types).sort();
  template.spouseNames = spouseNames;
  var html = template.evaluate().setWidth(400).setHeight(560); // Increased height for new fields
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Expense');
}

//function processForm(formData) {
function processForm(newExpenseItem, newExpenseType, expenseAmount, pap, expensePeriod, mode, split, paid, paidBy) {
  try {
    var myNumbers = new staticNumbers();
    var myUtils = new myUtil();

    /*

    var newExpenseItem = formData.name;
    var newExpenseType = formData.type;
    var expenseAmount = formData.amount;
    var pap = (formData.pap == 'true');
    var expensePeriod = formData.period;

    */

    console.log("newExpenseItem, newExpenseType, expenseAmount, pap, expensePeriod, mode, split, paid, paidBy " + newExpenseItem + newExpenseType + expenseAmount + pap + expensePeriod + mode + split + paid + paidBy);

    if (!newExpenseItem) throw "No New Expense";

    // Validate amount
    if (expenseAmount != "") {
      var parsedAmount = parseFloat(expenseAmount);
      if (isNaN(parsedAmount) || parsedAmount < 0) {
        throw "Amount Must Be Positive Number";
      }
      expenseAmount = parsedAmount.toFixed(2);
    } else {
      expenseAmount = -1; // Use a flag for unpopulated amount
    }

    //var mode = formData.mode; // Get the mode from the form data
    var date = new Date();
    var currentMonth = date.getMonth() + 1;

    if (mode == 'rm') {
      var m = currentMonth;
      var n = 12;
    } else if (mode == 'ry') {
      var m = 1;
      var n = 12;
    } else { // 'ot'
      var m = currentMonth;
      var n = m;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var amountColumnArrayIndex = myNumbers.expenseAmountColumn - 1;
    var inserted = false;
    var exists = false;
    var response5 = 'NO'; // Default to 'NO'
    var sheetsToUpdate = [];

    // Check for existing expense and prompt for update
    for (var i = m; i <= n; i++) {
      var sheet = ss.getSheets()[i];
      var expenseItems = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn, numOfRows).getValues().flat();
      var existingIndex = expenseItems.indexOf(newExpenseItem);

      if (existingIndex !== -1) {
        exists = true;
        sheetsToUpdate.push({ sheet: sheet, row: existingIndex + myNumbers.expenseFirstRow });
      }
    }

    if (exists) {
      response5 = SpreadsheetApp.getUi().alert("This expense already exists. Do you want to update it?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
      if (response5 == SpreadsheetApp.getUi().Button.YES) {
        // Update existing expense
        sheetsToUpdate.forEach(function (item) {
          var sheet = item.sheet;
          var row = item.row;

          if (expenseAmount != -1) {
            sheet.getRange(row, myNumbers.expenseAmountColumn).setValue(expenseAmount);
            // Logic for first pay column
            if (sheet.getRange(row, myNumbers.expenseFirstPayColumn).getValue() == "") {
              sheet.getRange(row, myNumbers.expenseFirstPayColumn).setValue(expenseAmount);
            } else if (sheet.getRange(row, myNumbers.expenseFirstPayColumn + 1).getValue() == "") {
              sheet.getRange(row, myNumbers.expenseFirstPayColumn + 1).setValue(expenseAmount);
            }
          }
          sheet.getRange(row, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
          sheet.getRange(row, myNumbers.expencePeriodColumn).setValue(expensePeriod);
          sheet.getRange(row, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
          sheet.getRange(row, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : 'N');
        });
        ss.toast("Expense updated successfully.", "Success", 5);
        return;
      } else {
        throw "Expense Already Exists";
      }
    }

    // Create new expense
    for (var i = m; i <= n; i++) {
      var sheet = ss.getSheets()[i];
      var expenseData = sheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseAmountColumn).getValues();

      for (var j = 0; j < numOfRows; j++) {
        if (expenseData[j][0].length == 0 && (expenseData[j][amountColumnArrayIndex] === "" || expenseData[j][amountColumnArrayIndex] == null)) {
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn).setValue(newExpenseItem);
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn).setValue(newExpenseType);

          if (expenseAmount != -1) {
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseAmountColumn).setValue(expenseAmount);
          }
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expencePeriodColumn).setValue(expensePeriod);
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
          sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : 'N');
          inserted = true;
          break;
        }
      }
      if (!inserted) {
        throw "No More Space To Add Expense";
      }
    }
    ss.toast("New expense created successfully.", "Success", 5);
  } catch (err) {
    if (err == "No New Expense") {
      return (-1);
    }
    if (err == "No More Space To Add Expense") {
      ss.toast("No More Space To Add Expense", "Error", 5);
      return (-2);
    }
    if (err == "Amount Must Be Positive Number") {
      ss.toast("Amount Must Be Positive Number", "Error", 5);
      return (-3);
    }
    if (err == "Expense Already Exists") {
      ss.toast("Expense already exists", "Notice", 5);
      return (0);
    } else {
      Logger.log(err);
      ss.toast("An unexpected error occurred: " + err, "Error", 5);
    }
  }
}