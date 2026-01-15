function addNewExpenseRM() { // Recurrent from current month
  addNewExpenseAI('rm');
}

function addNewExpenseRY() { // Recurrent from beginning of year
  addNewExpenseAI('ry');
}

function addNewExpenseOT() { // One time in current month 
  addNewExpenseAI('ot');
}

function deleteExpenseRM() { // Delete from current month onwards
  deleteExpenseAI('rm');
}

function deleteExpenseRY() { // Delete from beginning of year
  deleteExpenseAI('ry');
}

function deleteExpenseAI(mode) {
  if (!mode) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var expenseMap = {}; // Use object to dedupe by name, storing name -> type
  var myNumbers = new staticNumbers();

  // Get distinct expense names and types from sheets 1 to 11 (monthly tabs)
  // Note: Type is in column 1, Name/Description is in column 2
  for (var i = 1; i < Math.min(sheets.length, 12); i++) {
    var sheet = sheets[i];
    var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var values = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn, numRows, 2).getValues(); // Get type (col 1) and name (col 2)
    values.forEach(function (row) {
      var type = row[0] ? row[0].toString().trim() : '';
      var name = row[1] ? row[1].toString().trim() : '';
      if (name !== '' && !expenseMap[name]) {
        expenseMap[name] = type;
      }
    });
  }

  // Convert to array of objects and sort by name
  var expenseItems = Object.keys(expenseMap).sort().map(function (name) {
    return { name: name, type: expenseMap[name] };
  });

  var template = HtmlService.createTemplateFromFile('DeleteExpenseForm');
  template.mode = mode;
  template.expenseItems = expenseItems;
  var html = template.evaluate().setWidth(450).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Delete Expense');
}


function processDeleteForm(expenseName, mode) {
  try {
    var myNumbers = new staticNumbers();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!expenseName) throw "No Expense Selected";

    var date = new Date();
    var currentMonth = date.getMonth() + 1;

    if (mode == 'rm') {
      var m = currentMonth;
      var n = 12;
    } else { // 'ry'
      var m = 1;
      var n = 12;
    }

    var numOfRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var deletedCount = 0;

    // Confirm deletion
    var response = SpreadsheetApp.getUi().alert(
      "Are you sure you want to delete '" + expenseName + "' from " + (mode == 'ry' ? 'all months' : 'current month onwards') + "?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }

    // Process each sheet and delete where found
    for (var i = m; i <= n; i++) {
      var sheet = ss.getSheets()[i];
      var expenseItems = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn, numOfRows).getValues().flat();
      var existingIndex = expenseItems.indexOf(expenseName);

      if (existingIndex !== -1) {
        var row = existingIndex + myNumbers.expenseFirstRow;
        // Clear the entire row content for this expense (columns 1 through expenseAmountColumn + some extra)
        var numColsToClear = myNumbers.expensePaidColumn; // Clear up to the Paid column
        sheet.getRange(row, 1, 1, numColsToClear).clearContent();
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      ss.toast("Expense '" + expenseName + "' deleted from " + deletedCount + " month(s).", "Success", 5);
    } else {
      ss.toast("Expense '" + expenseName + "' was not found in any sheet.", "Notice", 5);
    }

  } catch (err) {
    if (err == "No Expense Selected") {
      return (-1);
    } else {
      Logger.log(err);
      SpreadsheetApp.getActiveSpreadsheet().toast("An unexpected error occurred: " + err, "Error", 5);
    }
  }
}

function addNewExpenseAI(mode) {
  if (!mode) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var expenseTypes = new Set();
  var expensePeriods = new Set();
  var myNumbers = new staticNumbers();

  // Get spouse names from Dashboard (B2 and C2)
  var dashboard = ss.getSheets()[0];
  var spouse1Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
  var spouse2Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();
  var spouseNames = [spouse1Name, spouse2Name];

  // Get distinct values from columns in sheets 2 to 12 (index 1 to 11), restricted to rows 3 to 50
  for (var i = 1; i < Math.min(sheets.length, 12); i++) {
    var sheet = sheets[i];
    var numRows = myNumbers.expenseLastRow - myNumbers.expenseFirstRow + 1;
    var rangeValues = sheet.getRange(myNumbers.expenseFirstRow, 1, numRows, myNumbers.expensePAPColumn).getValues();
    rangeValues.forEach(function (row) {
      // Type is in column 1 (index 0)
      if (row[myNumbers.expenseTypeColumn - 1]) expenseTypes.add(row[myNumbers.expenseTypeColumn - 1]);
      // Period is in column 10 (index 9)
      if (row[myNumbers.expencePeriodColumn - 1]) expensePeriods.add(row[myNumbers.expencePeriodColumn - 1]);
    });
  }

  var template = HtmlService.createTemplateFromFile('ExpenseForm');
  template.mode = mode;
  template.expenseTypes = Array.from(expenseTypes).sort();
  template.expensePeriods = Array.from(expensePeriods).sort();
  template.spouseNames = spouseNames;
  var html = template.evaluate().setWidth(400).setHeight(560); // Increased height for new fields
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Expense');
}

//function processForm(formData) {
function processForm(newExpenseItem, newExpenseType, expenseAmount, pap, expensePeriod, mode, split, paid, paidByName, paidByIndex) {
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

    console.log("newExpenseItem, newExpenseType, expenseAmount, pap, expensePeriod, mode, split, paid, paidByName, paidByIndex " + newExpenseItem + " " + newExpenseType + " " + expenseAmount + " " + pap + " " + expensePeriod + " " + mode + " " + split + " " + paid + " " + paidByName + " " + paidByIndex);

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


    // First pass: Check if expense exists in any sheet to determine if we need to prompt user
    for (var i = m; i <= n; i++) {
      var sheet = ss.getSheets()[i];
      var expenseItems = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn, numOfRows).getValues().flat();
      var existingIndex = expenseItems.indexOf(newExpenseItem);

      if (existingIndex !== -1) {
        exists = true;
        break; // Found at least one, no need to check further
      }
    }

    // If expense exists in at least one sheet, prompt user
    if (exists) {
      response5 = SpreadsheetApp.getUi().alert("This expense already exists in some months. Do you want to update existing entries and create new ones where missing?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
      if (response5 != SpreadsheetApp.getUi().Button.YES) {
        throw "Expense Already Exists";
      }
    }

    // Second pass: Process each sheet - update if exists, create if not
    for (var i = m; i <= n; i++) {
      var sheet = ss.getSheets()[i];
      var expenseItems = sheet.getRange(myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn, numOfRows).getValues().flat();
      var existingIndex = expenseItems.indexOf(newExpenseItem);

      if (existingIndex !== -1) {
        // Update existing expense in this sheet
        var row = existingIndex + myNumbers.expenseFirstRow;
        sheet.getRange(row, myNumbers.expenseFirstPayColumn, 1, 2).clearContent();

        if (expenseAmount != -1) {
          sheet.getRange(row, myNumbers.expenseAmountColumn).setValue(expenseAmount);
          if (paidByIndex != 0) {
            if (paidByIndex == 1) sheet.getRange(row, myNumbers.expenseSecondPayColumn).setValue(expenseAmount);
            if (paidByIndex == 2) sheet.getRange(row, myNumbers.expenseFirstPayColumn).setValue(expenseAmount);
            console.log("paidByIndex !=0, paidByName " + paidByIndex + " " + paidByName);
          }
        }
        sheet.getRange(row, myNumbers.expenseTypeColumn).setValue(newExpenseType);
        sheet.getRange(row, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
        sheet.getRange(row, myNumbers.expencePeriodColumn).setValue(expensePeriod);
        sheet.getRange(row, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
        sheet.getRange(row, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : '');

      } else {
        // Create new expense in this sheet at first empty row
        var expenseData = sheet.getRange(myNumbers.expenseFirstRow, 1, numOfRows, myNumbers.expenseAmountColumn).getValues();
        inserted = false;

        for (var j = 0; j < numOfRows; j++) {
          if (expenseData[j][0].length == 0 && (expenseData[j][amountColumnArrayIndex] === "" || expenseData[j][amountColumnArrayIndex] == null)) {
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseDescrColumn).setValue(newExpenseItem);
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseTypeColumn).setValue(newExpenseType);

            if (expenseAmount != -1) {
              sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseAmountColumn).setValue(expenseAmount);
              if (paidByIndex != 0) {
                if (paidByIndex == 1) sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseSecondPayColumn, 1).setValue(expenseAmount);
                if (paidByIndex == 2) sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenseFirstPayColumn).setValue(expenseAmount);
                console.log("paidByIndex !=0, paidByName " + paidByIndex + " " + paidByName);
              }
            }
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expensePAPColumn).setValue(pap ? 'PAP' : '');
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expencePeriodColumn).setValue(expensePeriod);
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expenceSplitColumn).setValue(split ? 'Y' : 'N');
            sheet.getRange(j + myNumbers.expenseFirstRow, myNumbers.expensePaidColumn).setValue(paid ? 'Y' : '');
            inserted = true;
            break;
          }
        }
        if (!inserted) {
          throw "No More Space To Add Expense";
        }
      }
    }
    ss.toast("Expense processed successfully.", "Success", 5);
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