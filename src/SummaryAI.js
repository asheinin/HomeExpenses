function summaryExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var myNumbers = new staticNumbers();

  var numofsheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  if (numofsheets == 1) {
    return;
  }

  const summarySheet = ss.getSheetByName('Summary');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var date = new Date();
  var currentMonth = date.getMonth();
  var currentYear = date.getFullYear();
  var fileName = ss.getName();

  var fileYear = fileName.split(" ").slice(-1).pop();

  Logger.log(currentYear + " " + " " + fileYear);

  //Check if this is current year file

  if (currentYear > fileYear) {
    currentMonth = 12;
    year = fileYear;
  } else {
    year = currentYear;
  }


  // Clear the summary sheet
  summarySheet.clear();

  summarySheet.clearContents();
  var chts = summarySheet.getCharts();
  for (var i = 0; i < chts.length; i++) {
    summarySheet.removeChart(chts[i]);
  }

  // Set the header row
  const header = ['Type', 'Description', 'Total Amount', ...months];
  const summary = ['Total'];
  summarySheet.appendRow(header);
  summarySheet.appendRow(summary);

  // Object to store aggregated data
  const data = {};

  // Iterate through each month
  months.forEach((month, index) => {
    const sheetName = `${month} ${year}`;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const range = sheet.getRange('A2:D50');
    const values = range.getValues();

    values.forEach(row => {
      const type = row[myNumbers.expenseTypeColumn - 1];
      const description = row[myNumbers.expenseDescrColumn - 1];
      const amount = row[myNumbers.expenseAmountColumn - 1];

      if (!type || !amount) return;

      if (!data[type]) {
        data[type] = {
          descriptions: new Set(),
          totalAmount: 0,
          monthlyAmounts: Array(12).fill(0)
        };
      }

      if (description) {
        data[type].descriptions.add(description);
      }
      if (index <= currentMonth) {
        data[type].totalAmount += amount;
      }
      data[type].monthlyAmounts[index] += amount;
    });
  });

  // Write aggregated data to the summary sheet
  Object.keys(data).forEach(type => {
    const row = [
      type,
      Array.from(data[type].descriptions).join(', '),
      data[type].totalAmount,
      ...data[type].monthlyAmounts
    ];
    summarySheet.appendRow(row);
  });


  // Set font color for future months and summary row, add formulas for summary row starting with total
  const lastRow = summarySheet.getLastRow();
  for (let i = myNumbers.summaryAmountColumn; i <= myNumbers.summaryAmountColumn + 13; i++) { // Adjusting for the first 3 columns
    if (i >= currentMonth + myNumbers.summaryAmountColumn + 1) { //font color for future months
      const range = summarySheet.getRange(1, i + 1, lastRow, 1);
      range.setFontColor('lightgrey');
    }

    const sumRange1 = summarySheet.getRange(myNumbers.summarySumRow + 1, i, lastRow - myNumbers.summarySumRow, 1);
    const sumRangeValue = summarySheet.getRange(myNumbers.summarySumRow, i, 1, 1);
    var formulaSum = '=SUM(' + sumRange1.getA1Notation() + ')';
    console.log(formulaSum);
    sumRangeValue.setValue(formulaSum);
  }

  // Set first two columns to bold
  const typeRange = summarySheet.getRange(1, myNumbers.expenseTypeColumn, lastRow, 1);
  const descriptionRange = summarySheet.getRange(1, myNumbers.expenseDescrColumn, lastRow, 1);
  const totalAmountRange = summarySheet.getRange(1, myNumbers.summaryAmountColumn, lastRow, 1);
  typeRange.setFontWeight('bold');
  descriptionRange.setFontWeight('bold');
  totalAmountRange.setFontWeight('bold');
  descriptionRange.setWrap(true);

  graph(lastRow, myNumbers.summarySumRow + 1);

  runAnalytics();

  runYearComparison();

}


