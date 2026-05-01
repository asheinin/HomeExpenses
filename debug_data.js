const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function debug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = [];
  months.forEach(m => {
    const sheet = ss.getSheetByName(m + " 2026");
    if (!sheet) return;
    const values = sheet.getRange('A2:D50').getValues();
    values.forEach(row => {
      const type = row[1]; // Type is col 2 (index 1)? Wait, myNumbers.expenseTypeColumn. Let's find out.
    });
  });
}
