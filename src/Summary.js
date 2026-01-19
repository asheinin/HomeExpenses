function Summary() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Summary");

  // Ensure we are working on the Summary sheet
  ss.setActiveSheet(summarySheet);

  // Generate the monthly expense summary data
  summaryExpenses();

  // Run historical spending analytics
  runAnalytics();

  // Generate the year-over-year comparison report
  runYearComparison();

}
