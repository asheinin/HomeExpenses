function graph(annualExpenses, startRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var summarySheet = ss.getSheetByName("Summary")
  ss.setActiveSheet(summarySheet);

// Add a pie chart to the summary sheet
  var chartBuilder = summarySheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(summarySheet.getRange(startRow, 1, annualExpenses, 1))
    .addRange(summarySheet.getRange(startRow, 3, annualExpenses, 1))
    .setPosition(1, 17, 0, 0)
    .setOption('title', 'Annual Expenses')
    .setOption('legend', {position: 'in'})
    .setOption('height', 600)
    .setOption('width', 800)
    .setOption('pieSliceText', 'value')
    .setOption('pieSliceTextStyle', {color: '#FFFFFF', fontName: 'Arial', fontSize: 12, bold: true});
  summarySheet.insertChart(chartBuilder.build())

}  

// help: https://ctrlq.org/code/20094-google-charts-dashboard-with-google-sheets

/*

function graph(annualExpenses) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var summarySheet = ss.getSheetByName("Summary")
  ss.setActiveSheet(summarySheet);

// Add a pie chart to the summary sheet
  var chartBuilder = summarySheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(summarySheet.getRange(3, 1, annualExpenses.length-2, 1))
    .addRange(summarySheet.getRange(3, 3, annualExpenses.length-2, 1))
    .setPosition(1, 17, 0, 0)
    .setOption('title', 'Annual Expenses')
    .setOption('legend', {position: 'in'})
    .setOption('height', 600)
    .setOption('width', 800)
    .setOption('pieSliceText', 'value')
    .setOption('pieSliceTextStyle', {color: '#FFFFFF', fontName: 'Arial', fontSize: 12, bold: true});
  summarySheet.insertChart(chartBuilder.build())

}  

// help: https://ctrlq.org/code/20094-google-charts-dashboard-with-google-sheets

*/