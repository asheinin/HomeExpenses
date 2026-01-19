function runAnnualExpensesChart(annualExpenses, startRow) {
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
    .setOption('legend', { position: 'in' })
    .setOption('height', 420)
    .setOption('width', 560)
    .setOption('pieSliceText', 'value')
    .setOption('pieSliceTextStyle', { color: '#FFFFFF', fontName: 'Arial', fontSize: 12, bold: true });
  summarySheet.insertChart(chartBuilder.build())
}

/**
 * Creates or updates the Historical Spending Analysis chart.
 * 
 * @param {Sheet} summarySheet - The sheet to insert the chart into
 * @param {number} startRow - The row where the analytics section starts
 * @param {Array} matrixData - The full data matrix
 * @param {Array} restOfData - The data matrix excluding the Year column
 * @param {Array} sortedTopTierTypes - List of top expense types
 * @param {number} maxTotalSpend - Maximum spending value for chart scaling
 * @param {Object} myNumbers - Static numbers config
 */
function drawHistoricalSpendingChart(summarySheet, startRow, matrixData, restOfData, sortedTopTierTypes, maxTotalSpend, myNumbers) {
  const charts = summarySheet.getCharts();
  charts.forEach(c => {
    if (c.getOptions().get('title') === 'Historical Spending Analysis') {
      summarySheet.removeChart(c);
    }
  });

  const colQ = 17;

  // Ranges
  const rangeYear = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsYearColumn, matrixData.length, 1);
  const rangeTotal = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsDataStartColumn, matrixData.length, 1);
  const rangeStacks = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsDataStartColumn + 1, matrixData.length, restOfData[0].length - 1);

  const seriesOptions = {};
  seriesOptions[0] = {
    type: 'line',
    dataLabel: 'value',
    lineSize: 0,
    pointSize: 0,
    visibleInLegend: false,
    dataLabelTextStyle: { bold: true, fontSize: 13, color: '#000' }
  };

  for (let i = 1; i <= sortedTopTierTypes.length + 1; i++) {
    seriesOptions[i] = { dataLabel: 'none' };
  }

  const chartBuilder = summarySheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(rangeYear)
    .addRange(rangeTotal)
    .addRange(rangeStacks)
    .setNumHeaders(1)
    .setOption('isStacked', true)
    .setOption('title', 'Historical Spending Analysis')
    .setOption('vAxis', {
      title: 'Amount ($)',
      gridlines: { count: 5 },
      viewWindow: { max: maxTotalSpend * 1.15 } // Add 15% padding at the top for labels
    })
    .setOption('hAxis.title', 'Year')
    .setOption('width', 665)
    .setOption('height', 434)
    .setOption('series', seriesOptions)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setPosition(startRow + 2, colQ, 0, 0)
    .build();

  summarySheet.insertChart(chartBuilder);
}

/**
 * Helper function to calculate the last row including all chart positions.
 * @param {Sheet} sheet - The sheet to analyze
 * @returns {number} - The last row occupied by data or charts
 */
function getLastRowIncludingCharts(sheet) {
  let lastRow = sheet.getLastRow();
  const charts = sheet.getCharts();

  charts.forEach(chart => {
    const containerInfo = chart.getContainerInfo();
    const anchorRow = containerInfo.getAnchorRow();
    const chartHeight = chart.getOptions().get('height') || 400;
    // Approximate rows: ~21 pixels per row in Google Sheets
    const chartEndRow = anchorRow + Math.ceil(chartHeight / 21);

    if (chartEndRow > lastRow) {
      lastRow = chartEndRow;
    }
  });

  return lastRow;
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