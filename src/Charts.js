/**
 * Creates and inserts a pie chart of annual expenses into the Summary sheet.
 * 
 * @param {number} annualExpenses - The number of expense rows to include in the chart.
 * @param {number} startRow - The starting row of the expense data in the Summary sheet.
 */
function drawAnnualExpensesChart(annualExpenses, startRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Summary");
  const myNumbers = new staticNumbers();

  // Define ranges: Column same as expenseTypeColumn and summaryAmountColumn (Amount)
  var typeRange = summarySheet.getRange(startRow, myNumbers.expenseTypeColumn, annualExpenses, 1);
  var amountRange = summarySheet.getRange(startRow, myNumbers.summaryAmountColumn, annualExpenses, 1);

  // Build the pie chart
  var chartBuilder = summarySheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(typeRange)
    .addRange(amountRange)
    .setPosition(1, myNumbers.summaryChartsStartColumn, 0, 0) // Position at row 1, column myNumbers.summaryChartsStartColum
    .setOption('title', 'Annual Expenses')
    .setOption('legend', { position: 'in' })
    .setOption('height', 420)
    .setOption('width', 560)
    .setOption('pieSliceText', 'value')
    .setOption('pieSliceTextStyle', { color: '#FFFFFF', fontName: 'Arial', fontSize: 12, bold: true });

  // Insert the chart into the sheet
  summarySheet.insertChart(chartBuilder.build());
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
  //const myNumbers = new staticNumbers();

  charts.forEach(c => {
    if (c.getOptions().get('title') === 'Historical Spending Analysis') {
      summarySheet.removeChart(c);
    }
  });

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
    .setPosition(startRow + 2, myNumbers.summaryChartsStartColumn, 0, 0)
    .build();

  summarySheet.insertChart(chartBuilder);
}




/**
 * Creates or updates the Year-Over-Year Monthly Comparison chart.
 * 
 * @param {Sheet} summarySheet - The sheet to insert the chart into
 * @param {Range} dataRange - The range containing the comparison data
 * @param {number} chartStartRow - The row where the chart should be positioned
 * @param {number} currentFileYear - The year of the current data
 * @param {number} previousYear - The year of the historical data
 * @param {Object} myNumbers - Static numbers config
 */
function DrawYoYComparisonChart(summarySheet, dataRange, chartStartRow, currentFileYear, previousYear, myNumbers) {
  // Remove existing Year Comparison chart if present
  const charts = summarySheet.getCharts();
  charts.forEach(c => {
    if (c.getOptions().get('title') === 'Year-Over-Year Monthly Comparison' ||
      c.getOptions().get('title') === 'Current vs Previous Year Monthly Comparison') {
      summarySheet.removeChart(c);
    }
  });

  const chartBuilder = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setNumHeaders(1)
    .setOption('title', 'Year-Over-Year Monthly Comparison')
    .setOption('isStacked', false)
    .setOption('vAxis', {
      title: 'Amount ($)',
      ticks: [0, 3000, 6000, 9000, 12000],
      viewWindow: { min: 0, max: 12000 },
      format: '$#,##0'
    })
    .setOption('hAxis', {
      title: 'Month',
      slantedText: true,
      slantedTextAngle: 45
    })
    .setOption('width', 665)
    .setOption('height', 280)
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#4285F4', '#BDBDBD']) // Blue for current year, Gray for previous
    .setOption('bar', { groupWidth: '70%' })
    .setOption('series', {
      0: { dataLabel: 'value', dataLabelTextStyle: { fontSize: 9 } },
      1: { dataLabel: 'value', dataLabelTextStyle: { fontSize: 9 } }
    })
    .setPosition(chartStartRow, myNumbers.summaryAnalyticsMonthColumn, 0, 0)
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


