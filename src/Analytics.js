function runAnalytics() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const myNumbers = new staticNumbers();

    // 1. Identify current year to include it
    const currentFileName = ss.getName();
    const currentId = ss.getId();
    const currentYearMatch = currentFileName.match(/Home payments (\d{4})/);
    const currentFileYear = currentYearMatch ? parseInt(currentYearMatch[1]) : parseInt(currentFileName.split(" ").slice(-1).pop());

    const nowYear = new Date().getFullYear();
    const limitYear = isNaN(currentFileYear) ? nowYear : currentFileYear;

    // 2. Process current file first (active spreadsheet)
    const rawYearlyData = [];
    const globalTypeTotals = {};
    const processedYears = new Set();

    const currentYearData = processSheetData(ss, currentFileYear, globalTypeTotals);
    if (currentYearData && !isNaN(currentFileYear)) {
        rawYearlyData.push(currentYearData);
        processedYears.add(currentFileYear);
    }

    // 3. Discover other files
    const files = DriveApp.searchFiles('title contains "Home payments " and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');

    while (files.hasNext()) {
        const file = files.next();
        if (file.getId() === currentId) continue; // Skip the active spreadsheet

        const fileName = file.getName();
        const yearMatch = fileName.match(/Home payments (\d{4})/);

        if (yearMatch) {
            const year = parseInt(yearMatch[1]);
            // Skip current year (already processed) and any duplicates
            if (year <= limitYear && !processedYears.has(year)) {
                try {
                    const histSS = SpreadsheetApp.open(file);
                    const historicalData = processSheetData(histSS, year, globalTypeTotals);
                    if (historicalData) {
                        rawYearlyData.push(historicalData);
                        processedYears.add(year);
                    }
                } catch (e) {
                    console.error("Could not process file: " + fileName + ". Error: " + e.message);
                }
            }
        }
    }

    if (rawYearlyData.length === 0) {
        ui.alert("No 'Home payments' files found for years up to " + limitYear);
        return;
    }

    // 3. Identify Top tier types (all types that were in the Top 3 of ANY year)
    const topTierTypes = new Set();
    let maxTotalSpend = 0;
    rawYearlyData.forEach(res => {
        if (res.totalSpend > maxTotalSpend) maxTotalSpend = res.totalSpend;
        const yearTop1 = Object.keys(res.typeTotals)
            .sort((a, b) => res.typeTotals[b] - res.typeTotals[a])
            .slice(0, 1);
        yearTop1.forEach(t => topTierTypes.add(t));
    });

    const sortedTopTierTypes = Array.from(topTierTypes).sort((a, b) => globalTypeTotals[b] - globalTypeTotals[a]);

    // 4. Build the matrix data (Types only, no descriptions)
    const matrixData = [];
    const matrixHeader = ["Year", "Total Spend (Total)", ...sortedTopTierTypes, "Others"];
    matrixData.push(matrixHeader);

    rawYearlyData.sort((a, b) => a.year - b.year).forEach(res => {
        const yearTop1 = Object.keys(res.typeTotals)
            .sort((a, b) => res.typeTotals[b] - res.typeTotals[a])
            .slice(0, 1);

        let yearOthersSum = res.totalSpend;
        const row = [res.year, res.totalSpend];

        sortedTopTierTypes.forEach(type => {
            if (yearTop1.includes(type)) {
                const amt = res.typeTotals[type];
                row.push(amt || 0);
                yearOthersSum -= amt;
            } else {
                row.push(0);
            }
        });

        row.push(Math.max(0, yearOthersSum));
        matrixData.push(row);
    });

    // 6. Internal Data Placement on ACTIVE Summary Tab
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) return ui.alert("Summary sheet not found in the current spreadsheet.");

    // Remove existing Historical section and charts to prevent duplication
    const lastRowOfData = summarySheet.getLastRow();
    const existingHeaders = summarySheet.getRange(1, 1, Math.max(lastRowOfData, 1), 25).getValues();

    let startRow = getLastRowIncludingCharts(summarySheet) + 1;

    for (let i = 0; i < existingHeaders.length; i++) {
        const headerVal = existingHeaders[i][myNumbers.summaryAnalyticsYearColumn - 1];
        if (headerVal === "Historical Spending Summary" || existingHeaders[i][16] === "Historical Spending Summary") {
            startRow = i + 1;
            // Clear content across all analytics columns
            summarySheet.getRange(startRow, 1, Math.max(lastRowOfData - startRow + 30, 1), 25).clearContent();
            break;
        }
    }

    // Set Section Title at configured Year Column
    summarySheet.getRange(startRow + 1, myNumbers.summaryAnalyticsYearColumn).setValue("Historical Spending Summary")
        .setFontWeight("normal")
        .setFontFamily("Roboto")
        .setFontSize(14)
        .setFontColor("#666666");

    // "Trend" header for sparklines (Column 2)
    const trendCol = myNumbers.summaryAnalyticsYearColumn + 1;
    summarySheet.getRange(startRow + 2, trendCol).setValue("Trend");

    // Split matrix data into Year column and the rest
    const yearData = matrixData.map(row => [row[0]]);
    const restOfData = matrixData.map(row => row.slice(1));

    const yearRange = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsYearColumn, yearData.length, 1);
    yearRange.setValues(yearData);

    const dataRange = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsDataStartColumn, restOfData.length, restOfData[0].length);
    dataRange.setValues(restOfData);

    // Generate and set SPARKLINE formulas for Column B (Trend)
    // Formula format: =SPARKLINE(C4,{"charttype","bar";"max",MAX_VAL})
    const trendRange = summarySheet.getRange(startRow + 3, trendCol, matrixData.length - 1, 1);
    const sparklineFormulas = [];
    for (let i = 0; i < matrixData.length - 1; i++) {
        const amtCell = summarySheet.getRange(startRow + 3 + i, myNumbers.summaryAnalyticsDataStartColumn).getA1Notation();
        sparklineFormulas.push(['=SPARKLINE(' + amtCell + ',{"charttype","bar";"max",' + maxTotalSpend + '})']);
    }
    trendRange.setFormulas(sparklineFormulas);

    // Formatting
    yearRange.setFontWeight("bold").setBackground("#f8f9fa").setFontColor("#000000");
    summarySheet.getRange(startRow + 2, trendCol).setFontWeight("bold").setBackground("#f8f9fa").setFontColor("#000000");
    const headerRange = summarySheet.getRange(startRow + 2, myNumbers.summaryAnalyticsDataStartColumn, 1, restOfData[0].length);
    headerRange.setFontWeight("bold").setBackground("#f8f9fa").setFontColor("#000000");
    const numericDataRange = summarySheet.getRange(startRow + 3, myNumbers.summaryAnalyticsDataStartColumn, restOfData.length - 1, restOfData[0].length);
    numericDataRange.setNumberFormat("$#,##0;;").setFontColor("#000000");
    //summarySheet.autoResizeColumns(colA, matrixData[0].length);



    // 7. Create/Update Chart
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
        .setOption('width', 950)
        .setOption('height', 620)
        .setOption('series', seriesOptions)
        .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
        .setPosition(startRow + 2, colQ, 0, 0)
        .build();

    summarySheet.insertChart(chartBuilder);
}


/**
 * Compares monthly expenses between current year and previous year.
 * Creates a grouped bar chart on the Summary tab showing month-by-month comparison.
 * Should be called after runAnalytics().
 */
function runYearComparison() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const myNumbers = new staticNumbers();

    // 1. Get current file year
    const currentFileName = ss.getName();
    const currentFileYear = parseInt(currentFileName.split(" ").slice(-1).pop());

    if (isNaN(currentFileYear)) {
        console.log("Could not determine current file year from filename: " + currentFileName);
        return;
    }

    const previousYear = currentFileYear - 1;

    // 2. Find previous year file
    const files = DriveApp.searchFiles('title = "Home payments ' + previousYear + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');

    if (!files.hasNext()) {
        console.log("Previous year file not found: Home payments " + previousYear);
        return; // No previous year file, skip comparison
    }

    const prevFile = files.next();
    let prevSS;
    try {
        prevSS = SpreadsheetApp.open(prevFile);
    } catch (e) {
        console.error("Could not open previous year file: " + e.message);
        return;
    }

    // 3. Get Summary sheet and calculate lastRow including all charts
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) {
        console.log("Summary sheet not found");
        return;
    }

    const lastRowWithCharts = getLastRowIncludingCharts(summarySheet);
    const lastRowOfData = summarySheet.getLastRow();

    // 4. Extract monthly data from both files' Dashboard tabs
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    const currentDashboard = ss.getSheets()[0]; // Dashboard is first sheet
    const prevDashboard = prevSS.getSheets()[0];

    const currentMonthlyTotals = [];
    const prevMonthlyTotals = [];

    for (let i = 0; i < 12; i++) {
        const row = myNumbers.dashFirstMonthRow + i;

        // Current year data
        const currentVal = currentDashboard.getRange(row, myNumbers.dashAmountTotalBeforeSplitColumn).getValue();
        currentMonthlyTotals.push(parseFloat(currentVal) || 0);

        // Previous year data
        const prevVal = prevDashboard.getRange(row, myNumbers.dashAmountTotalBeforeSplitColumn).getValue();
        prevMonthlyTotals.push(parseFloat(prevVal) || 0);
    }

    // 5. Build comparison data matrix
    const comparisonData = [];
    comparisonData.push(["Month", currentFileYear.toString(), previousYear.toString()]);

    for (let i = 0; i < 12; i++) {
        comparisonData.push([months[i], currentMonthlyTotals[i], prevMonthlyTotals[i]]);
    }

    // 6. Write data to sheet (hidden area for chart data source)
    const dataStartRow = lastRowOfData + 3;
    const dataStartCol = myNumbers.summaryChartsStartColumn;

    // Clear any existing comparison data
    const existingDataRange = summarySheet.getRange(dataStartRow, dataStartCol, 15, 3);
    existingDataRange.clearContent();

    // Write new data
    const dataRange = summarySheet.getRange(dataStartRow, dataStartCol, comparisonData.length, 3);
    dataRange.setValues(comparisonData);
    dataRange.setFontColor("#FFFFFF"); // Hide the data (white on white)
    dataRange.setFontSize(1);

    // 7. Remove existing Year Comparison chart if present
    const charts = summarySheet.getCharts();
    charts.forEach(c => {
        if (c.getOptions().get('title') === 'Year-Over-Year Monthly Comparison') {
            summarySheet.removeChart(c);
        }
    });

    // 8. Create grouped bar chart
    const chartStartRow = lastRowOfData + 3;

    // Find max value for chart scaling
    const allValues = [...currentMonthlyTotals, ...prevMonthlyTotals];
    const maxValue = Math.max(...allValues);

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
        .setOption('width', 950)
        .setOption('height', 400)
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


/**
 * Extracts expense data from a spreadsheet's Summary tab.
 * 
 * @param {Spreadsheet} ss - The spreadsheet to process
 * @param {number} year - The year being processed
 * @param {Object} globalTypeTotals - Reference to update global type totals
 * @returns {Object|null} - The processed data object or null if Summary tab not found
 */
function processSheetData(ss, year, globalTypeTotals) {
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) return null;

    let yearlyTotal = 0;
    const yearlyTypeTotals = {};

    const lastRow = summarySheet.getLastRow();
    if (lastRow >= 3) {
        // Assume static column offsets: Type (Col 1), Amount (Col 3)
        // Note: myNumbers.expenseTypeColumn and myNumbers.summaryAmountColumn could be used,
        // but the logic here matches the original fixed indices (Cols A, B, C)
        const dataRange = summarySheet.getRange(3, 1, lastRow - 2, 3).getValues();
        dataRange.forEach(row => {
            const type = (row[0] || "Uncategorized").toString().trim();
            const amount = parseFloat(row[2]) || 0;

            if (amount > 0) {
                yearlyTypeTotals[type] = (yearlyTypeTotals[type] || 0) + amount;
                globalTypeTotals[type] = (globalTypeTotals[type] || 0) + amount;
                yearlyTotal += amount;
            }
        });
    }

    return {
        year: year,
        totalSpend: yearlyTotal,
        typeTotals: yearlyTypeTotals
    };
}
