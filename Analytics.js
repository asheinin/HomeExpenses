function runAnalytics() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const myNumbers = new staticNumbers();

    // 1. Identify current year to include it
    const currentFileName = ss.getName();
    const currentFileYear = parseInt(currentFileName.split(" ").slice(-1).pop());
    const nowYear = new Date().getFullYear();
    const limitYear = isNaN(currentFileYear) ? nowYear : currentFileYear;

    // 2. Discover files (including current year)
    const files = DriveApp.searchFiles('title contains "Home payments " and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');
    const results = [];
    const allTypes = new Set();

    while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        const yearMatch = fileName.match(/Home payments (\d{4})/);

        if (yearMatch) {
            const year = parseInt(yearMatch[1]);
            if (year <= limitYear) {
                try {
                    const histSS = SpreadsheetApp.open(file);
                    const summarySheet = histSS.getSheetByName("Summary");

                    if (summarySheet) {
                        let totalSpend = 0;
                        const expenses = [];
                        const typeTotals = {};

                        const lastRow = summarySheet.getLastRow();
                        if (lastRow >= 3) {
                            const dataRange = summarySheet.getRange(3, 1, lastRow - 2, 3).getValues(); // Cols A, B, C
                            dataRange.forEach(row => {
                                const type = (row[0] || "Uncategorized").toString().trim();
                                const amount = parseFloat(row[2]) || 0;

                                if (amount > 0) {
                                    allTypes.add(type);
                                    typeTotals[type] = (typeTotals[type] || 0) + amount;
                                    totalSpend += amount;
                                    expenses.push({ type: type, amount: amount });
                                }
                            });
                        }

                        // Sort descending by amount and take top 3
                        expenses.sort((a, b) => b.amount - a.amount);
                        const top3 = expenses.slice(0, 3);

                        results.push({
                            year: year,
                            totalSpend: totalSpend,
                            typeTotals: typeTotals,
                            top3: top3
                        });
                    }
                } catch (e) {
                    console.error("Could not process file: " + fileName + ". Error: " + e.message);
                }
            }
        }
    }

    if (results.length === 0) {
        ui.alert("No 'Home payments' files found for years up to " + limitYear);
        return;
    }

    // Sort results by year
    results.sort((a, b) => a.year - b.year);
    const sortedTypes = Array.from(allTypes).sort();

    // 3. Create/Update Analytics Spreadsheet
    const analyticsFileName = "Home Expenses Analytics";
    let analyticsSS;
    const existingFiles = DriveApp.getFilesByName(analyticsFileName);

    if (existingFiles.hasNext()) {
        analyticsSS = SpreadsheetApp.open(existingFiles.next());
    } else {
        analyticsSS = SpreadsheetApp.create(analyticsFileName);
    }

    // --- Report Sheet ---
    let reportSheet = analyticsSS.getSheetByName("Summary Report");
    if (!reportSheet) {
        reportSheet = analyticsSS.getSheets()[0];
        reportSheet.setName("Summary Report");
    }
    reportSheet.clear();

    const headers = [
        "Year", "Total Spend",
        "Top 1 Amount", "Top 1 Type",
        "Top 2 Amount", "Top 2 Type",
        "Top 3 Amount", "Top 3 Type"
    ];
    reportSheet.appendRow(headers);

    results.forEach(res => {
        const row = [res.year, res.totalSpend];
        for (let i = 0; i < 3; i++) {
            if (res.top3[i]) {
                row.push(res.top3[i].amount, res.top3[i].type);
            } else {
                row.push("", "");
            }
        }
        reportSheet.appendRow(row);
    });

    const lastRowReport = reportSheet.getLastRow();
    reportSheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f3f3f3");
    [2, 3, 5, 7].forEach(colIndex => {
        reportSheet.getRange(2, colIndex, lastRowReport - 1, 1).setNumberFormat("$#,##0.00");
    });
    reportSheet.autoResizeColumns(1, 8);

    // --- Charts & Data Matrix Sheet ---
    let chartSheet = analyticsSS.getSheetByName("Charts");
    if (!chartSheet) {
        chartSheet = analyticsSS.insertSheet("Charts");
    }
    chartSheet.clear();

    // Data matrix for chart
    const matrixHeader = ["Year", ...sortedTypes];
    chartSheet.appendRow(matrixHeader);

    results.forEach(res => {
        const row = [res.year];
        sortedTypes.forEach(type => {
            row.push(res.typeTotals[type] || 0);
        });
        chartSheet.appendRow(row);
    });

    // Formatting matrix
    const matrixRange = chartSheet.getRange(1, 1, results.length + 1, sortedTypes.length + 1);
    chartSheet.getRange(2, 2, results.length, sortedTypes.length).setNumberFormat("$#,##0");

    // Create Stacked Column Chart
    const chart = chartSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(matrixRange)
        .setOption('isStacked', true)
        .setPosition(results.length + 3, 1, 0, 0)
        .setOption('title', 'Year-over-Year Expenses by Type')
        .setOption('vAxis.title', 'Amount ($)')
        .setOption('hAxis.title', 'Year')
        .setOption('width', 800)
        .setOption('height', 500)
        .build();

    chartSheet.insertChart(chart);
    chartSheet.autoResizeColumns(1, sortedTypes.length + 1);

    ui.alert("Analytics complete! View report and charts here: " + analyticsSS.getUrl());
}
