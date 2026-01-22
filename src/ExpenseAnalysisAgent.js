/**
 * Expense Analysis Agent
 * 
 * Evaluates monthly data against previous year and YoY comparisons.
 * Generates annual expense forecasts with recurrent expense trends.
 * Identifies high/above-normal expense spikes using AI analysis.
 * 
 * @author Alexander Sheinin
 */

/**
 * Main entry point for the Expense Analysis Agent.
 * Can be called from menu, directly, or via web app.
 * 
 * @param {boolean} returnData - If true, returns data object instead of showing dialog
 * @returns {Object|void} - Analysis results if returnData is true
 */
function runExpenseAnalysisAgent(returnData) {
    const myNumbers = new staticNumbers();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    // 1. Determine current context
    const now = new Date();
    const currentMonthIndex = now.getMonth(); // 0-11
    const currentMonthName = months[currentMonthIndex];
    const currentYear = now.getFullYear();

    // Get previous month info
    const prevMonthDate = new Date(currentYear, currentMonthIndex - 1, 1);
    const prevMonthIndex = prevMonthDate.getMonth();
    const prevMonthName = months[prevMonthIndex];
    const prevMonthYear = prevMonthDate.getFullYear();

    console.log(`Expense Analysis Agent running for: ${currentMonthName} ${currentYear}`);

    // 2. Get current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentFileYear = parseInt(ss.getName().split(" ").slice(-1).pop());

    // 3. Gather comparison data
    const comparisonData = getMonthlyComparisonData(
        ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months
    );

    // 4. Calculate annual forecast
    const forecastData = calculateAnnualForecast(
        ss, currentMonthIndex, currentYear, myNumbers, months
    );

    // 5. Detect expense spikes
    const spikeAnalysis = detectExpenseSpikes(
        comparisonData.current, comparisonData.yearAgo, myNumbers
    );

    // 6. Generate AI analysis
    const aiAnalysis = generateAgentAnalysis(comparisonData, forecastData, spikeAnalysis);

    // 7. Compile results
    const results = {
        timestamp: now.toISOString(),
        currentMonth: currentMonthName,
        currentYear: currentYear,
        comparison: comparisonData,
        forecast: forecastData,
        spikes: spikeAnalysis,
        aiInsights: aiAnalysis
    };

    // 8. Display or return
    if (returnData === true) {
        return results;
    } else {
        displayAgentResults(results, myNumbers);
    }
}


/**
 * Recalculates the expense analysis with custom assumption values.
 * Called from the dialog when user edits assumptions and clicks Recalculate.
 * 
 * @param {number} groceries - Monthly groceries estimate
 * @param {number} onlinePurchases - Monthly online purchases estimate
 * @param {number} gasoline - Monthly gasoline estimate
 * @param {number} misc - Monthly miscellaneous estimate
 */
function recalculateWithAssumptions(groceries, onlinePurchases, gasoline, misc) {
    const myNumbers = new staticNumbers();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    const now = new Date();
    const currentMonthIndex = now.getMonth();
    const currentMonthName = months[currentMonthIndex];
    const currentYear = now.getFullYear();

    const prevMonthDate = new Date(currentYear, currentMonthIndex - 1, 1);
    const prevMonthIndex = prevMonthDate.getMonth();
    const prevMonthYear = prevMonthDate.getFullYear();

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Custom assumptions from user input
    const customAssumptions = {
        groceries: parseFloat(groceries) || 0,
        onlinePurchases: parseFloat(onlinePurchases) || 0,
        gasoline: parseFloat(gasoline) || 0,
        misc: parseFloat(misc) || 0
    };

    const comparisonData = getMonthlyComparisonData(
        ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months
    );

    // Pass custom assumptions to forecast calculation
    const forecastData = calculateAnnualForecast(
        ss, currentMonthIndex, currentYear, myNumbers, months, customAssumptions
    );

    const spikeAnalysis = detectExpenseSpikes(
        comparisonData.current, comparisonData.yearAgo, myNumbers
    );

    const aiAnalysis = generateAgentAnalysis(comparisonData, forecastData, spikeAnalysis);

    const results = {
        timestamp: now.toISOString(),
        currentMonth: currentMonthName,
        currentYear: currentYear,
        comparison: comparisonData,
        forecast: forecastData,
        spikes: spikeAnalysis,
        aiInsights: aiAnalysis
    };

    displayAgentResults(results, myNumbers);
}


/**
 * Gathers monthly comparison data: current, previous month, and YoY.
 */
function getMonthlyComparisonData(ss, currentMonthIndex, currentYear, prevMonthIndex, prevMonthYear, myNumbers, months) {
    const currentMonthName = months[currentMonthIndex];
    const prevMonthName = months[prevMonthIndex];

    // Current month stats (may be partial if mid-month)
    const currentSS = getSpreadsheetForYear(currentYear);
    const currentStats = currentSS ? getMonthStats(currentSS, currentMonthName, currentYear, myNumbers) : null;

    // Previous month stats
    const prevSS = (prevMonthYear === currentYear) ? currentSS : getSpreadsheetForYear(prevMonthYear);
    const prevStats = prevSS ? getMonthStats(prevSS, prevMonthName, prevMonthYear, myNumbers) : null;

    // Year-over-year (same month last year)
    const yearAgoYear = currentYear - 1;
    const yearAgoSS = getSpreadsheetForYear(yearAgoYear);
    const yearAgoStats = yearAgoSS ? getMonthStats(yearAgoSS, currentMonthName, yearAgoYear, myNumbers) : null;

    // Calculate changes
    let vsPrevMonth = null;
    let vsYearAgo = null;

    if (currentStats && prevStats && prevStats.totalSpend > 0) {
        const diff = currentStats.totalSpend - prevStats.totalSpend;
        vsPrevMonth = {
            difference: diff,
            percentChange: ((diff / prevStats.totalSpend) * 100).toFixed(1)
        };
    }

    if (currentStats && yearAgoStats && yearAgoStats.totalSpend > 0) {
        const diff = currentStats.totalSpend - yearAgoStats.totalSpend;
        vsYearAgo = {
            difference: diff,
            percentChange: ((diff / yearAgoStats.totalSpend) * 100).toFixed(1)
        };
    }

    return {
        current: currentStats,
        currentMonthName: currentMonthName,
        previous: prevStats,
        previousMonthName: prevMonthName,
        previousMonthYear: prevMonthYear,
        yearAgo: yearAgoStats,
        yearAgoYear: yearAgoYear,
        vsPrevMonth: vsPrevMonth,
        vsYearAgo: vsYearAgo
    };
}


/**
 * Calculates annual expense forecast.
 * Uses posted expenses YTD + YoY trends + non-posted monthly projections.
 * 
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {number} currentMonthIndex - Current month (0-11)
 * @param {number} currentYear - Current year
 * @param {Object} myNumbers - Static numbers configuration
 * @param {Array} months - Month names array
 * @param {Object} customAssumptions - Optional custom assumption values
 * @returns {Object} Forecast data
 */
function calculateAnnualForecast(ss, currentMonthIndex, currentYear, myNumbers, months, customAssumptions) {
    // Get YTD posted expenses from Dashboard
    const dashboard = ss.getSheets()[0];
    let ytdPosted = 0;
    const monthlyTotals = [];

    for (let i = 0; i < 12; i++) {
        const row = myNumbers.dashFirstMonthRow + i;
        const val = parseFloat(dashboard.getRange(row, myNumbers.dashAmountTotalBeforeSplitColumn).getValue()) || 0;
        monthlyTotals.push(val);
        if (i <= currentMonthIndex) {
            ytdPosted += val;
        }
    }

    // Get YoY data for trend analysis
    const prevYear = currentYear - 1;
    const prevYearSS = getSpreadsheetForYear(prevYear);
    let prevYearTotal = 0;
    let prevYearMonthlyAvg = 0;

    if (prevYearSS) {
        const prevDashboard = prevYearSS.getSheets()[0];
        for (let i = 0; i < 12; i++) {
            const row = myNumbers.dashFirstMonthRow + i;
            const val = parseFloat(prevDashboard.getRange(row, myNumbers.dashAmountTotalBeforeSplitColumn).getValue()) || 0;
            prevYearTotal += val;
        }
        prevYearMonthlyAvg = prevYearTotal / 12;
    }

    // Non-posted monthly projections - use custom values if provided, else defaults
    const groceriesMonthly = (customAssumptions && customAssumptions.groceries !== undefined)
        ? customAssumptions.groceries
        : (myNumbers.agentGroceriesMonthly || 800);
    const onlinePurchasesMonthly = (customAssumptions && customAssumptions.onlinePurchases !== undefined)
        ? customAssumptions.onlinePurchases
        : (myNumbers.agentOnlinePurchasesMonthly || 800);
    const gasolineMonthly = (customAssumptions && customAssumptions.gasoline !== undefined)
        ? customAssumptions.gasoline
        : (myNumbers.agentGasolineMonthly || 400);
    const miscMonthly = (customAssumptions && customAssumptions.misc !== undefined)
        ? customAssumptions.misc
        : (myNumbers.agentMiscMonthly || 1000);
    const nonPostedMonthly = groceriesMonthly + onlinePurchasesMonthly + gasolineMonthly + miscMonthly;

    // Remaining months in the year
    const remainingMonths = 11 - currentMonthIndex; // 0-indexed, so December (11) means 0 remaining

    // Calculate projections for remaining months
    // Use YoY monthly average as baseline, adjusted by current year trend
    let avgPostedThisYear = currentMonthIndex > 0 ? ytdPosted / (currentMonthIndex + 1) : 0;
    let projectedPostedPerMonth = avgPostedThisYear > 0 ? avgPostedThisYear : prevYearMonthlyAvg;

    // Project remaining months
    const projectedRemaining = remainingMonths * (projectedPostedPerMonth + nonPostedMonthly);
    const annualForecast = ytdPosted + projectedRemaining;

    // YoY comparison for forecast
    const yoyForecastDiff = prevYearTotal > 0 ? annualForecast - prevYearTotal : null;
    const yoyForecastPct = prevYearTotal > 0 ? ((yoyForecastDiff / prevYearTotal) * 100).toFixed(1) : null;

    return {
        ytdPosted: ytdPosted,
        monthsCompleted: currentMonthIndex + 1,
        remainingMonths: remainingMonths,
        avgMonthlyThisYear: avgPostedThisYear,
        nonPostedMonthly: nonPostedMonthly,
        nonPostedBreakdown: {
            groceries: groceriesMonthly,
            onlinePurchases: onlinePurchasesMonthly,
            gasoline: gasolineMonthly,
            misc: miscMonthly
        },
        projectedRemaining: projectedRemaining,
        annualForecast: annualForecast,
        previousYearTotal: prevYearTotal,
        vsLastYear: {
            difference: yoyForecastDiff,
            percentChange: yoyForecastPct
        }
    };
}


/**
 * Detects expense spikes by comparing current month against YoY data.
 * Flags categories >30% above normal as "above normal" and >50% as "high/spike".
 */
function detectExpenseSpikes(currentStats, yearAgoStats, myNumbers) {
    const spikes = [];
    const aboveNormal = [];

    if (!currentStats || !yearAgoStats) {
        return { spikes, aboveNormal, hasAnomalies: false };
    }

    const currentCategories = currentStats.categoryTotals || {};
    const yearAgoCategories = yearAgoStats.categoryTotals || {};

    for (const category in currentCategories) {
        const currentAmount = currentCategories[category];
        const yearAgoAmount = yearAgoCategories[category] || 0;

        if (yearAgoAmount > 0) {
            const percentChange = ((currentAmount - yearAgoAmount) / yearAgoAmount) * 100;

            if (percentChange > 50) {
                spikes.push({
                    category: category,
                    currentAmount: currentAmount,
                    yearAgoAmount: yearAgoAmount,
                    percentChange: percentChange.toFixed(1),
                    severity: 'HIGH'
                });
            } else if (percentChange > 30) {
                aboveNormal.push({
                    category: category,
                    currentAmount: currentAmount,
                    yearAgoAmount: yearAgoAmount,
                    percentChange: percentChange.toFixed(1),
                    severity: 'ABOVE_NORMAL'
                });
            }
        } else if (currentAmount > 500) {
            // New category with significant spending
            spikes.push({
                category: category,
                currentAmount: currentAmount,
                yearAgoAmount: 0,
                percentChange: 'NEW',
                severity: 'NEW_CATEGORY'
            });
        }
    }

    return {
        spikes: spikes,
        aboveNormal: aboveNormal,
        hasAnomalies: spikes.length > 0 || aboveNormal.length > 0
    };
}


/**
 * Generates AI-powered analysis using Gemini.
 */
function generateAgentAnalysis(comparisonData, forecastData, spikeAnalysis) {
    const formatCurrency = (val) => new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(val || 0);

    let prompt = `You are a household expense analysis agent. Analyze this data and provide 3-4 actionable insights.
Be concise, helpful, and use bold text for key numbers. Format as HTML list (<ul><li>...</li></ul>).Also, please advice of assumptions made in html (projections only, list assumptions for gasoline, misc, groceries, online purchases)

CURRENT MONTH: ${comparisonData.currentMonthName}
- Total Spend: ${formatCurrency(comparisonData.current?.totalSpend)}
- Top Category: ${comparisonData.current?.topCategory?.name} (${formatCurrency(comparisonData.current?.topCategory?.amount)})
- Highest Single Expense: ${comparisonData.current?.highestSpend?.description} (${formatCurrency(comparisonData.current?.highestSpend?.amount)})

COMPARISONS:
- vs Previous Month (${comparisonData.previousMonthName}): ${comparisonData.vsPrevMonth ? comparisonData.vsPrevMonth.percentChange + '%' : 'N/A'}
- vs Same Month Last Year: ${comparisonData.vsYearAgo ? comparisonData.vsYearAgo.percentChange + '%' : 'N/A'}

ANNUAL FORECAST:
- YTD Posted: ${formatCurrency(forecastData.ytdPosted)}
- Projected Annual Total: ${formatCurrency(forecastData.annualForecast)}
- vs Last Year Annual: ${forecastData.vsLastYear?.percentChange ? forecastData.vsLastYear.percentChange + '%' : 'N/A'}

EXPENSE ANOMALIES:`;

    if (spikeAnalysis.hasAnomalies) {
        spikeAnalysis.spikes.forEach(s => {
            prompt += `\n- HIGH SPIKE: ${s.category} at ${formatCurrency(s.currentAmount)} (${s.percentChange}% vs last year)`;
        });
        spikeAnalysis.aboveNormal.forEach(s => {
            prompt += `\n- Above Normal: ${s.category} at ${formatCurrency(s.currentAmount)} (${s.percentChange}% vs last year)`;
        });
    } else {
        prompt += '\n- No significant anomalies detected.';
    }

    prompt += '\n\nProvide insights focusing on: spending trends, areas of concern, and actionable recommendations.';

    const response = callGemini(prompt);
    return response ? response.trim() : null;
}


/**
 * Displays agent results in a formatted HTML dialog.
 * (Interactive version for spreadsheet use)
 */
function displayAgentResults(results, myNumbers) {
    const html = generateAgentHtml(results, { isReadOnly: false, isStandalone: false, isEmbedded: false });

    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(500)
        .setHeight(780);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Expense Analysis Agent');
}


/**
 * Web App entry point - serves the mobile-responsive HTML page.
 * Sets to read-only mode for standalone access.
 */
function doGet(e) {
    try {
        const results = runExpenseAnalysisAgent(true);
        const isEmbedded = e && e.parameter && e.parameter.embed === 'true';
        const html = generateAgentHtml(results, {
            isReadOnly: true,
            isStandalone: !isEmbedded,
            isEmbedded: isEmbedded
        });

        return HtmlService.createHtmlOutput(html)
            .setTitle('Expense Analysis')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    } catch (error) {
        return HtmlService.createHtmlOutput(`<h1>Error</h1><p>${error.toString()}</p>`);
    }
}


/**
 * Handles POST requests (only used for spreadsheet-like interactive contexts if needed, 
 * but currently web app is read-only).
 */
function doPost(e) {
    // For now, web app is read-only. In the future, this could handle specific web-only actions.
    return doGet(e);
}


/**
 * Unified HTML generator for both spreadsheet dialog and web app.
 * 
 * @param {Object} results - Analysis results object
 * @param {Object} options - { isReadOnly: boolean, isStandalone: boolean, isEmbedded: boolean }
 * @returns {string} - Complete HTML page
 */
function generateAgentHtml(results, options) {
    const isReadOnly = options.isReadOnly || false;
    const isStandalone = options.isStandalone || false;
    const isEmbedded = options.isEmbedded || false;
    const formatCurrency = (val) => new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(val || 0);

    let html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Expense Analysis - ${results.currentMonth} ${results.currentYear}</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: ${isEmbedded ? 'transparent' : (isStandalone ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' : '#fff')};
            min-height: 100vh;
            padding: ${isEmbedded ? '0' : (isStandalone ? '10px' : '0')};
            color: #333;
            overflow-x: hidden;
        }
        .container { 
            max-width: 600px; 
            margin: 0 auto;
            position: relative;
            background: ${isStandalone || isEmbedded ? 'transparent' : '#fff'};
        }
        
        .header {
            background: ${isStandalone || isEmbedded ? 'rgba(255,255,255,0.95)' : 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'};
            color: ${isStandalone || isEmbedded ? '#333' : 'white'};
            border-radius: ${isStandalone || isEmbedded ? '12px' : '0 0 12px 12px'};
            padding: 16px;
            text-align: center;
            margin-bottom: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            position: relative;
            ${isEmbedded ? 'margin-top: 5px;' : ''}
        }
        .header h1 { font-size: 18px; margin-bottom: 2px; }
        .header .date { color: ${isStandalone || isEmbedded ? '#667eea' : 'rgba(255,255,255,0.9)'}; font-weight: 600; font-size: 13px; }
        
        .menu-btn {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            font-size: 20px;
            color: ${isStandalone || isEmbedded ? '#667eea' : 'white'};
            -webkit-tap-highlight-color: transparent;
            z-index: 10;
        }
        .menu-btn:active { background: rgba(102, 126, 234, 0.1); }

        .card {
            background: white;
            border-radius: 12px;
            padding: 14px;
            margin-bottom: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            border: 1px solid #eee;
        }
        .card-title {
            font-size: 13px; font-weight: 700; color: #555;
            margin-bottom: 10px; padding-bottom: 5px;
            border-bottom: 2px solid #667eea;
        }
        
        .metric {
            display: flex; justify-content: space-between; align-items: center;
            padding: 8px 0; border-bottom: 1px solid #f9f9f9;
        }
        .metric:last-child { border-bottom: none; }
        .metric-label { color: #666; font-size: 12px; }
        .metric-value { font-weight: 700; font-size: 14px; }
        .metric-value.big { font-size: 18px; color: #667eea; }
        .positive { color: #2E7D32; }
        .negative { color: #D32F2F; }
        
        .spike { 
            background: #FFEBEE; border-left: 4px solid #D32F2F; 
            padding: 8px; margin: 4px 0; border-radius: 0 6px 6px 0;
            display: flex; justify-content: space-between; font-size: 12px;
        }
        .above-normal { 
            background: #FFF3E0; border-left: 4px solid #FF9800; 
            padding: 8px; margin: 4px 0; border-radius: 0 6px 6px 0;
            display: flex; justify-content: space-between; font-size: 12px;
        }
        
        .ai-section {
            background: #E8F5E9; border-radius: 12px;
            padding: 14px; margin-bottom: 10px; border: 1px solid #C8E6C9;
        }
        .ai-title { font-size: 15px; font-weight: 700; color: #2E7D32; margin-bottom: 10px; }
        .ai-section ul { padding-left: 16px; }
        .ai-section li { margin-bottom: 8px; line-height: 1.5; font-size: 14px; }
        
        /* Modal Overlay */
        .modal-overlay {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.5);
            display: none;
            align-items: center;
            justify-content: center;
            padding: 20px;
            z-index: 1000;
        }
        .modal-content {
            background: white;
            border-radius: 16px;
            padding: 20px;
            width: 100%;
            max-width: 400px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            position: relative;
        }
        .modal-close {
            position: absolute;
            right: 15px;
            top: 15px;
            font-size: 24px;
            cursor: pointer;
            color: #999;
        }
        
        .assumptions-title { font-size: 15px; font-weight: 700; color: #F57C00; margin-bottom: 15px; border-bottom: 2px solid #FFE082; padding-bottom: 5px; }
        .item-row { display: flex; justify-content: space-between; align-items: center; padding: 8px 0; font-size: 13px; border-bottom: 1px solid #f9f9f9; }
        .item-label { color: #666; }
        .item-value { font-weight: 600; }
        
        .total-row {
            display: flex; justify-content: space-between; align-items: center;
            padding: 12px 0; margin-top: 10px; border-top: 2px solid #FFB74D;
            font-weight: 700; color: #F57C00; font-size: 15px;
        }
        
        input[type="number"] {
            width: 90px; padding: 8px; border: 1px solid #ddd;
            border-radius: 6px; font-size: 14px; text-align: right;
        }
        .recalculate-btn {
            width: 100%; padding: 14px; margin-top: 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; border: none; border-radius: 10px;
            font-size: 14px; font-weight: 700; cursor: pointer;
        }
        .recalculate-btn:disabled { opacity: 0.6; }

        .footer {
            text-align: center; color: ${isStandalone || isEmbedded ? 'rgba(0,0,0,0.4)' : '#aaa'};
            font-size: 10px; padding: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Expense Analysis</h1>
            <div class="date">${results.currentMonth} ${results.currentYear}</div>
            <div class="menu-btn" onclick="toggleAssumptions(true)" title="Forecast Assumptions">⋮</div>
        </div>
        
        <div class="card">
            <div class="card-title">Monthly Comparison</div>
            <div class="metric">
                <span class="metric-label">Current Month Total</span>
                <span class="metric-value">${formatCurrency(results.comparison.current?.totalSpend)}</span>
            </div>`;

    if (results.comparison.vsPrevMonth) {
        const pct = parseFloat(results.comparison.vsPrevMonth.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += `
            <div class="metric">
                <span class="metric-label">vs ${results.comparison.previousMonthName}</span>
                <span class="metric-value ${colorClass}">${pct > 0 ? '+' : ''}${pct}%</span>
            </div>`;
    }

    if (results.comparison.vsYearAgo) {
        const pct = parseFloat(results.comparison.vsYearAgo.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += `
            <div class="metric">
                <span class="metric-label">vs ${results.currentMonth} ${results.comparison.yearAgoYear} (YoY)</span>
                <span class="metric-value ${colorClass}">${pct > 0 ? '+' : ''}${pct}%</span>
            </div>`;
    }

    html += `
        </div>
        
        <div class="card">
            <div class="card-title">Annual Forecast</div>
            <div class="metric">
                <span class="metric-label">YTD Posted (${results.forecast.monthsCompleted} mo)</span>
                <span class="metric-value">${formatCurrency(results.forecast.ytdPosted)}</span>
            </div>
            <div class="metric">
                <span class="metric-label">Projected Annual Total</span>
                <span class="metric-value big">${formatCurrency(results.forecast.annualForecast)}</span>
            </div>`;

    if (results.forecast.vsLastYear.percentChange) {
        const pct = parseFloat(results.forecast.vsLastYear.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += `
            <div class="metric">
                <span class="metric-label">vs Last Year (${formatCurrency(results.forecast.previousYearTotal)})</span>
                <span class="metric-value ${colorClass}">${pct > 0 ? '+' : ''}${pct}%</span>
            </div>`;
    }

    html += `
        </div>`;

    if (results.spikes.hasAnomalies) {
        html += `
        <div class="card">
            <div class="card-title">⚠️ Expense Alerts</div>`;
        results.spikes.spikes.forEach(s => {
            html += `<div class="spike"><strong>${s.category}</strong><span class="negative">+${s.percentChange}%</span></div>`;
        });
        results.spikes.aboveNormal.forEach(s => {
            html += `<div class="above-normal"><strong>${s.category}</strong><span class="negative">+${s.percentChange}%</span></div>`;
        });
        html += `</div>`;
    }

    if (results.aiInsights) {
        html += `<div class="ai-section"><div class="ai-title">🤖 AI Analysis</div><div>${results.aiInsights}</div></div>`;
    }

    /* Modal for Assumptions */
    html += `
    <div class="modal-overlay" id="modalOverlay" onclick="if(event.target==this) toggleAssumptions(false)">
        <div class="modal-content">
            <div class="modal-close" onclick="toggleAssumptions(false)">×</div>
            <div class="assumptions-title">📝 Forecast Assumptions (Monthly)</div>`;

    const categories = [
        { id: 'groceries', label: 'Groceries', val: results.forecast.nonPostedBreakdown.groceries },
        { id: 'onlinePurchases', label: 'Online Purchases', val: results.forecast.nonPostedBreakdown.onlinePurchases },
        { id: 'gasoline', label: 'Gasoline', val: results.forecast.nonPostedBreakdown.gasoline },
        { id: 'misc', label: 'Miscellaneous', val: results.forecast.nonPostedBreakdown.misc }
    ];

    categories.forEach(cat => {
        html += `<div class="item-row">
            <span class="item-label">${cat.label}</span>
            <span class="item-value">
                ${isReadOnly
                ? formatCurrency(cat.val)
                : `$ <input type="number" id="${cat.id}" value="${cat.val}" min="0" step="50" oninput="updateTotal()">`}
            </span>
        </div>`;
    });

    html += `
            <div class="total-row">
                <span>Total Monthly Non-Posted</span>
                <span id="totalMonthly">${formatCurrency(results.forecast.nonPostedMonthly)}</span>
            </div>`;

    if (!isReadOnly) {
        html += `<button class="recalculate-btn" id="recalcBtn" onclick="recalculateForecast()">🔄 Recalculate Forecast</button>`;
    }

    html += `
        </div>
    </div>
    
    <div class="footer">Generated by Expense Analysis Agent</div>
    </div>`;

    html += `
    <script>
        function toggleAssumptions(show) {
            document.getElementById('modalOverlay').style.display = show ? 'flex' : 'none';
            if (show) document.body.style.overflow = 'hidden';
            else document.body.style.overflow = 'auto';
        }

        function updateTotal() {
            var g = parseFloat(document.getElementById('groceries').value) || 0;
            var o = parseFloat(document.getElementById('onlinePurchases').value) || 0;
            var ga = parseFloat(document.getElementById('gasoline').value) || 0;
            var m = parseFloat(document.getElementById('misc').value) || 0;
            var total = g + o + ga + m;
            document.getElementById('totalMonthly').textContent = '$' + total.toLocaleString();
        }
        
        function recalculateForecast() {
            var btn = document.getElementById('recalcBtn');
            var originalText = btn.textContent;
            btn.disabled = true;
            btn.textContent = '⏳ Recalculating...';
            
            var groceries = parseFloat(document.getElementById('groceries').value) || 0;
            var online = parseFloat(document.getElementById('onlinePurchases').value) || 0;
            var gasoline = parseFloat(document.getElementById('gasoline').value) || 0;
            var misc = parseFloat(document.getElementById('misc').value) || 0;
            
            google.script.run
                .withSuccessHandler(function() {
                    toggleAssumptions(false);
                })
                .withFailureHandler(function(err) {
                    alert('Error: ' + err);
                    btn.disabled = false;
                    btn.textContent = originalText;
                })
                .recalculateWithAssumptions(groceries, online, gasoline, misc);
        }
    </script>
    </body></html>`;
    return html;
}

