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
 */
function calculateAnnualForecast(ss, currentMonthIndex, currentYear, myNumbers, months) {
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

    // Non-posted monthly projections (configurable in staticNumbers)
    const groceriesMonthly = myNumbers.agentGroceriesMonthly || 800;
    const onlinePurchasesMonthly = myNumbers.agentOnlinePurchasesMonthly || 800;
    const gasolineMonthly = myNumbers.agentGasolineMonthly || 400;
    const nonPostedMonthly = groceriesMonthly + onlinePurchasesMonthly + gasolineMonthly;

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
            gasoline: gasolineMonthly
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
Be concise, helpful, and use bold text for key numbers. Format as HTML list (<ul><li>...</li></ul>).

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
 */
function displayAgentResults(results, myNumbers) {
    const formatCurrency = (val) => new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(val || 0);

    let html = `
    <style>
      body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 0; padding: 20px; color: #333; }
      .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; margin: -20px -20px 20px -20px; text-align: center; }
      .header h2 { margin: 0; font-size: 20px; }
      .header .date { font-size: 12px; opacity: 0.9; margin-top: 5px; }
      .section { background: #f9f9f9; border-radius: 8px; padding: 15px; margin-bottom: 15px; }
      .section-title { font-size: 14px; font-weight: bold; color: #555; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 5px; }
      .metric { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #eee; }
      .metric:last-child { border-bottom: none; }
      .metric-label { color: #666; }
      .metric-value { font-weight: bold; }
      .positive { color: #2E7D32; }
      .negative { color: #D32F2F; }
      .neutral { color: #666; }
      .spike { background: #FFEBEE; border-left: 4px solid #D32F2F; padding: 8px 12px; margin: 5px 0; border-radius: 0 4px 4px 0; }
      .above-normal { background: #FFF3E0; border-left: 4px solid #FF9800; padding: 8px 12px; margin: 5px 0; border-radius: 0 4px 4px 0; }
      .ai-section { background: #E8F5E9; border-radius: 8px; padding: 15px; margin-bottom: 15px; }
      .ai-title { font-size: 14px; font-weight: bold; color: #2E7D32; margin-bottom: 10px; }
      ul { margin: 0; padding-left: 20px; }
      li { margin-bottom: 8px; line-height: 1.5; }
    </style>

    <div class="header">
      <h2>📊 Expense Analysis Report</h2>
      <div class="date">${results.currentMonth} ${results.currentYear}</div>
    </div>

    <div class="section">
      <div class="section-title">Monthly Comparison</div>
      <div class="metric">
        <span class="metric-label">Current Month Total</span>
        <span class="metric-value">${formatCurrency(results.comparison.current?.totalSpend)}</span>
      </div>`;

    if (results.comparison.vsPrevMonth) {
        const pct = parseFloat(results.comparison.vsPrevMonth.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        const sign = pct > 0 ? '+' : '';
        html += `
      <div class="metric">
        <span class="metric-label">vs ${results.comparison.previousMonthName} ${results.comparison.previousMonthYear}</span>
        <span class="metric-value ${colorClass}">${sign}${results.comparison.vsPrevMonth.percentChange}%</span>
      </div>`;
    }

    if (results.comparison.vsYearAgo) {
        const pct = parseFloat(results.comparison.vsYearAgo.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        const sign = pct > 0 ? '+' : '';
        html += `
      <div class="metric">
        <span class="metric-label">vs ${results.comparison.currentMonthName} ${results.comparison.yearAgoYear} (YoY)</span>
        <span class="metric-value ${colorClass}">${sign}${results.comparison.vsYearAgo.percentChange}%</span>
      </div>`;
    }

    html += `
    </div>

    <div class="section">
      <div class="section-title">Annual Forecast</div>
      <div class="metric">
        <span class="metric-label">YTD Posted (${results.forecast.monthsCompleted} months)</span>
        <span class="metric-value">${formatCurrency(results.forecast.ytdPosted)}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Projected Annual Total</span>
        <span class="metric-value" style="font-size: 18px; color: #667eea;">${formatCurrency(results.forecast.annualForecast)}</span>
      </div>`;

    if (results.forecast.vsLastYear.percentChange) {
        const pct = parseFloat(results.forecast.vsLastYear.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        const sign = pct > 0 ? '+' : '';
        html += `
      <div class="metric">
        <span class="metric-label">vs Last Year (${formatCurrency(results.forecast.previousYearTotal)})</span>
        <span class="metric-value ${colorClass}">${sign}${results.forecast.vsLastYear.percentChange}%</span>
      </div>`;
    }

    html += `
      <div class="metric">
        <span class="metric-label">Monthly Non-Posted Estimate</span>
        <span class="metric-value neutral">${formatCurrency(results.forecast.nonPostedMonthly)}</span>
      </div>
      <div class="metric" style="font-size: 11px; color: #888;">
        <span class="metric-label">Groceries: ${formatCurrency(results.forecast.nonPostedBreakdown.groceries)} | Online: ${formatCurrency(results.forecast.nonPostedBreakdown.onlinePurchases)} | Gas: ${formatCurrency(results.forecast.nonPostedBreakdown.gasoline)}</span>
      </div>
    </div>`;

    // Spike alerts
    if (results.spikes.hasAnomalies) {
        html += `
    <div class="section">
      <div class="section-title">⚠️ Expense Alerts</div>`;

        results.spikes.spikes.forEach(s => {
            html += `
        <div class="spike">
          <strong>${s.category}</strong>: ${formatCurrency(s.currentAmount)} 
          <span style="float: right; color: #D32F2F;">+${s.percentChange}% vs YoY</span>
        </div>`;
        });

        results.spikes.aboveNormal.forEach(s => {
            html += `
        <div class="above-normal">
          <strong>${s.category}</strong>: ${formatCurrency(s.currentAmount)} 
          <span style="float: right; color: #FF9800;">+${s.percentChange}% vs YoY</span>
        </div>`;
        });

        html += `</div>`;
    }

    // AI Insights
    if (results.aiInsights) {
        html += `
    <div class="ai-section">
      <div class="ai-title">🤖 AI Analysis</div>
      <div>${results.aiInsights}</div>
    </div>`;
    }

    html += `
    <div style="text-align: center; font-size: 11px; color: #999; margin-top: 20px;">
      Generated by Expense Analysis Agent
    </div>`;

    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(480)
        .setHeight(650);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Expense Analysis Agent');
}


/**
 * Web App entry point for external calls.
 * Deploy this script as a web app to call from other Apps Scripts.
 * 
 * @param {Object} e - Event object from web request
 * @returns {TextOutput} - JSON response with analysis results
 */
function doGet(e) {
    try {
        const results = runExpenseAnalysisAgent(true);
        return ContentService.createTextOutput(JSON.stringify(results))
            .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}
