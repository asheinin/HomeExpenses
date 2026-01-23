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





