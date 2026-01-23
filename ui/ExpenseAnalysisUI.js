/**
 * Expense Analysis UI
 * 
 * Handles display and interaction logic for the Expense Analysis Agent.
 * Includes web app entry points (doGet, doPost) and HTML generation.
 * 
 * @author Alexander Sheinin
 */

/**
 * Displays agent results in a formatted HTML dialog.
 * (Interactive version for spreadsheet use)
 */
function displayAgentResults(results, myNumbers) {
    const html = generateAgentHtml(results, { isReadOnly: false, isStandalone: false, isEmbedded: false });

    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(500)
        .setHeight(650);

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
    <title>Expense Analysis - \${results.currentMonth} \${results.currentYear}</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: \${isEmbedded ? 'transparent' : (isStandalone ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' : '#fff')};
            min-height: 100vh;
            padding: \${isEmbedded ? '0' : (isStandalone ? '10px' : '0')};
            color: #333;
            overflow-x: hidden;
        }
        .container { 
            max-width: 600px; 
            margin: 0 auto;
            position: relative;
            background: \${isStandalone || isEmbedded ? 'transparent' : '#fff'};
        }
        
        .header {
            background: \${isStandalone || isEmbedded ? 'rgba(255,255,255,0.95)' : 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'};
            color: \${isStandalone || isEmbedded ? '#333' : 'white'};
            border-radius: \${isStandalone || isEmbedded ? '12px' : '0 0 12px 12px'};
            padding: 16px;
            text-align: center;
            margin-bottom: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            position: relative;
            \${isEmbedded ? 'margin-top: 5px;' : ''}
        }
        .header h1 { font-size: 22px; margin-bottom: 2px; }
        .header .date { color: \${isStandalone || isEmbedded ? '#667eea' : 'rgba(255,255,255,0.9)'}; font-weight: 600; font-size: 15px; }
        
        .menu-btn {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            width: 38px;
            height: 38px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            font-size: 24px;
            color: \${isStandalone || isEmbedded ? '#667eea' : 'white'};
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
            font-size: 16px; font-weight: 700; color: #555;
            margin-bottom: 10px; padding-bottom: 5px;
            border-bottom: 2px solid #667eea;
        }
        
        .metric {
            display: flex; justify-content: space-between; align-items: center;
            padding: 10px 0; border-bottom: 1px solid #f9f9f9;
        }
        .metric:last-child { border-bottom: none; }
        .metric-label { color: #666; font-size: 14px; }
        .metric-value { font-weight: 700; font-size: 16px; }
        .metric-value.big { font-size: 22px; color: #667eea; }
        .positive { color: #2E7D32; }
        .negative { color: #D32F2F; }
        
        .spike { 
            background: #FFEBEE; border-left: 4px solid #D32F2F; 
            padding: 10px; margin: 6px 0; border-radius: 0 6px 6px 0;
            display: flex; justify-content: space-between; font-size: 14px;
        }
        .above-normal { 
            background: #FFF3E0; border-left: 4px solid #FF9800; 
            padding: 10px; margin: 6px 0; border-radius: 0 6px 6px 0;
            display: flex; justify-content: space-between; font-size: 14px;
        }
        
        .ai-section {
            background: #E8F5E9; border-radius: 12px;
            padding: 16px; margin-bottom: 10px; border: 1px solid #C8E6C9;
        }
        .ai-title { font-size: 18px; font-weight: 700; color: #2E7D32; margin-bottom: 12px; }
        .ai-section ul { padding-left: 20px; }
        .ai-section li { margin-bottom: 10px; line-height: 1.5; font-size: 16px; }
        
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
            padding: 24px;
            width: 100%;
            max-width: 440px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            position: relative;
        }
        .modal-close {
            position: absolute;
            right: 15px;
            top: 15px;
            font-size: 28px;
            cursor: pointer;
            color: #999;
        }
        
        .assumptions-title { font-size: 18px; font-weight: 700; color: #F57C00; margin-bottom: 15px; border-bottom: 2px solid #FFE082; padding-bottom: 8px; }
        .item-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; font-size: 15px; border-bottom: 1px solid #f9f9f9; }
        .item-label { color: #666; }
        .item-value { font-weight: 600; }
        
        .total-row {
            display: flex; justify-content: space-between; align-items: center;
            padding: 14px 0; margin-top: 12px; border-top: 2px solid #FFB74D;
            font-weight: 700; color: #F57C00; font-size: 18px;
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
            text-align: center; color: \${isStandalone || isEmbedded ? 'rgba(0,0,0,0.4)' : '#aaa'};
            font-size: 10px; padding: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Expense Analysis</h1>
            <div class="date">\${results.currentMonth} \${results.currentYear}</div>
            <div class="menu-btn" onclick="toggleAssumptions(true)" title="Forecast Assumptions">⋮</div>
        </div>
        
        <div class="card">
            <div class="card-title">Monthly Comparison</div>
            <div class="metric">
                <span class="metric-label">Current Month Total</span>
                <span class="metric-value">\${formatCurrency(results.comparison.current?.totalSpend)}</span>
            </div>\`;

    if (results.comparison.vsPrevMonth) {
        const pct = parseFloat(results.comparison.vsPrevMonth.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += \`
            <div class="metric">
                <span class="metric-label">vs \${results.comparison.previousMonthName}</span>
                <span class="metric-value \${colorClass}">\${pct > 0 ? '+' : ''}\${pct}%</span>
            </div>\`;
    }

    if (results.comparison.vsYearAgo) {
        const pct = parseFloat(results.comparison.vsYearAgo.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += \`
            <div class="metric">
                <span class="metric-label">vs \${results.currentMonth} \${results.comparison.yearAgoYear} (YoY)</span>
                <span class="metric-value \${colorClass}">\${pct > 0 ? '+' : ''}\${pct}%</span>
            </div>\`;
    }

    html += \`
        </div>
        
        <div class="card">
            <div class="card-title">Annual Forecast</div>
            <div class="metric">
                <span class="metric-label">YTD Posted (\${results.forecast.monthsCompleted} mo)</span>
                <span class="metric-value">\${formatCurrency(results.forecast.ytdPosted)}</span>
            </div>
            <div class="metric">
                <span class="metric-label">Projected Annual Total</span>
                <span class="metric-value big">\${formatCurrency(results.forecast.annualForecast)}</span>
            </div>\`;

    if (results.forecast.vsLastYear.percentChange) {
        const pct = parseFloat(results.forecast.vsLastYear.percentChange);
        const colorClass = pct > 0 ? 'negative' : 'positive';
        html += \`
            <div class="metric">
                <span class="metric-label">vs Last Year (\${formatCurrency(results.forecast.previousYearTotal)})</span>
                <span class="metric-value \${colorClass}">\${pct > 0 ? '+' : ''}\${pct}%</span>
            </div>\`;
    }

    html += \`
        </div>\`;

    if (results.spikes.hasAnomalies) {
        html += \`
        <div class="card">
            <div class="card-title">⚠️ Expense Alerts</div>\`;
        results.spikes.spikes.forEach(s => {
            html += \`<div class="spike"><strong>\${s.category}</strong><span class="negative">+\${s.percentChange}%</span></div>\`;
        });
        results.spikes.aboveNormal.forEach(s => {
            html += \`<div class="above-normal"><strong>\${s.category}</strong><span class="negative">+\${s.percentChange}%</span></div>\`;
        });
        html += \`</div>\`;
    }

    if (results.aiInsights) {
        html += \`<div class="ai-section"><div class="ai-title">🤖 AI Analysis</div><div>\${results.aiInsights}</div></div>\`;
    }

    /* Modal for Assumptions */
    html += \`
    <div class="modal-overlay" id="modalOverlay" onclick="if(event.target==this) toggleAssumptions(false)">
        <div class="modal-content">
            <div class="modal-close" onclick="toggleAssumptions(false)">×</div>
            <div class="assumptions-title">📝 Forecast Assumptions (Monthly)</div>\`;

    const categories = [
        { id: 'groceries', label: 'Groceries', val: results.forecast.nonPostedBreakdown.groceries },
        { id: 'onlinePurchases', label: 'Online Purchases', val: results.forecast.nonPostedBreakdown.onlinePurchases },
        { id: 'gasoline', label: 'Gasoline', val: results.forecast.nonPostedBreakdown.gasoline },
        { id: 'misc', label: 'Miscellaneous', val: results.forecast.nonPostedBreakdown.misc }
    ];

    categories.forEach(cat => {
        html += \`<div class="item-row">
            <span class="item-label">\${cat.label}</span>
            <span class="item-value">
                \${isReadOnly
                ? formatCurrency(cat.val)
                : \`$ <input type="number" id="\${cat.id}" value="\${cat.val}" min="0" step="50" oninput="updateTotal()">\`}
            </span>
        </div>\`;
    });

    html += \`
            <div class="total-row">
                <span>Total Monthly Non-Posted</span>
                <span id="totalMonthly">\${formatCurrency(results.forecast.nonPostedMonthly)}</span>
            </div>\`;

    if (!isReadOnly) {
        html += \`<button class="recalculate-btn" id="recalcBtn" onclick="recalculateForecast()">🔄 Recalculate Forecast</button>\`;
    }

    html += \`
        </div>
    </div>
    
    <div class="footer">Generated by Expense Analysis Agent</div>
    </div>\`;

    html += \`
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
                    // If in a Web App context, refreshing the page shows the updated forecast results
                    if (window.location.href.indexOf('macros/s/') !== -1) {
                        window.top.location.reload();
                    }
                })
                .withFailureHandler(function(err) {
                    alert('Error: ' + err);
                    btn.disabled = false;
                    btn.textContent = originalText;
                })
                .recalculateWithAssumptions(groceries, online, gasoline, misc);
        }
    </script>
    </body></html>\`;
    return html;
}
