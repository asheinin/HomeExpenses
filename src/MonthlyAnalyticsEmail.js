/**
 * Analyzes expenses for the month that just ended and sends an email report.
 * Designed to be triggered on the 1st of each month.
 */
function sendMonthlySummaryEmail() {
  const myNumbers = new staticNumbers();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // 1. Determine the month to analyze (the month that just ended)
  const now = new Date();
  const targetDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const targetMonthIndex = targetDate.getMonth();
  const targetMonthName = months[targetMonthIndex];
  const targetYear = targetDate.getFullYear();

  console.log(`Analyzing month: ${targetMonthName} ${targetYear}`);

  // 2. Get current stats
  const currentSS = getSpreadsheetForYear(targetYear);
  if (!currentSS) {
    console.error(`Could not find spreadsheet for year ${targetYear}`);
    return;
  }
  const currentStats = getMonthStats(currentSS, targetMonthName, targetYear, myNumbers);

  if (!currentStats || currentStats.totalSpend === 0) {
    console.log(`No data found for ${targetMonthName} ${targetYear}. Skipping email.`);
    return;
  }

  // 3. Get previous month stats (for comparison)
  const prevMonthDate = new Date(targetYear, targetMonthIndex - 1, 1);
  const prevMonthName = months[prevMonthDate.getMonth()];
  const prevMonthYear = prevMonthDate.getFullYear();
  const prevSS = (prevMonthYear === targetYear) ? currentSS : getSpreadsheetForYear(prevMonthYear);
  const prevStats = prevSS ? getMonthStats(prevSS, prevMonthName, prevMonthYear, myNumbers) : null;

  // 4. Get same month last year stats (for YoY comparison)
  const yearAgoYear = targetYear - 1;
  const yearAgoSS = getSpreadsheetForYear(yearAgoYear);
  const yearAgoStats = yearAgoSS ? getMonthStats(yearAgoSS, targetMonthName, yearAgoYear, myNumbers) : null;

  // 5. Build HTML Email Body
  const htmlBody = buildInsightsEmailHtml(targetMonthName, targetYear, currentStats, prevStats, yearAgoStats);

  // 6. Get recipients from Dashboard
  const dashboard = currentSS.getSheets()[0];
  const spouse1Email = dashboard.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse1NameColumn).getValue();
  const spouse2Email = dashboard.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse2NameColumn).getValue();
  const spouse1Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
  const spouse2Name = dashboard.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();

  if (!spouse1Email && !spouse2Email) {
    console.error("No email addresses found in Dashboard.");
    return;
  }

  const subject = `Monthly Expense Insights: ${targetMonthName} ${targetYear}`;
  const recipients = [spouse1Email, spouse2Email].filter(e => e).join(',');

  // 7. Send the email
  sendMail(recipients, subject, htmlBody, true, "");
  console.log(`Email sent to: ${recipients}`);
}

/**
 * Extracts statistics for a specific month from a spreadsheet.
 */
function getMonthStats(ss, monthName, year, myNumbers) {
  const sheetName = `${monthName} ${year}`;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  const lastRow = myNumbers.expenseLastRow;
  const firstRow = myNumbers.expenseFirstRow;
  const data = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, myNumbers.expenseAmountColumn).getValues();

  let totalSpend = 0;
  let highestSpend = { description: "None", amount: 0, category: "" };
  const categoryTotals = {};

  data.forEach(row => {
    const category = (row[myNumbers.expenseTypeColumn - 1] || "Uncategorized").toString().trim();
    const description = (row[myNumbers.expenseDescrColumn - 1] || "").toString().trim();
    const amount = parseFloat(row[myNumbers.expenseAmountColumn - 1]) || 0;

    if (amount > 0) {
      totalSpend += amount;

      // Update highest spend
      if (amount > highestSpend.amount) {
        highestSpend = { description, amount, category };
      }

      // Update category totals
      categoryTotals[category] = (categoryTotals[category] || 0) + amount;
    }
  });

  // Find top category
  let topCategory = { name: "None", amount: 0 };
  for (const cat in categoryTotals) {
    if (categoryTotals[cat] > topCategory.amount) {
      topCategory = { name: cat, amount: categoryTotals[cat] };
    }
  }

  return {
    totalSpend,
    highestSpend,
    topCategory,
    categoryTotals
  };
}

/**
 * Helper to find the spreadsheet for a given year.
 */
function getSpreadsheetForYear(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getName().includes(year.toString())) return ss;

  const files = DriveApp.searchFiles(`title = "Home payments ${year}" and mimeType = "${MimeType.GOOGLE_SHEETS}"`);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return null;
}

/**
 * Constructs the HTML body for the insights email.
 */
function buildInsightsEmailHtml(month, year, current, prev, yearAgo) {
  const formatCurrency = (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'CAD' }).format(val);

  let html = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #eee; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #4CAF50; color: white; padding: 20px; text-align: center;">
        <h1 style="margin: 0; font-size: 24px;">${month} ${year} Expense Summary</h1>
      </div>
      
      <div style="padding: 20px;">
        <div style="background-color: #f9f9f9; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
          <small style="color: #666; text-transform: uppercase; letter-spacing: 1px;">Total Spend</small>
          <div style="font-size: 32px; font-weight: bold; color: #2E7D32; margin: 5px 0;">${formatCurrency(current.totalSpend)}</div>
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
          <tr>
            <td style="padding: 10px; border-bottom: 1px solid #eee; width: 50%;">
              <strong style="display: block; color: #666; font-size: 12px; text-transform: uppercase;">Highest Spend</strong>
              <span style="font-size: 16px;">${current.highestSpend.description}</span>
            </td>
            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right;">
              <span style="font-size: 18px; font-weight: bold; color: #D32F2F;">${formatCurrency(current.highestSpend.amount)}</span>
            </td>
          </tr>
          <tr>
            <td style="padding: 10px; border-bottom: 1px solid #eee;">
              <strong style="display: block; color: #666; font-size: 12px; text-transform: uppercase;">Top Category</strong>
              <span style="font-size: 16px;">${current.topCategory.name}</span>
            </td>
            <td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right;">
              <span style="font-size: 18px; font-weight: bold; color: #1976D2;">${formatCurrency(current.topCategory.amount)}</span>
            </td>
          </tr>
        </table>

        <h3 style="color: #444; margin-top: 30px; border-left: 4px solid #4CAF50; padding-left: 10px;">Comparisons</h3>
        <table style="width: 100%; border-collapse: collapse;">
          <thead>
            <tr style="background-color: #f5f5f5;">
              <th style="padding: 10px; text-align: left; font-size: 12px; color: #888;">PERIOD</th>
              <th style="padding: 10px; text-align: right; font-size: 12px; color: #888;">TOTAL</th>
              <th style="padding: 10px; text-align: right; font-size: 12px; color: #888;">CHANGE</th>
            </tr>
          </thead>
          <tbody>`;

  // Previous Month Comparison
  if (prev) {
    const diff = current.totalSpend - prev.totalSpend;
    const pct = ((diff / prev.totalSpend) * 100).toFixed(1);
    const color = diff > 0 ? '#D32F2F' : '#2E7D32';
    const sign = diff > 0 ? '+' : '';
    html += `
      <tr>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee;">Previous Month</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee; text-align: right;">${formatCurrency(prev.totalSpend)}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee; text-align: right; color: ${color}; font-weight: bold;">${sign}${pct}%</td>
      </tr>`;
  }

  // Year Ago Comparison
  if (yearAgo) {
    const diff = current.totalSpend - yearAgo.totalSpend;
    const pct = ((diff / yearAgo.totalSpend) * 100).toFixed(1);
    const color = diff > 0 ? '#D32F2F' : '#2E7D32';
    const sign = diff > 0 ? '+' : '';
    html += `
      <tr>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee;">Same Month Last Year</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee; text-align: right;">${formatCurrency(yearAgo.totalSpend)}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #eee; text-align: right; color: ${color}; font-weight: bold;">${sign}${pct}%</td>
      </tr>`;
  }

  html += `
          </tbody>
        </table>

        <!-- Insights Section -->
        ${generateInsights(current, prev, yearAgo)}
      </div>

      <div style="background-color: #f5f5f5; color: #888; padding: 15px; text-align: center; font-size: 12px;">
        This is an automated report from your Home Expenses system.
      </div>
    </div>
  `;

  return html;
}

/**
 * Generates textual insights based on data, using AI if available.
 */
function generateInsights(current, prev, yearAgo) {
  // Try AI insights first
  const aiInsights = generateAiInsights(current, prev, yearAgo);
  if (aiInsights) {
    return `
      <div style="margin-top: 30px; background-color: #E8F5E9; padding: 15px; border-radius: 8px;">
        <h4 style="margin: 0 0 10px 0; color: #2E7D32;">AI Monthly Analysis</h4>
        <div style="color: #444; line-height: 1.6;">
          ${aiInsights}
        </div>
      </div>
    `;
  }

  // Fallback to original hardcoded logic
  const insights = [];

  // Insight 1: Highest category proportion
  if (current.totalSpend > 0) {
    const topPct = ((current.topCategory.amount / current.totalSpend) * 100).toFixed(0);
    insights.push(`Your top category **${current.topCategory.name}** accounted for **${topPct}%** of your total spending.`);
  }

  // Insight 2: YoY Comparison
  if (yearAgo) {
    const diff = current.totalSpend - yearAgo.totalSpend;
    if (diff < 0) {
      insights.push(`Great news! You spent **${new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(Math.abs(diff))} less** than the same month last year.`);
    } else {
      insights.push(`You spent **${new Intl.NumberFormat('en-CA', { style: 'currency', currency: 'CAD' }).format(diff)} more** than the same month last year.`);
    }
  }

  // Insight 3: Large single expense
  if (current.highestSpend.amount > current.totalSpend * 0.3) {
    insights.push(`A single expense (**${current.highestSpend.description}**) was responsible for over **30%** of your monthly total.`);
  }

  if (insights.length === 0) return "";

  return `
    <div style="margin-top: 30px; background-color: #E8F5E9; padding: 15px; border-radius: 8px;">
      <h4 style="margin: 0 0 10px 0; color: #2E7D32;">Monthly Insights</h4>
      <ul style="margin: 0; padding-left: 20px; color: #444; line-height: 1.6;">
        ${insights.map(i => `<li style="margin-bottom: 5px;">${i}</li>`).join('')}
      </ul>
    </div>
  `;
}

/**
 * Calls Gemini AI to generate personalized insights.
 */
function generateAiInsights(current, prev, yearAgo) {
  const formatCurrency = (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'CAD' }).format(val);

  let prompt = `Analyze these household expenses and provide 3-4 concise, helpful insights in HTML list format (<ul><li>...</li></ul>). 
  Be encouraging and highlight trends. Use <strong> for emphasis.

  CONTEXT:
  - This Month Total Spend: ${formatCurrency(current.totalSpend)}
  - Top Spending Category: ${current.topCategory.name} (${formatCurrency(current.topCategory.amount)})
  - Largest Single Expense: ${current.highestSpend.description} (${formatCurrency(current.highestSpend.amount)})
  
  CATEGORY BREAKDOWN:
  ${Object.entries(current.categoryTotals).map(([cat, amt]) => `- ${cat}: ${formatCurrency(amt)}`).join('\n')}
  `;

  if (prev) {
    prompt += `\nCOMPARISON WITH PREVIOUS MONTH:
    - Previous Month Total: ${formatCurrency(prev.totalSpend)}
    - Change: ${formatCurrency(current.totalSpend - prev.totalSpend)} (${(((current.totalSpend - prev.totalSpend) / prev.totalSpend) * 100).toFixed(1)}%)`;
  }

  if (yearAgo) {
    prompt += `\nCOMPARISON WITH SAME MONTH LAST YEAR:
    - Last Year Total: ${formatCurrency(yearAgo.totalSpend)}
    - Change: ${formatCurrency(current.totalSpend - yearAgo.totalSpend)} (${(((current.totalSpend - yearAgo.totalSpend) / yearAgo.totalSpend) * 100).toFixed(1)}%)`;
  }

  const response = callGemini(prompt);
  return response ? response.trim() : null;
}
