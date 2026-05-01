function createEOYDocument() {
  var template = HtmlService.createTemplateFromFile('ui/TaxReceiptDialog');
  
  var availableYears = [];
  var files = DriveApp.searchFiles('title contains "Home payments" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');
  while (files.hasNext()) {
    var file = files.next();
    var match = file.getName().match(/Home payments\s*(\d{4})/i);
    if (match) {
      availableYears.push(parseInt(match[1]));
    }
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeYearMatch = ss.getName().match(/Home payments\s*(\d{4})/i);
  var activeYear = activeYearMatch ? parseInt(activeYearMatch[1]) : new Date().getFullYear();
  if (availableYears.indexOf(activeYear) === -1) {
    availableYears.push(activeYear);
  }
  
  availableYears.sort(function(a, b){return a-b});
  template.availableYears = availableYears;

  var htmlDlg = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi()
    .showModalDialog(htmlDlg, 'Tax Receipt Dates');
}

function processEOYDocument(startMonthStr, endMonthStr) {
  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ss.getName();
  var dashSheet = ss.getSheets()[0];
  var editors = ss.getEditors();

  try {
    var currentTime = new Date();
    var formattedDate = Utilities.formatDate(currentTime, "GMT", "MMMM-dd-yyyy");
    
    var fileName = ss.getName();
    var activeYear = parseInt(fileName.split(" ").slice(-1).pop()) || currentTime.getFullYear();
    
    // Parse dates
    var startDate = new Date(startMonthStr + "-01T00:00:00");
    var endDate = new Date(endMonthStr + "-01T00:00:00");
    
    // Calculate display period
    var displayStartDate = Utilities.formatDate(startDate, "GMT", "MMMM 1, yyyy");
    var endDateObj = new Date(endDate.getFullYear(), endDate.getMonth() + 1, 0); // Last day of month
    var displayEndDate = Utilities.formatDate(endDateObj, "GMT", "MMMM d, yyyy");
    
    // Create list of target months
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const targetSheets = [];
    let d = new Date(startDate.getTime());
    while (d <= endDate) {
      targetSheets.push({ monthStr: months[d.getMonth()], year: d.getFullYear() });
      d.setMonth(d.getMonth() + 1);
    }
    
    // Check if required spreadsheets exist
    const yearsNeeded = Array.from(new Set(targetSheets.map(ts => ts.year)));
    const spreadsheetsByYear = {};
    
    for (let i = 0; i < yearsNeeded.length; i++) {
      let year = yearsNeeded[i];
      if (year === activeYear) {
        spreadsheetsByYear[year] = ss;
      } else {
        const files = DriveApp.searchFiles('title = "Home payments ' + year + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');
        if (files.hasNext()) {
          spreadsheetsByYear[year] = SpreadsheetApp.open(files.next());
        } else {
          throw new Error("Spreadsheet for year " + year + " ('Home payments " + year + "') does not exist. Please adjust the dates.");
        }
      }
    }
    
    // Aggregate Data
    const data = {};
    targetSheets.forEach(ts => {
      let sheetSS = spreadsheetsByYear[ts.year];
      const sheetName = `${ts.monthStr} ${ts.year}`;
      
      console.log(`Processing month: ${sheetName} from file: ${sheetSS.getName()}`);
      
      const sheet = sheetSS.getSheetByName(sheetName);
      if (!sheet) return;

      const range = sheet.getRange('A2:D50');
      const values = range.getValues();

      values.forEach(row => {
        const type = row[myNumbers.expenseTypeColumn - 1];
        const description = row[myNumbers.expenseDescrColumn - 1];
        let amount = row[myNumbers.expenseAmountColumn - 1];

        if (!type || !amount) return;

        // Ensure amount is a float
        amount = parseFloat(amount) || 0;

        if (!data[type]) {
          data[type] = {
            descriptions: new Set(),
            totalAmount: 0
          };
        }

        if (description) {
          data[type].descriptions.add(description);
        }
        data[type].totalAmount += amount;
        
        if (type.toString().trim().toLowerCase() === "eat out") {
          console.log(`Eat Out found in ${sheetName}: amount=${amount}, running total=${data[type].totalAmount}`);
        }
      });
    });
    
    // Build Rows Data
    const header = ['Type', 'Description', 'Total Amount'];
    const rowsData = [header];
    let grandTotal = 0;
    
    Object.keys(data).forEach(type => {
        const row = [
            type,
            Array.from(data[type].descriptions).join(', '),
            data[type].totalAmount
        ];
        rowsData.push(row);
        grandTotal += data[type].totalAmount;
    });
    rowsData.push(['Total', '', grandTotal]);
    
    // Format currency
    for (var i = 1; i < rowsData.length; i++) {
      rowsData[i][2] = isNaN(parseFloat(rowsData[i][2])) ? "" : "$" + rowsData[i][2].toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,');
    }

    var newFileName = name + " Tax Receipt " + formattedDate;

    // Call existing summary processes for analytics
    summaryExpenses();

    var doc = DocumentApp.create(newFileName);
    var docURL = doc.getUrl();

    var files = DriveApp.getFilesByName(newFileName);
    var file;
    while (files.hasNext()) {
      file = files.next();
    }

    if (editors && editors.length > 0) {
      file.addEditors(editors);
    }

    var currentFile = DriveApp.getFileById(ss.getId());
    var parentFold = currentFile.getParents();
    if (parentFold.hasNext()) {
      var folder = parentFold.next();
      var theId = folder.getId();
      var targetFolder = DriveApp.getFolderById(theId);
      targetFolder.addFile(file);
    }

    var body = doc.getBody();
    var address = dashSheet.getRange(myNumbers.dashAddressRow, myNumbers.dashAddressColumn).getValue();

    body.insertParagraph(0, name + " Tax Receipt ")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.insertParagraph(1, 'Address: ' + address)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(2, 'Print Date: ' + currentTime.getDate() + "/" + (currentTime.getMonth()+1) + "/" + currentTime.getFullYear())
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.insertParagraph(3, 'Period: ' + displayStartDate + ' to ' + displayEndDate)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
      
    var table = body.appendTable(rowsData);
    table.getRow(0).editAsText().setBold(true);

    notifyNewFile(newFileName, docURL);

    // Run historical spending analytics
    runAnalytics();

    // Generate the year-over-year comparison report
    runYearComparison();

  }
  catch (err) {
    console.error(err);
    ss.toast("Error generating document: " + err, "Error", 5);
  }
}

