function createEOYDocument() {
  var htmlDlg = HtmlService.createHtmlOutputFromFile('ui/TaxReceiptDialog')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi()
    .showModalDialog(htmlDlg, 'Tax Receipt Dates');
}

function processEOYDocument(startDateStr, endDateStr) {
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
    var startDate = new Date(startDateStr + "T00:00:00");
    var endDate = new Date(endDateStr + "T00:00:00");
    
    // Create list of target months
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const targetSheets = [];
    let d = new Date(startDate.getTime());
    while (d <= endDate) {
      targetSheets.push({ monthStr: months[d.getMonth()], year: d.getFullYear() });
      d.setMonth(d.getMonth() + 1);
    }
    
    // Check if we need to open a previous year spreadsheet
    let prevSS = null;
    const yearsNeeded = Array.from(new Set(targetSheets.map(ts => ts.year)));
    if (yearsNeeded.some(y => y !== activeYear)) {
      const prevYear = yearsNeeded.find(y => y !== activeYear);
      const files = DriveApp.searchFiles('title = "Home payments ' + prevYear + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');
      if (files.hasNext()) {
        prevSS = SpreadsheetApp.open(files.next());
      }
    }
    
    // Aggregate Data
    const data = {};
    targetSheets.forEach(ts => {
      let sheetSS = (ts.year === activeYear) ? ss : (prevSS || ss);
      const sheetName = `${ts.monthStr} ${ts.year}`;
      const sheet = sheetSS.getSheetByName(sheetName);
      if (!sheet) return;

      const range = sheet.getRange('A2:D50');
      const values = range.getValues();

      values.forEach(row => {
        const type = row[myNumbers.expenseTypeColumn - 1];
        const description = row[myNumbers.expenseDescrColumn - 1];
        const amount = row[myNumbers.expenseAmountColumn - 1];

        if (!type || !amount) return;

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
    body.insertParagraph(3, 'Period: ' + startDateStr + ' to ' + endDateStr)
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
    Logger.log(err);
    ss.toast("Error generating document: " + err, "Error", 5);
  }
}

