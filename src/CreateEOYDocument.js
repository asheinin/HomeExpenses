function createEOYDocument() {

  var myNumbers = new staticNumbers();
  var myUtils = new myUtil();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ss.getName();
  var sheet = ss.getSheets()[0];
  var summarySheet = ss.getSheetByName("Summary")

  var editors = ss.getEditors();

  Logger.log(editors);

  try {


    var currentTime = new Date();
    var month = currentTime.getMonth() + 1;
    var day = currentTime.getDate();
    var currYear = currentTime.getFullYear();

    var fileName = ss.getName();

    var fileYear = fileName.split(" ").slice(-1).pop();

    Logger.log(currYear + " " + " " + fileYear);

    //Check if this is current year file

    if (currYear > fileYear) month = 11;

    var formattedDate = Utilities.formatDate(currentTime, "GMT", "MMMM-dd-yyyy");

    var newFileName = name + " Tax Receipt " + formattedDate;

    //check if file already exists

    //if (DriveApp.getFilesByName(newFileName).hasNext() === true) throw "FileAlreadyExists";  

    //copy sourceSheet from one spreadsheet to another

    //var copyDoc = myUtils.saveFile(newFileName); 
    //if (copyDoc == -1) throw "CopyFailed";

    if ((month != 11) || (month != 1)) {

      var response = myUtils.dialogYN('You may not have all expenses for ' + fileYear + ' yet. Continue?');

      if (response != 'YES') return;

    }

    summaryExpenses();

    var doc = DocumentApp.create(newFileName);
    var docURL = doc.getUrl();

    var files = DriveApp.getFilesByName(newFileName);
    while (files.hasNext()) {
      var file = files.next();
    }

    file.addEditors(editors);

    Logger.log(editors);

    var currentFile = DriveApp.getFileById(ss.getId());
    var parentFold = currentFile.getParents();
    var folder = parentFold.next();
    var theId = folder.getId();
    var targetFolder = DriveApp.getFolderById(theId);
    targetFolder.addFile(file);

    Logger.log('targetFolder name: ' + targetFolder.getName());

    /*
    var currentDoc = DriveApp.getFileById(ss.getId());
    var fileParents = currentDoc.getParents();
    while ( fileParents.hasNext() ) {
       var folder = fileParents.next();
       Logger.log(folder.getName());
    }
    */

    var body = doc.getBody();
    var address = sheet.getRange(myNumbers.dashAddressRow, myNumbers.dashAddressColumn).getValue();
    var rowsData = summarySheet.getRange(1, 1, summarySheet.getLastRow(), 3).getValues();
    Logger.log(rowsData);

    for (var i = 1; i < summarySheet.getLastRow(); i++) {
      rowsData[i][2] = isNaN(parseFloat(rowsData[i][2])) ? "" : "$" + rowsData[i][2].toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,');
      /*
      if (rowsData[i][1] == "") {
        rowsData.splice(i, 1);
        i++
      } 
      */
    }

    body.insertParagraph(0, name + " Tax Receipt ")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.insertParagraph(1, 'Address: ' + address)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(2, 'Print Date: ' + day + "/" + month + "/" + currYear)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    table = body.appendTable(rowsData);
    table.getRow(0).editAsText().setBold(true);

    notifyNewFile(newFileName, docURL);

    // Run historical spending analytics
    runAnalytics();

    // Generate the year-over-year comparison report
    runYearComparison();


  }
  catch (err) {
    if (err == "CopyFailed") {
      ss.toast("No file created", "Error", 5);
      return (-1);
    }
    if (err == "FileAlreadyExists") {
      ss.toast("Next Year File Already Exists", "Error", 5);
      return (-2);
    } else {
      Logger.log(err);
    }
  }

  return;

}
