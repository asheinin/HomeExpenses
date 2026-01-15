function notifyNewFile(fileName,fileURL,notes) {
  
    if (!notes) notes = " ";
  
    var myNumbers = new staticNumbers();
    var myUtils = new myUtil();
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    Logger.log(fileName);
    
    var subject = 'Your File '+ fileName + ' is ready'
    var htmlBodyText = "<br><strong>File is ready:  </strong><br>";
    htmlBodyText += "<br>" + fileName + "<br>";
    htmlBodyText += '<br><a href=' + fileURL + ' style="font-size:150%"> ' + fileName + ' </a>'+'<br><br>';
    htmlBodyText += notes;
    
    SpreadsheetApp.getActiveSpreadsheet().toast('File is ready: ' + fileName, 'File is ready', 3);
    
    
    
    var sp1Email = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse1NameColumn).getValue();
    var sp2Email = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse2NameColumn).getValue();
    
    var noReply = true;
    
    Logger.log(sp1Email);
  
    sendMail(sp1Email, subject,htmlBodyText,noReply);
    if (sp1Email != sp2Email) sendMail(sp2Email, subject,htmlBodyText,noReply); 
}
