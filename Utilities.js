// Google Spreadsheets Data API
// Author - Alexander Sheinin


//######################################################################

/**
 * utility class with public methods and fields
 * <pre>
 * this.stringTrim(); 
 * this.saveFile(); 
 * </pre>
 */


function myUtil() {
   this.stringTrim(); 
   this.saveFile();
   this.dialogYN();
   this.dialogQA();
}


myUtil.prototype.stringTrim  = function(str) {
      str = str + " ";
      return str.replace(/\s*/g, '');
}




myUtil.prototype.saveFile = function(name) {
  
  var myNumbers = new staticNumbers();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  
  try {      
    
         if (!name) throw "CopyFailed";
    
         //copy sourceSheet from one spreadsheet to another
 
         var currentDoc = DriveApp.getFileById(ss.getId());
     
         //target doc
         var copyDoc = currentDoc.makeCopy(name);
         if (!copyDoc) throw "CopyFailed";
     } 
           
     catch(err)  
       {   
       
          if (err == "CopyFailed")
            return(-1);
         
         if (err == "FileExists") {
            return(-2);
         } else {
           return(-1);
         }  
       
       }
        
     return(copyDoc);
} 


myUtil.prototype.dialogYN = function(question) {
  
 if (!question) return -1;
  
 // Display a dialog box with a message and "Yes" and "No" buttons.
 var ui = SpreadsheetApp.getUi();
 var response = ui.alert(question, ui.ButtonSet.YES_NO);

 // Process the user's response.
 if (response == ui.Button.YES) {
   Logger.log('The user clicked "Yes."');
 } else {
   Logger.log('The user clicked "No" or the dialog\'s close button.');
 }
  
 return (response); 
  
}



myUtil.prototype.dialogQA = function(title, question) {

 if (!title) return -1; 

 var ui = SpreadsheetApp.getUi();
 var response = ui.prompt(title, question, ui.ButtonSet.OK_CANCEL);

 // Process the user's response.
 if (response.getSelectedButton() == ui.Button.OK) {
   Logger.log('The expense is %s.', response.getResponseText());
 } else if (response.getSelectedButton() == ui.Button.CANCEL) {
   Logger.log('The user clicked Cancel');
   return -1
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
   return -1
 }
 
 return (response); 
  
}  

myUtil.prototype.dialogOK = function(question) {
  
 if (!question) return -1;
  
 // Display a dialog box with a message and "Yes" and "No" buttons.
 var ui = SpreadsheetApp.getUi();
 var response = ui.alert(question, ui.ButtonSet.OK);

 // Process the user's response.

 if (response == ui.Button.OK) {
   Logger.log('The user clicked "OK"');
 } else {
   Logger.log('The user clicked the dialog\'s close button.');
 }
  
 return (response); 
  
}


