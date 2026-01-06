function findAmount() {
  
  var myNumbers = new staticNumbers();
  
  var emailDate = new Date();

  var formattedDate = Utilities.formatDate(emailDate, "GMT", "MMMM yyyy");
  var name = ""
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  ss.setActiveSheet(sheet);
  
  var date = new Date();
  var currentMonth = date.getMonth();

  var currYear = date.getFullYear();
 
  var fileName = ss.getName();
  
  var fileYear = fileName.split(" ").slice(-1).pop(); 
 
  var year = "curr";
  
  //year can be curr or next. For curr year, process all except for 0 month. For next year process only if month is 0. For all other years return
  
  if ((currYear == fileYear)&&(currentMonth == 0)) return;
  if (currYear < fileYear) return;
  if (currYear-fileYear > 1) return;
  if (currYear-fileYear == 1) {
    if (currentMonth == 0) {
      year = "next";
    } else {
      return;
    }  
  }    
  
  Logger.log("current month: " + currentMonth + " " + year);
  
  var sp1Name = sheet.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse1NameColumn).getValue();
  var sp2Name = sheet.getRange(myNumbers.dashNamesRow, myNumbers.dashSpouse2NameColumn).getValue();
  
  var sp1Email = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse1NameColumn).getValue();
  var sp2Email = sheet.getRange(myNumbers.dashEmailsRow, myNumbers.dashSpouse2NameColumn).getValue();
  
  var dataRange = sheet.getRange(myNumbers.dashFirstMonthRow, 1, 12, myNumbers.dashColumns);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  var match = false;
  
  for (var i=1; i<data.length; i++) {
   
    var formattedDate1 = Utilities.formatDate(data[i][0], "GMT", "MMMM yyyy");
    
    //Logger.log(formattedDate1 + " "  + formattedDate + " "  + year);
    
    if ((formattedDate1 == formattedDate)||(year == "next")){
       match = true;
       if (year != "next") {
          var formattedDateMonthAgo = Utilities.formatDate(data[i-1][0], "GMT", "MMMM yyyy");
       } else {
          var formattedDateMonthAgo = Utilities.formatDate(data[data.length][0], "GMT", "MMMM yyyy");
          i = data.length -1;
       }  
        
       var amountSp1Owes = data[i-1][myNumbers.dashSp1BalanceColumn-1];
       var amountSp2Owes = data[i-1][myNumbers.dashSp2BalanceColumn-1];
       if (amountSp1Owes>0) {
         name = sp1Name; 
         var amount = amountSp1Owes;
       } else { 
         name = sp2Name; 
         var amount = amountSp2Owes;
       }
      
       console.log("data: " + i + " " + myNumbers.dashAmountTotalColumn);
      // console.log("data: "  + amountSp1Owes + " " + data);
       console.log("data: " + (data[i-1][myNumbers.dashSp2BalanceColumn - 1]).toFixed(2));
       
       var _test = i-1;
       var _test1 = myNumbers.dashAmountTotalColumn-1;

       console.log(data[i-1]);

       var totalMonth = data[i-1][myNumbers.dashAmountTotalColumn-1].toFixed(2);
       var sp1Part = data[i-1][myNumbers.dashSp1PartColumn-1].toFixed(2);
       var sp2Part = data[i-1][myNumbers.dashSp2PartColumn-1].toFixed(2);
       var sp1Paid = data[i-1][myNumbers.dashSp1PaidColumn-1].toFixed(2);
       var sp2Paid = data[i-1][myNumbers.dashSp2PaidColumn-1].toFixed(2);
       var sp1ToSp2 = data[i-1][myNumbers.dashSp1ToSp2Column-1].toFixed(2);
       var sp2ToSp1 = data[i-1][myNumbers.dashSp2ToSp1Column-1].toFixed(2);
       var sp1SBalanceUsed = (data[i-1][myNumbers.dashSp1BalanceUsedColumn-1] == "") ? 0 : data[i-1][myNumbers.dashSp1BalanceUsedColumn-1].toFixed(2);
       var sp2SBalanceUsed = (data[i-1][myNumbers.dashSp2BalanceUsedColumn-1] == "") ? 0 : data[i-1][myNumbers.dashSp2BalanceUsedColumn-1].toFixed(2);
       var totalSp1Paid = parseFloat(sp1Paid) + parseFloat(sp1ToSp2) + parseFloat(sp1SBalanceUsed) - parseFloat(sp2ToSp1);
       var totalSp2Paid = parseFloat(sp2Paid) + parseFloat(sp2ToSp1) + parseFloat(sp2SBalanceUsed) - parseFloat(sp1ToSp2);
    }
  }
  
  if (!match) return;
  
  var noReply = true;
  var subject = "Monthly Property Account for " + formattedDateMonthAgo;
  var htmlBodyText = "<br>In " + formattedDateMonthAgo + "<br>";
  if (amount < 10) {
    htmlBodyText += "<br><strong> All Paid </strong><br>";
  } else {  
     htmlBodyText += "<br><strong>" + name + " " + "to pay $" + Math.round(amount).toFixed(2) + "</strong><br>";
  }  
  htmlBodyText += "<br>----------------------------------------------<br>";
  htmlBodyText += "<br>" + "Total amount: $" + " " + totalMonth + "<br>";
  htmlBodyText += "<br>" + sp1Name + "'s part: $" + " " + sp1Part + "<br>";
  htmlBodyText += "<br>" + sp2Name + "'s part: $" + " " + sp2Part + "<br>";
  htmlBodyText += "<br>" + sp1Name + " paid: $" + " " + totalSp1Paid + "<br>";
  htmlBodyText += "<br>" + sp2Name + " paid: $" + " " + totalSp2Paid + "<br>";
  htmlBodyText += "<br>----------------------------------------------<br>";
  
 
  Logger.log(sp1Email + " " + sp2Email);
  
 // sendMail(sp1Email, subject,htmlBodyText,noReply);
  if (sp1Email != sp2Email) sendMail(sp2Email, subject,htmlBodyText,noReply); 
  
  return;
}


