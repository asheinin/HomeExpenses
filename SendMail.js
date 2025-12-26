function sendMail(imail, subject,htmlBodyText,noReply,cc) {


  
   MailApp.sendEmail(
       imail,
       subject,
       subject,
          { htmlBody: htmlBodyText,
            cc: cc,
            name: MAILER 
          }
          
         
    );
 

}
