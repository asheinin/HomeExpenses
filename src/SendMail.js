function sendMail(imail, subject,htmlBodyText,noReply,cc) {
  var myNumbers = new staticNumbers();

  MailApp.sendEmail(
       imail,
       subject,
       subject,
          { htmlBody: htmlBodyText,
            cc: cc,
            name: myNumbers.mailerName
          }
    );
}
