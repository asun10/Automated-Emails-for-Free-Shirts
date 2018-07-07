//Code uses Google Scripts

var EMAIL_SENT = "EMAIL_SENT";

function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 869;   // Number of rows to process
  var dataRange = sheet.getRange(startRow, 1, numRows, 11) //11 is the number of columns in my spreadsheet
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var rowData = data[i];
    var emailAddress = rowData[1];
    var recipient = rowData[0];
    var message1 = rowData[2];
    var message2 = rowData[3];
    var message3 = rowData[4];
    var message4 = rowData[5];
    var message5 = rowData[6];
    var message6 = rowData[7];
    var message7 = rowData[8];
    var message8 = rowData[9];
    //Create personalized message with each college name
    var message = 'Dear ' + recipient + ' admissions' + ',\n\n' + message1 + '\n\n' + message2 + ' ' + recipient + ' ' +
    message3 + '\n\n' + message4 + '\n\n' + message5 + '\n' + message6 + '\n\n'+ message7 + '\n' + message8; 
    var subject = 'Question';
    MailApp.sendEmail(emailAddress, subject, message);
    sheet.getRange(startRow + i, 11).setValue(EMAIL_SENT); //Mark each row as "Email_Sent" once it has been sent
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }

var spreadsheet = SpreadsheetApp.getActive();
var menuItems = [
  {name: 'Send Emails', functionName: 'sendEmails2'}
];
//spreadsheet.addMenu('Send Emails', menuItems);
