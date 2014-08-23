function sendFeedbackEmails() {
  var EMAIL_SENT = "EMAIL_SENT";
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 4;  // First row of data to process
  var numRows = sheet.getLastRow() - 3;   // Number of rows to process
  // Fetch the range of cells A2:A:10
  var boilerplateRange = sheet.getRange(2, 1, 1, 5);
  var dataRange = sheet.getRange(startRow, 1, numRows, 11);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var boilerplate = boilerplateRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[2];  // Third column
    // Construct message
    var timestamp = Utilities.formatDate(new Date(), "CST", "HH:mm:ss");
    var t = HtmlService.createTemplateFromFile('tmpl');
    t.name =         row[0];
    t.introMessage = boilerplate[0][2];
    t.stationOne =   row[3];
    t.runCadence =   row[4];
    t.minnMarch =    row[5];
    t.battleHymn =   row[6];
    t.stationThree = row[7];
    t.ourMinn =      row[8];
    t.generalComments = row[9];
    t.closingMessage = boilerplate[0][3];
    t.signature = boilerplate[0][4];
    t.timestamp = timestamp;
    var message = t.evaluate().getContent();
    var formattedDate = Utilities.formatDate(boilerplate[0][1], "CST", "MM-dd-yyyy");
    var subject = boilerplate[0][0] + " " + formattedDate;
    var emailSent = row[10];     // Last Column
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      MailApp.sendEmail(emailAddress, subject, message, {htmlBody: message});
      sheet.getRange(startRow + i, 11).setValue(EMAIL_SENT);
      SpreadsheetApp.flush();
      Logger.log("Sent message to %s", row[0]);
    } else {
      Logger.log("Skipping message to %s, already sent.", row[0]);
    }
  }
}
