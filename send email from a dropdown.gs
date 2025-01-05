function notif_ticket_raised(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Get the header row and find the column index for "method"
  var headerRow = 1; // Assuming headers are in the first row
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var methodColumnIndex = headers.indexOf("method") + 1; // Adding 1 because column indices are 1-based

  // Check if the edited sheet is "Workdesk" and if the edit is happening in the "method" column
  // Also, check if the value is "ticket_raised"
  if (sheet.getName() == "Workdesk" && range.getColumn() == methodColumnIndex && range.getValue() == "ticket_raised") {
    sendNotification(sheet, range.getRow());
  }
}

function sendNotification(sheet, row) {
  // Get the values from columns D, E, F, G, H, and I in the same row where "ticket_raised" is selected
  var colD = sheet.getRange(row, 4).getValue(); // Column D
  var colE = sheet.getRange(row, 5).getValue(); // Column E
  var colF = sheet.getRange(row, 6).getValue(); // Column F
  var colG = sheet.getRange(row, 7).getValue(); // Column G
  var colH = sheet.getRange(row, 8).getValue(); // Column H
  var colI = sheet.getRange(row, 9).getValue(); // Column I
  
  // Compose the email body
  var body = "Non-identifiable bank payment // Please help!\n\n";
  body += "Here are the details. @support-team -->Please help us match this payment\n";
  body += "Valutadatum:" + colD + "\n";
  body += "Empfängername/Auftraggeber:" + colE + "\n";
  body += "IBAN/Kontonummer:" + colF + "\n";
  body += "BIC/BLZ:" + colG + "\n";
  body += "Verwendungszweck:" + colH + "\n";
  body += "Betrag in EUR:" + colI + "\n";
  
  // Send the email
  var emailAddress = "xxx"; // Your email address
  var subject = "Non-identifiable bank payment // Please help";
  MailApp.sendEmail(emailAddress, subject, body);
}
