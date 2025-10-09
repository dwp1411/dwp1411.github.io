function Email_Update() {
  var sheetId = '1OnD9GXgPzAv64WvQoqKB-yOjwlL63v9NdNoLXvb273A';

  // Open the sheet
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Master');

  // Set a specific recipient for testing
  var recipients = 'dpool@crateandbarrel.com';

  // Get the range of cells you want to include in the PDF
  var range = sheet.getRange('B1:O30');

  // Export the range as a PDF in LANDSCAPE
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?exportFormat=pdf&format=pdf' +
    '&size=letter&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false&gid=' + range.getSheet().getSheetId() + '&range=' + range.getA1Notation();

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  var pdfBlob = response.getBlob().setName('191 Shift Update.pdf');

  // Get formatted date
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');

  // Send the email
  var subject = '191 Shift Update ' + formattedDate;
  var body = 'Please find below End of Shift Update for 191.';

  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: body,
    attachments: [pdfBlob]
  });
}