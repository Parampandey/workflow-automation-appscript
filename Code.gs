function sendEmailNotification(e) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  var name = sheet.getRange(lastRow, 1).getValue();
  var email = sheet.getRange(lastRow, 2).getValue();
  var department = sheet.getRange(lastRow, 3).getValue();
  var request = sheet.getRange(lastRow, 4).getValue();

  var subject = "New Workflow Entry Submitted";

  var message =
    "New Request Details:\n\n" +
    "Name: " + name + "\n" +
    "Email: " + email + "\n" +
    "Department: " + department + "\n" +
    "Request: " + request;

  MailApp.sendEmail("parampandey09773@gmail.com", subject, message);
}
function sendDailyReport() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var total = data.length - 1;

  var subject = "Daily Workflow Summary";
  var message = "Total entries today: " + total;

  MailApp.sendEmail("akb312777@gmail.com", subject, message);
}
function setupValidation() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Approved", "Rejected"])
    .build();

  sheet.getRange("E2:E1000").setDataValidation(rule);
}


