// Menu
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('✅ Check', 'checkEmailAddresses')
  .addItem('Information', 'showHowTo')
  .addToUi();
}

// Check Email Addresses function
function checkEmailAddresses() {
  var row = 1;
  var col = 1;
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(row,col);
  var lastRowContent = sheet.getLastRow();
  var dataRange = sheet.getRange(row, col, lastRowContent, 2);
  var data = dataRange.getValues(); // Fetch values for each row in the Range.
  SpreadsheetApp.getActiveSpreadsheet().toast("👌 Validation started", "Status", -1);
  for (var i = 0; i < lastRowContent; i++) {
    var rowData = data[i];
    var email = rowData[0];  // First column
    var status = rowData[1]; // Second column
    if (email != "" && status == "") {
      var response = validateEmailAddress(email);
      sheet.getRange(row + i, 2).setValue(response);
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Validation is done", "Status", -1);
}

// Validate Email Address function
function validateEmailAddress(email) {
  var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  var result = re.test(email);
  Logger.log(result);
  return result;
}