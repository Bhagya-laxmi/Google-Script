//function to retrive data from the spreadsheet containing email address and message

function myFunction() {
  Browser.msgBox("hello")
  var check= 'check';
  if( check == 'check!')
  {
  Browser.msgBox(check)}
  else {Browser.msgBox('Mistake')}
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

  //var sh0 = sheet.getSheetByName("Sheet1");
  var cols = sheet.getLastColumn();
  Browser.msgBox('last column: ' + cols)
  
  var startRow = 2; // First row of data to process
  var numRows = 1; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 4);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var iLastRow = SpreadsheetApp.getActiveSheet().getMaxRows();
  
  Browser.msgBox(iLastRow)
  for (var i in data) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1] +row[2]; // Second column
   
    var subject = 'Sending emails from a Spreadsheet';
    MailApp.sendEmail(emailAddress, subject, message);
    
    
  }
  var data4 = data[4];
  if( data4 == null)
  {
  Browser.msgBox("No data");
  }
}
