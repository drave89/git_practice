var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Approval').addItem('Send E-mail to Approver', 'sendEmails2').addSeparator().addToUi(); //.addSubMenu('The rest').addItem('Second Menu','secondItem').addToUi();
  //ss.addMenu("Accrual JE: Approval Workflow", 'Approval');
  //  spreadsheet.toast("Input business unit approver information -> Click 'Approval' -> Authorize Script -> Email sent.", "Get Started", -1);
}

// This constant is written in column C for rows for which an email
// has been sent successfully.
//var EMAIL_SENT = "EMAIL_SENT";

function sendEmails2(){ 
  var d = new Date();
  var curr_time = d.toLocaleTimeString();
  var curr_date = d.getDate();
  var curr_month = d.getMonth() + 1;
  var curr_year = d.getFullYear();
  var theDate = curr_month + "/" + curr_date + "/" + curr_year + " " + curr_time;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet ();
  var sheet = ss.getSheetByName("JE ");
  var sheet1 = ss.getSheetByName("Checks");
  var source = sheet.getRange("C23").getValues();
  var source2 = sheet.getRange("G16").getValues();
  if (source != 0) { 
    Browser.msgBox('THE DEBITS AND CREDITS DO NOT BALANCE. PLEASE CHECK YOUR ENTRY AND TRY AGAIN.', Browser.Buttons.OK);
  } else if (source2 == "") {
    Browser.msgBox('Please ensure your business unit approver e-mail is filled in cell G16.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B2").getValues() == "INVALID") {  
    Browser.msgBox('GL Accounts 77117000, 77117100, and 77118000 in your entry cannot be posted with Z400 and Z700 Internal Orders. Please check your entry and try again.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B6").getValues() == "INVALID") {  
    Browser.msgBox('GL Accounts 77117000, 77117100, and 77118000 in your entry cannot be posted with any Cost Centers. Please check your entry and try again.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B10").getValues() == "INVALID") {  
    Browser.msgBox('GL Accounts 7XXXXXXX in your entry cannot be posted with any Profit Centers, as the CC or IO identfied is already assigned with a Profit Center. Please check your entry and try again.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B14").getValues() == "INVALID") {  
    Browser.msgBox('GL Accounts 7XXXXXXX in your entry can only be posted with either a Cost Center OR an Internal Order, not both. Please check your entry and try again.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B18").getValues() == "INVALID") {  
    Browser.msgBox('If you post in a Profit Center, please do not include any information in the Cost Center or Internal Order fields.', Browser.Buttons.OK);
  } else if (sheet1.getRange("B22").getValues() == "INVALID") {  
    Browser.msgBox('For accounts not starting with 7, please post in a Profit Center', Browser.Buttons.OK);
  } else{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = DriveApp.getFileById(ss.getId());
    sheet.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW)
    
    var EMAIL_SENT = "EMAIL_SENT";
    //var sheet = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Template').activate();
    
    var sheet = SpreadsheetApp.getActiveSheet();
    
    var startRow = 2;  // First row of data to process
    var numRows = 1;   // Number of rows to process
    // Fetch the range of cells A2:B3
    var dataRange = sheet.getRange(startRow, 1, numRows, 4)
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
      var row = data[i];
      var emailAddress = row[0];  // First column
      var replyToAddress = "UploadsProcessing@atb.com"
      var message = row[1]+"\n\n"+ss.getUrl();        // Second column
      var name = SpreadsheetApp.getActiveSpreadsheet().getName();
      var subject = row[2]+ " " +name +" "+ theDate ;
      var ccAddress = row[3];     // Fourth column
      //if (ccAddress != EMAIL_SENT) {  // Prevents sending duplicates
      //var subject = "Approval Request: Create Accrual Entries";
      MailApp.sendEmail(emailAddress, subject, message, {
        cc: ccAddress,
        replyTo: replyToAddress
      });
      //sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      var sheet = SpreadsheetApp.getActiveSpreadsheet();
      SpreadsheetApp.setActiveSheet(sheet.getSheets()[0]);
    }
  }
}
//}

//function onEdit() {
// This script prevents cells from being updated. When a user edits a cell on the master sheet,
// it is checked against the same cell on a helper sheet. If the value on the helper sheet is
// empty, the new value is stored on both sheets.
// If the value on the helper sheet is not empty, it is copied to the cell on the master sheet,
// effectively undoing the change.
// The exception is that the first few rows and the first few columns can be left free to edit by
// changing the firstDataRow and firstDataColumn variables below to greater than 1.
// To create the helper sheet, go to the master sheet and click the arrow in the sheet's tab at
// the tab bar at the bottom of the browser window and choose Duplicate, then rename the new sheet
// to Helper.
// To change a value that was entered previously, empty the corresponding cell on the helper sheet,
// then edit the cell on the master sheet.
// You can hide the helper sheet by clicking the arrow in the sheet's tab at the tab bar at the
// bottom of the browser window and choosing Hide Sheet from the pop-up menu, and when necessary,
// unhide it by choosing View > Hidden sheets > Helper.
// See https://productforums.google.com/d/topic/docs/gnrD6_XtZT0/discussion

// modify these variables per your requirements
// var masterSheetName = "JE " // sheet where the cells are protected from updates
//  var helperSheetName = "Helper" // sheet where the values are copied for later checking

// range where edits are "write once": M20:V30, i.e., rows 20-30 and columns 13-22
// var firstDataRow = 20; // only take into account edits on or below this row
//  var lastDataRow = 100; // only take into account edits on or above this row
//  var firstDataColumn = 1; // only take into account edits on or to the right of this column
//var lastDataColumn = 50; // only take into account edits on or to the left of this column

//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var masterSheet = ss.getActiveSheet();
//  if (masterSheet.getName() != masterSheetName) return;

// var masterCell = masterSheet.getActiveCell();
//  if (masterCell.getRow() < firstDataRow || masterCell.getColumn() < firstDataColumn) return;

// var helperSheet = ss.getSheetByName(helperSheetName);
//  var helperCell = helperSheet.getRange(masterCell.getA1Notation());
//  var newValue = masterCell.getValue();
//  var oldValue = helperCell.getValue();

//  if (oldValue == "") {
//    helperCell.setValue(newValue);
//  } else {
//    masterCell.setValue(oldValue);
//  }
//}