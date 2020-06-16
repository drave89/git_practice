function uploadBalance() {
  download.clearContents();
  var downloads = info.getRange(3, 2, 3).getValues(); 
  var sourceSheetA = SpreadsheetApp.openById(downloads[0])
  var sourceValA = sourceSheetA.getActiveSheet().getDataRange().getValues();
  download.getRange(1, 1, sourceValA.length, sourceValA[0].length).setValues(sourceValA);
  info.getRange(2, 4).setFormula("=unique(master!J:J)");
  info.getRange(2, 4).setNumberFormat('@');
  
  for(var i = 1; i < 3; i++) {
    var sourceSpreadsheet = SpreadsheetApp.openById(downloads[i]);
    var sourceVal = sourceSpreadsheet.getActiveSheet().getRange(8, 1, sourceSpreadsheet.getLastRow(), sourceSpreadsheet.getLastColumn()).getValues();
    var downloadLastRow = download.getLastRow(); 
    download.getRange(downloadLastRow, 1, sourceVal.length, sourceVal[0].length).setValues(sourceVal);
  }
}

function generateID () {
  var title = download.getRange(6, 1).setValue("Item #");
  var balance = download.getRange(6, 16).setValue("Balance");
  var targetRange = download.getRange(8, 1, download.getLastRow());
  var rangeLastRow = targetRange.getNumRows() - 7; 
  var itemArr = [];
  
  for (var i = 0; i < rangeLastRow - 7; i++) {
    itemArr.push([i]); 
  }
  var writeRange = download.getRange(8, 1, itemArr.length);
  writeRange.setValues(itemArr);
}

function concatenateFtid () {
  var range = download.getRange(8, 4, download.getLastRow());
  var referenceStart = 8; 
  var concat = range.setFormula("=concatenate(C8,G8)")  
 }

function uploadBalances () {
  uploadBalance(); 
  concatenateFtid();
  generateID(); 
}
