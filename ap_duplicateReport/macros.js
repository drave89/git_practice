function sortRange() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 2, ascending: true}, {column: 6, ascending: true}]);
};

//function highlightRow() {
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('344:344').activate();
//  spreadsheet.getActiveRangeList().setBackground('#ffff00');
//};