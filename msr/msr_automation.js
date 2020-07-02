// test


function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Monthly Status Reports')
  .addItem('Upload balances', 'downloadSheet')
  .addItem('Generate reports', 'createSheets')
  .addItem('Refresh exceptions', 'generateException')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Email')
              .addItem('Send all emails', 'sendEmails')
              .addItem('Send single email', 'singleEmail')
              .addItem('Send follow  up email', 'followUp')
             )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Form')
             .addItem('Update form', 'amendForm')
              .addItem('Retrieve form responses', 'getFormResponses')             
             )
  .addItem('Delete sheets', 'deleteSheets')
  .addItem('Roll forward', 'rollForward')
  .addToUi(); 
}

function setProtection() {
  //set protection on sheets to only have mandy and suzanne have edit access and remove all others, except for protected ranges
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var infoSheet = spreadsheet.getSheetByName("Info"); 
  var exceptionSheet = spreadsheet.getSheetByName("Exceptions");
  var mandy = "amclennan@atb.com"; 
  var suzanne = "slemieux-barrett@atb.com"; 
  var dominic = "dturcotte@atb.com"
  var financeOps = [mandy, suzanne, dominic]; 
  var exceptionProtect = exceptionSheet.protect();
  var protection = infoSheet.protect();
  exceptionProtect.removeEditors(exceptionProtect.getEditors())
  exceptionProtect.addEditors(financeOps); 
  protection.removeEditors(protection.getEditors()); 
  protection.addEditors(financeOps);
}

function infoSheet () {
  var attestRange = info.getRange(8, 1, info.getLastRow() - 7);
  var url_email_rec = info.getRange(8, 4, info.getLastRow() - 7, 3)
  attestRange.clearContent(); 
  url_email_rec.clearContent(); 
  var genDate = info.getRange(3, 4).setFormula("=today()"); 
  var month_end = info.getRange(4, 4).getValue();
  var dueDate = info.getRange(5, 4).setFormula("=workday(D3,10)").getValue(); 
  info.getRange(7, 1).setFormula("=unique(master!J:J)"); //grabbing all the unique attestation possiblities
  info.getRange(7, 1).setNumberFormat('@'); 
  var attestRange = info.getRange(7, 1, info.getLastRow() - 6)
  var sortRange = info.getRange(8, 1, info.getLastRow() - 7)
  attestRange.copyTo(attestRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  sortRange.sort({column: 1, ascending: true})
}

function downloadSheet () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheets = spreadsheet.getSheets(); 
  var info = spreadsheet.getSheetByName("Info");  
  var errors = sheets[2];
  var urls = info.getRange(3, 2, 3).getValues(); 
  var downloadArr = []; 

  Logger.log(urls); 
  for(var i = 0; i < urls.length; i++) {
    var values = SpreadsheetApp.openByUrl(urls[i]).getDataRange().getValues()
    values.splice(0, 7) // removing first 7 rows of each dataset 
    downloadArr.push(values)
  }
  
  var downloadArr = downloadArr.flat(); // getValues is a 2d array, but it is doing x 3, making it a 6d array. Flattening it makes it back to a 2d array. 
  
  var download = sheets[4];
  var key = 0; 
  var downCount = 1; 
  var downTransit = 2; 
  var downFtid = 3; 
  var downAttest = 4; 
  var onMaster = 5; 
  var downAccount = 6; 
  var downProduct = 7; 
  var downDate = 11; 

  var downBalance = 15; 
  var downRisk = 16; 
  var downBSorIS = 17; 
  var downExpectedBal = 18; 
  var downGl = 19; 
  var downPurpose = 20; 
  
  var master = sheets[5]
  var masterValues = master.getDataRange().getValues(); // grabbing all master values and indices
  var masterProduct = 0; 
  var masterTransit = 1;
  var masterAccount = 2; 
  var masterFtid = 3; 
  var masterGl = 4; 
  var masterDesc = 5; 
  var masterCurr = 6; 
  var masterExpBal = 7; 
  var masterAccountability = 9;
  var masterPurp = 11; 
  var masterRisk = 12; 
  var masterBSorIS = 13;
  
  var map = new Map(); // setting up a mapping table for each masterFtid value and their respective properties into an object
  for (var i = 0; i < masterValues.length; i++) {
    var eachRow = masterValues[i]; 
    map.set(eachRow[masterFtid], {
      "accountability": eachRow[masterAccountability], 
      "product": eachRow[masterProduct],
      "transit": eachRow[masterTransit], 
      "account": eachRow[masterAccount], 
      "gl": eachRow[masterGl],
      "desc": eachRow[masterDesc], 
      "expBal": eachRow[masterExpBal], 
      "purpose": eachRow[masterPurp], 
      "risk": eachRow[masterRisk],
      "bsOrIs": eachRow[masterBSorIS]
    })
  }
  var errorArr = []; 
  
  for (var i = 0; i < downloadArr.length; i++) {
    var row = downloadArr[i] //grab each row in the download array
    row[key] = i + 1 // create the unique key by simply adding 1 to the counter
    row[downFtid] = row[downTransit].toString() + row[downAccount].toString(); // concatenating the ftid
    var info = map.get(row[downFtid]) // getting the concatenated ftid from the above mapping function
    
    // if "getting" the concatted ftid returns a value, return true (meaning it was found on the master file), else false (meaning it was not on the master file) 
    if(info != undefined || info != null) {
      row[downAttest] = info.accountability  // from the ftid object above, grab the accountability
      row[downRisk] = info.risk // same here
      row[downBSorIS] = info.bsOrIs // same here
      row[downExpectedBal] = info.expBal
      row[downGl] = info.gl
      row[downPurpose] = info.purpose
      row[onMaster] = "true"
    } else {
      row[downAttest] = ""  // from the ftid object above, grab the accountability
      row[downRisk] = "" // same here
      row[downBSorIS] = "" // same here
      row[downExpectedBal] = ""
      row[downGl] = ""
      row[downPurpose] = ""
      row[onMaster] = "false"
      errorArr.push(row);
    }
  }

  
  var header = ["Key", "Country", "Transit", "FTID", "Attestation Accountability", "On Master?", "Account", "Description", "Product", "Account Holder", "BP Name", "Date", "Currency", "Debit Amount", "Cred. Amnt", "Bal", "Risk Rating", "IS or BS", "Exp. Bal.", "General Ledger", "Purpose"]
  downloadArr.unshift(header) //putting the header row as the first row
  Logger.log(downloadArr[downloadArr.length - 1]); 
  download.getRange(1, 1, downloadArr.length, downloadArr[0].length).setValues(downloadArr) // setting the values of the download array to the download sheet
  
  if(errorArr.length == 0) {
    SpreadsheetApp.getUi().alert("All FTID's from download files appear on Master file!")
  } else {
    SpreadsheetApp.getUi().alert("Found " + errorArr.length + " FTID's on download file that do not appear on Master file.") 
    errors.getRange(2, 1, errorArr.length, errorArr[0].length).setValues(errorArr);
  }
}

function createSheets () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var exceptionSheet = spreadsheet.getSheetByName("Exceptions"); 
  var downSheet = spreadsheet.getSheetByName("Download"); 
  var downRange = downSheet.getRange(2, 1, downSheet.getLastRow() - 1, downSheet.getLastColumn()).getValues(); // get all possible attestation values from the final download sheet 
  var unique = []; 
  var templateSheet = spreadsheet.getSheetByName("MSR_template")
  var dateVal = downRange[2][11]
  templateSheet.getRange(1, 1).setValue("Status Report on Internal Accounts for " + getDate(dateVal))
  var attestVal = downRange.forEach(function (row) {
    unique.push(row[4])
  })

  var onlyUnique = function (value, index, self) { // callback function to return only unique values
    return self.indexOf(value) === index;
  } 
  
  // defining indices for sheet that is being created
  var index = 0; 
  var country = 1; 
  var transit = 2; 
  var number = 3; 
  var description = 4; 
  var product = 5; 
  var date = 6; 
  var curr = 7; 
  var balance = 8; 
  var gl = 9; 
  var expBal = 10;
  var flag = 11; 
  var blank = ''
  var accountability = 15; 
  var purpose = 20; 
  
  //defining indices for existing download sheet
  var key = 0; 
  var downCount = 1; 
  var downTransit = 2; 
  var downFtid = 3; 
  var downAttest = 4; 
  var onMaster = 5; 
  var downAccount = 6; 
  var downDescription = 7; 
  var downProduct = 8; 
  var downDate = 11; 
  var downCurr = 12; 
  var downBalance = 15; 
  var downRisk = 16; 
  var downBSorIS = 17; 
  var downExpectedBal = 18; 
  var downGl = 19;  
  
  var uniqueAttest = unique.filter(onlyUnique).sort() //callback function to filter only unique values
  var twoDArr = []; 
  
  uniqueAttest.forEach(function (element) {
    twoDArr.push([element]);
  })
  
  Logger.log(twoDArr); 
                       
                    
                       
  var infoSheet = spreadsheet.getSheetByName("Info") 
  var infoStartRow = 8; 
  var infoRange = infoSheet.getRange(infoStartRow, 1, uniqueAttest.length).setValues(twoDArr); 
  var exceptionArr = []; 
  
  for(var i = 0; i < uniqueAttest.length; i++) {
    var sheet = spreadsheet.insertSheet(uniqueAttest[i], spreadsheet.getSheets().length, {template: templateSheet})
    var newSheet = sheet.getName(); // for each unique value (i.e. each attestation value) create one sheet and name it to the current loop counter value
    spreadsheet.getSheetByName(newSheet).getRange(2, 4).setValue(newSheet); 
    
    var newArr = downRange.filter(element => element[downAttest] == uniqueAttest[i])
    var formattedArr = []; 
    newArr.forEach(function (element) {
      var flag = determineFlag (element[downBalance], element[downExpectedBal], element[downRisk], element[downBSorIS])
      formattedArr.push([
        element[key], 
        element[downCount], 
        element[downTransit],
        element[downAccount],
        element[downDescription], 
        element[downProduct], 
        element[downDate],
        element[downCurr],
        element[downBalance],
        element[downGl],
        element[downExpectedBal], 
        flag, 
        "", 
        "", 
        "", 
        element[downAttest], 
        element[purpose]        
      ]);
      
      if(flag == "FLAG") {
        exceptionArr.push([
        element[key], 
        element[downCount], 
        element[downTransit],
        element[downAccount],
        element[downDescription], 
        element[downProduct], 
        element[downDate],
        element[downCurr],
        element[downBalance],
        element[downGl],
        element[downExpectedBal], 
        flag, 
        "", 
        "", 
        "", 
        element[downAttest], 
        element[purpose]        
      ]);      
      }
    })
    spreadsheet.getSheetByName(newSheet).getRange(4, 1, formattedArr.length, formattedArr[0].length).setValues(formattedArr)
    var protection = sheet.protect(); 
    var unprotected = sheet.getRange(4, 14, sheet.getLastRow(), 2); 
    var unprotect = protection.setUnprotectedRanges([unprotected]).addEditors(["AMcLennan@atb.com", "SLemieux-Barrett@atb.com"]).setDomainEdit(false); 
//    var formLink = "https://docs.google.com/forms/d/e/1FAIpQLSfoubasNB7mXwboMcatd_KqBytQxRi3qfY8NaB6xxbsgH2icg/viewform"
//    sheet.getRange(1, 5).setFormula("=HYPERLINK('https://docs.google.com/forms/d/e/1FAIpQLSfoubasNB7mXwboMcatd_KqBytQxRi3qfY8NaB6xxbsgH2icg/viewform', 'Link to Form 1017'")
  }
  
  if(exceptionArr.length < 1) {
    SpreadsheetApp.getUi().alert("No exceptions found!") 
  } else {
    SpreadsheetApp.getUi().alert(exceptionArr.length + " exceptions found!")
    exceptionArr.unshift(["Item #", "Bank Ctry", "Transit", "Account Number", "Account Description", "Product",	"Date", "Currency", "Balance", "General Ledger", "Expected Balance", "Flag", "Finance Comments", "AOE Comments", "Correction Date", "Attestation Accountability", "Purpose"])
    exceptionSheet.getRange(2, 1, exceptionArr.length, exceptionArr[0].length).setValues(exceptionArr); 
    exceptionSheet.activate(); 
  }
  getUrl(); 
  setProtection(); 
}

function determineFlag (balance, expectedBalance, riskRatingValue, categoryValue) {
  
  if (balance == 0) {
    return "CLEAR"; 
  } else if (balance > 0 && expectedBalance == "Zero" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG"; 
  } else if (balance < 0 && expectedBalance == "Zero" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG";
  } else if (balance > 0 && expectedBalance == "Zero-delayed - DR" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG";  
  } else if (balance > 0 && expectedBalance == "Non-zero - DR" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG";  
  } else if (balance < 0 && expectedBalance == "Zero-delayed - CR" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG";
  } else if (balance < 0 && expectedBalance == "Non-zero - CR" && riskRatingFlag(balance, riskRatingValue, categoryValue) == true) {
    return "FLAG";
  } else {
    //      Logger.log("ERROR at row " + [j] + " on sheet" + sheet.getName()); 
    return "CLEAR"
    //      Logger.log(" balance: " + balance + " riskCol: " + riskRatingValue + " bsIsCol: " + categoryValue + " riskrating eval: " + riskRatingFlag(balance, riskRatingValue, categoryValue))
    
  }
  //    row[flagCol] = computedFlag;   
}
 
function riskRatingFlag(balance, riskRating, bs_is) {
  
  if (balance == 0) {
    return false
  } else if ((riskRating == "High" && bs_is == "B/S") || (riskRating == "High" && bs_is == "I/S")) {
    return true
  } else if (riskRating == 'Medium' && bs_is == 'B/S' && Math.abs(balance) > 500) {
    return true
  } else if (riskRating == 'Low' && bs_is == 'B/S' && Math.abs(balance) > 1500) {
    return true 
  } else if (riskRating == 'Medium' && bs_is == 'I/S' && Math.abs(balance) > 250) {
    return true
  } else if (riskRating == 'Low' && bs_is == 'I/S' && Math.abs(balance) > 500) {
    return true
  } else {
    return false
  }
    
}

function deleteSheets () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var getSheets = spreadsheet.getSheets();
  
  for (var i = 6; i < getSheets.length; i++) {
    spreadsheet.deleteSheet(getSheets[i])
  }
} 

function generateException () {
  //function to grab all exceptions of each sheet, and populate them on "exceptions" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets(); //get array of all sheets
  var flagArr = []; //initialize empty array for flag values
  var errArr = []; //initialize empty array for error values
  var ui = SpreadsheetApp.getUi();


  for (var j = 6; j < sheets.length; j++) {
    var sourceSheet = spreadsheet.getSheetByName(sheets[j].getName());
    var range = sourceSheet.getRange(4, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).getValues();
    //    Logger.log(range[0][11]);     
    for (var i = 0; i < range.length; i++) {
      var row = range[i];
      var flag = 11; 
      var aoeCom = 13; 
      
      if(row[flag] == "FLAG") {
        flagArr.push(row)
      }
      if(row[flag] == "ERROR") {
        errArr.push(row);
      }
      if(row[aoeCom] !== "" && row[flag] == "CLEAR" ) { //if comments in aoe column are not blank AND are clear, push to row
        flagArr.push(row); 
      }
    }
  }
  
  if(flagArr.length < 1 || flagArr == undefined) {
    ui.alert('No flags found'); 
  } else {
    
    var writeFlag = spreadsheet.getSheetByName("Exceptions").getRange(4, 1, flagArr.length, sourceSheet.getLastColumn());
    
    writeFlag.setValues(flagArr);
     
    ui.alert(flagArr.length + ' flags found');
  }
  
  if(errArr.length < 1 || errArr == undefined) {
    ui.alert('All accounts are \'CLEAR\' or \'FLAG\''); 
  } else {
    
    var writeError = spreadsheet.getSheetByName("Exceptions").getRange(flagArr.length + 1, 1, errArr.length, sourceSheet.getLastColumn()); //getting range to write error values
    writeError.setValues(errArr); //writing error values
    ui.alert('Could not set \'CLEAR\' or \'FLAG\' to ' + errArr.length + ' accounts.'); 
  }
}

function clearExceptions () {
  //function to clear exceptions and leave headings
  var exceptionRange = exceptions.getRange(4, 1, exceptions.getLastRow() - 3, exceptions.getLastColumn());
  exceptionRange.clear();
}

function clearErrors () {
  var errorRange = spreadsheet.getSheetByName("Errors"); 
  errorRange.clearContents(); 
}


function getUrl() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Info");
  var range = sheet.getRange(8,1, sheet.getLastRow() - 7, sheet.getLastColumn());
  var values = range.getValues(); 
  var sheets = spreadsheet.getSheets(); 
  var valuesLength = values.length; 
  var ssU = spreadsheet.getUrl();
  var urlArr = [];
  
  
  //build url from "url" + grid id" 
  for (var i = 6; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    var gid = sheets[i].getDataRange().getGridId(); 
    var url = ssU + "#gid=" + gid;
    urlArr.push([url])
  }
  sheet.getRange(8, 4, urlArr.length, urlArr[0].length).setValues(urlArr); 
}

function sendEmails () {
  var formApp = FormApp.openById("1mkZz7M_92lP0EbIqIxBas6Qt1kr8f_ejj9qpvleDw5o");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Info");
  var range = sheet.getRange(8, 1, sheet.getLastRow() - 7, 5);
  var values = range.getValues();
  var email_range = sheet.getRange(8, 2, sheet.getLastRow() - 7, 1); 
  var email_val = email_range.getValues();
  //  Logger.log(email_val); // list of emails 
  //  Logger.log(range[1].length); //5
  var date = new Date(); 
  var ui = SpreadsheetApp.getUi();
  var startRange = 8; 
//  var genDate = info.getRange(3, 4).setFormula("=today()"); 
  var month_end = sheet.getRange(4, 4).getValue();
  var dueDate = sheet.getRange(5, 4).getValue(); 
  var formatYear = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "YYYY");
  var formatMonth = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMMM");
  var formatDay = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "d");
  var formUrl = formApp.getPublishedUrl(); 
  
  var response = ui.alert('Send email(s) to ' + email_val.length + ' recipient(s)?', ui.ButtonSet.OK_CANCEL);
  var emailEval = HtmlService.createTemplateFromFile('emailOut');   

  if(response == ui.Button.OK) {
    for (var i = 0; i < email_val.length; i++) {

      var index = i + 1; 
      var allRows = values[i]; 
      var attest = 0; 
      var toLine = 1;
      var cc = 2;  
      var url = 3; 
      var confirm = 4; 

      emailEval.month_end = month_end; 
      emailEval.formatYear = formatYear; 
      emailEval.formatMonth = formatMonth; 
      emailEval.formatDay = formatDay; 
      emailEval.formUrl = formUrl; 
      emailEval.attest = allRows[attest]; 
      emailEval.url = allRows[url]; 
      
      var evaluatedEmail = emailEval.evaluate().getContent(); 

      var message = {
        to: allRows[toLine], 
        cc: allRows[cc], 
        replyTo: "FinanceOps-Processing@atb.com",
        subject: "Internal Accounts Monthly Status Report for " + month_end + " " + formatYear,
        htmlBody: evaluatedEmail,
        name: "FinanceOps-Processing@atb.com" 
      }
      
      if(allRows[toLine] != "" || allRows[toLine] == undefined) {
        MailApp.sendEmail(message); 
        allRows[confirm] = 'true';
        sheet.getRange(i + startRange, 5).setValue('true')
        Logger.log('true ' + 'at index ' + i + 'is ' + allRows[confirm]);  
      } else {
        ui.alert('Error at index ' + index)
        allRows[confirm] = 'false'; 
        Logger.log('false ' + 'at index ' + i + 'is ' + allRows[confirm])
        sheet.getRange(i + startRange, 5).setValue('false')
      }
    }
    range.setValues(values);
  } else {
    ui.alert('Email cancelled'); 
  }

}


function getDate (date) {
//  var date = new Date(); 
  var formatYear = Utilities.formatDate(date, Session.getScriptTimeZone(), "YYYY")
  var formatMonth = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM")
  var formattedDate = formatMonth + " " + formatYear; 
  var msr_date = msr_template.getRange(1,1).setValue("Status Report on Internal Accounts for " + formattedDate)
}

function singleEmail () {
  var formApp = FormApp.openById("1mkZz7M_92lP0EbIqIxBas6Qt1kr8f_ejj9qpvleDw5o");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var info = spreadsheet.getSheetByName("Info") 
  var values = info.getRange(8, 1, info.getLastRow() - 7, 5).getValues();
  var attest = 0; 
  var email = 1; 
  var startRow = 8;
  var cc = 2; 
  var url = 3; 
  var sent = 4; 
  var selection = info.getActiveRange().getRowIndex() - startRow;  
  var selectedRow = values[selection];
  //  var genDate = info.getRange(3, 4).setFormula("=today()"); 
  var month_end = info.getRange(4, 4).getValue();
  var dueDate = info.getRange(5, 4).getValue();
  var formatYear = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "YYYY");
  var formatMonth = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMMM");
  var formatDay = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "d");
  var formUrl = formApp.getPublishedUrl(); 
  var emailEval = HtmlService.createTemplateFromFile('emailOut');

  
  emailEval.month_end = month_end; 
  emailEval.formatYear = formatYear; 
  emailEval.formatMonth = formatMonth; 
  emailEval.formatDay = formatDay; 
  emailEval.formUrl = formUrl; 
  emailEval.attest = selectedRow[attest]; 
  emailEval.url = selectedRow[url]; 
  var evaluatedEmail = emailEval.evaluate().getContent(); 
  
  var message = {
    to: selectedRow[email], 
    cc: selectedRow[cc],
    replyTo: "FinanceOps-Processing@atb.com",
    subject: "Internal Accounts Monthly Status Report for " + month_end + " " + formatYear,
    htmlBody: evaluatedEmail,
    name: "FinanceOps-Processing@atb.com" 
  } 
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send email to ' + selectedRow[email] + '?', ui.ButtonSet.OK_CANCEL); 
  
  if(response == ui.Button.OK) {
    MailApp.sendEmail(message); 
    ui.alert('Email sent to ' + selectedRow[email]); 
    var range = info.getRange(selection + startRow, sent + 1).setValue("TRUE")
    } else {
    ui.alert('Email cancelled')
    var range = info.getRange(selection + startRow, sent + 1).setValue("FALSE")
    }
}

function followUp () {
  var formApp = FormApp.openById("1mkZz7M_92lP0EbIqIxBas6Qt1kr8f_ejj9qpvleDw5o");
  var date = new Date(); 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Info") 
  var values = sheet.getRange(8, 1, sheet.getLastRow() - 7, 5).getValues();
  var attest = 0; 
  var email = 1; 
  var startRow = 8;
  var cc = 2; 
  var url = 3; 
  var sent = 4; 
  var selection = sheet.getActiveRange().getRowIndex() - startRow;  
  var selectedRow = values[selection];
//  var genDate = info.getRange(3, 4).setFormula("=today()"); 
  var month_end = sheet.getRange(4, 4).getValue();
  var dueDate = sheet.getRange(5, 4).getValue();
  var formatYear = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "YYYY");
  var formatMonth = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "MMMM");
  var formatDay = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "d");
  var formUrl = formApp.getPublishedUrl();   
  
  var emailEval = HtmlService.createTemplateFromFile('followUpEmail');   
  emailEval.month_end = month_end; 
  emailEval.formatYear = formatYear; 
  emailEval.formatMonth = formatMonth; 
  emailEval.formatDay = formatDay; 
  emailEval.formUrl = formUrl; 
  emailEval.attest = selectedRow[attest]; 
  emailEval.url = selectedRow[url]; 
  
  var evaluatedEmail = emailEval.evaluate().getContent();
  
  var message = {
    to: selectedRow[email], 
    cc: selectedRow[cc],
    replyTo: "FinanceOps-Processing@atb.com",
    subject: "Follow up: Internal Accounts Monthly Status Report for " + month_end + " " + formatYear,
    htmlBody: evaluatedEmail, 
    name: "FinanceOps-Processing@atb.com" 
  } 
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Follow up with ' + selectedRow[email] + '?', ui.ButtonSet.OK_CANCEL); 
  
  if(response == ui.Button.OK) {
    MailApp.sendEmail(message); 
    ui.alert('Email sent to ' + selectedRow[email]); 
    var range = sheet.getRange(selection + startRow, sent + 1).setValue("TRUE")
    } else {
    ui.alert('Email cancelled')
    var range = sheet.getRange(selection + startRow, sent + 1).setValue("FALSE")
    }
}

function rollForward () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets(); 
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Roll forward to: (Enter month)', ui.ButtonSet.OK_CANCEL); //prompt user to enter what they want the next month spreadsheet to be called
  var response = prompt.getResponseText(); //grab user response
  
  if(prompt.getSelectedButton() == ui.Button.OK) {
    var createdSheet = spreadsheet.copy('MSR '+ response).getId(); 
    var newSheet = SpreadsheetApp.openById(createdSheet); 
    var new_sheets = newSheet.getSheets();
    var download_sheet = 4; 
    
    
    for(var i = 6; i < new_sheets.length; i++) {
      newSheet.deleteSheet(new_sheets[i]);  
      new_sheets[download_sheet].clear(); 
    }
  } else {
    ui.alert('Operation cancelled');
  }
}

function amendForm () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Info");
  var startRow = 8; 
  var attestValues = sheet.getRange(8, 1, sheet.getLastRow() - startRow + 1).getValues().flat(); 
  
  var form = FormApp.openById("1mkZz7M_92lP0EbIqIxBas6Qt1kr8f_ejj9qpvleDw5o");
  var questions = form.getItems();
  var aoe_question = questions[1]; 
  var choices = aoe_question.asCheckboxItem().setChoiceValues(attestValues);
}

function getFormResponses () {
  var formSpreadsheet = SpreadsheetApp.openById("1Q2b9WBAoJd4A01mDp8Nh1OqHJeKbxJ3EIsl23X0PY4w");
//  var formSheet = formSpreadsheet.getActiveSheet();
  var formSheet = formSpreadsheet.getSheets()[0]; 
  var info = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Info")
  var formDataRange = formSheet.getRange(2, 1, formSheet.getLastRow()-1, formSheet.getLastColumn());  
  var formDataVal = formDataRange.getValues();
  var msr_range = info.getRange(8, 1, info.getLastRow() - 7, info.getLastColumn());  
  var msr_val = msr_range.getValues();
  
  //converting each attestation accountability to a string since sheets converts it back to a number if the accountability is just a transit
  for(var s = 0; s < msr_val.length; s++) {
    msr_val[s][0] = msr_val[s][0].toString()
  }
  
  //splitting the responses into separate indices in the response array since the form returns it all as a single string even though multiple options are selected
  var formAoE = 2; 
  var attestReceived = 5; 
  var newArr = []; 
  for(var i = 0; i < formDataVal.length; i++) {
    var rows = formDataVal[i]; 
    var aoes = rows[2]; 
    newArr.push(aoes.split(', '));
  }
  
  //above code results in the array looking like: [1, 2, [3, 4, 5]]. Flattening turns it into [1, 2, 3, 4, 5] for easier looping
  var flatArr = newArr.flat(); 

  //loop through each element in the flattened array and match with values on msr attestation accountability values. setting the index value 
  for(var i = 0; i < msr_val.length; i++) {
    if(flatArr.includes(msr_val[i][0])) {
      msr_val[i][5] = "TRUE" 
    } else {
      msr_val[i][5] = "FALSE"
    }
  }
  
  //looping through (again) and setting the value of attestReceived (5) to the actual cell
  for(var b = 0; b < msr_val.length; b++) {
    var attest_range = info.getRange(b + 8, 6);
    attest_range.setValue(msr_val[b][5])
  }
}; 


function getDate (date) {
//  var date = new Date (); 
  var month = date.getMonth() + 1; 
  var year = date.getFullYear(); 
  var formattedDate = Utilities.formatDate(new Date (year, month - 1), "GMT", "MMMMM YYYY"); 
  
  return formattedDate; 
}