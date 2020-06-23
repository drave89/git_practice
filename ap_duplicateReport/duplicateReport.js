//below are global variables that are used throughout many functions in this script
//if any sheet names change, ensure that they are updated in these variables


const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
const vendorSetupSheet = spreadsheet.getSheetByName("Vendor Setup List"); 
const vendorNrtList = spreadsheet.getSheetByName("NRT Vendors"); 
const vendorExceptions = spreadsheet.getSheetByName("Vendor Payment Exceptions")
const EXTRACT_SHEET = "Copy of Duplicate Report April 7, 2020"
const exceptionSheet = spreadsheet.getSheetByName("Exceptions");
const possibleDuplicates = spreadsheet.getSheetByName("possibleDuplicates"); 
const currentRate = 1.32; //defining rate to check against  
const start_row = 11; //defining the start row of the extract sheet
const lowerThreshold = -50; //check anything over and under these two values
const upperThreshold = 50;
const highAmountThreshold = 25000; 

var duplicateArray = []; 
var headerArray = ["DocumentDate", "PostingDate", "ClearingDate", "netDueDate", "Reference", "AmountDocCurr", "DocCurr", "AmountLocCurr", "LocCurr", "Text", "DocumentNum", "ClearingDoc", "PayType", "Account", "ProfitCenter", "Type", "VendorName", "VendorCurr", "VendorPm", "VendorNRT", "VendorException", "Vendor Instructions", "DerivedRate"]
var exceptionHeader = ["DocumentDate", "PostingDate", "ClearingDate", "netDueDate", "Reference", "AmountDocCurr", "DocCurr", "AmountLocCurr", "LocCurr", "Text", "DocumentNum", "ClearingDoc", "PayType", "Account", "ProfitCenter", "Type", "VendorName", "VendorCurr", "VendorPm", "VendorNRT", "VendorException", "Vendor Instructions", "DerivedRate", "Exception"]; 


function onOpen() {
  var ui = SpreadsheetApp.getUi()
  .createMenu("Duplicate analysis")
  .addItem("Parse extract", "parseExtract").addItem("Identify duplicates", "identifyDuplicates")
  .addToUi()
  
  }

//setting up object for "Vendor Setup List", using that table as the baseline object
function vendorSetup () {
  var values = vendorSetupSheet.getDataRange().getValues(); 
  var vendorObj = {}; 
  var vendorNumIndex = 0; 
  var vendorNameIndex = 1; 
  var vendorCurrIndex = 2; 
  var vendorPmIndex = 3; 
  
  //for each row in the "Vendor Setup List" table, create an object to establish our base lookup destinations
  for(var i = 0; i < values.length; i++) {
    var row = values[i]; 
    var vendorNum = row[vendorNumIndex].toString(); //converting to string since sheets keeps thinking this is a number
    var vendorName = row[vendorNameIndex]; 
    var vendorCurr = row[vendorCurrIndex]; 
    var vendorPm = row[vendorPmIndex]; 
    var isNrt = false; //defaulting nrt value to false (we'll change this later if vendor is indeed nrt)
    var confirmation = false; //defaulting confirmation to false (we'll change this later if vendor has confirmation values
    var instructions = ""; 
    
    //if the vendorObj[vendorNum] does not contain the vendor number, create it and default the values as below
    //which in this function are the values found in the vendor setup table
    if(vendorObj[vendorNum] == undefined) { 
      vendorObj[vendorNum] = {
        "vendorName": vendorName, 
        "vendorCurr": vendorCurr, 
        "vendorPm": vendorPm, 
        "isNrt": isNrt,
        "confirmation": confirmation, 
        "instructions": instructions
      }
    }
  }
  //we use this as a callback function later on to retrieve the values from this object to do the lookups
  return vendorObj; 
}

//using the vendorObj from the "vendorSetup" function to populate the nrtValues
//if the vendor is found on the "NRT Vendors" table, change the isNrt variable to true
function nrtVendors () {
  var values = vendorNrtList.getDataRange().getValues(); 
  var vendorObj = vendorSetup(); 
  var vendorNumIndex = 0; 
  var vendorNameIndex = 1; 
  var vendorNrtIndex = 2; 
  var isNrt = true;
  
  for(var i = 1; i < values.length; i++) {
    var rows = values[i]; 
    var id = rows[vendorNumIndex].toString();  
    
    //if in the event that the vendor number isn't found on the original object, create it and initialize the name to the name found
    //and initialize the nrtValue as true (since it was found on the NRT table. If it was found, update the object with the "isNrt" property as true
    if(vendorObj[id] == null) {
      vendorObj[id] = {
        "vendorName": rows[vendorNameIndex], 
        "isNrt": isNrt      
      } 
    } else {
      vendorObj[id]["isNrt"] = isNrt
    }
  }
  return vendorObj
}

//similar to the other two, we are looping through each table set out in each sheet and updating the object from the vendorSetup function
//with values found on the table
function vendorException() {
  var values = vendorExceptions.getDataRange().getValues(); 
  var vendorObj = nrtVendors();  
  var vendorNumIndex = 0; 
  var vendorName = 1; 
  var vendorException = 2; 
  var vendorInstructions = 3; 
  
  for(var i = 0; i < values.length; i++) {
    var rows = values[i]; 
    var id = rows[vendorNumIndex].toString(); 
    
    //if id (vendor number) doesn't exist in the object, create it and init the values as what's found on the table
    if(vendorObj[id] == null) {
      vendorObj[id] = {
        "vendorName": rows[vendorNumIndex],
        "exception": rows[vendorException],
        "instructions": rows[vendorInstructions]
      }
    } else {
      
      //if it is found, update the properties of the id with what's found on the table
      vendorObj[id]["exception"] = rows[vendorException]; 
      vendorObj[id]["instructions"] = rows[vendorInstructions];
    }         
  }
  return vendorObj; 
}

function parseExtract (sheet) {
  var sheet = spreadsheet.getSheetByName(EXTRACT_SHEET); 
  var range = sheet.getDataRange(); 
  var extractValues = range.getValues(); 
  var vendorObj = vendorException(); //grabbing the values from the vendor object we created earlier
  
  var writeArr = []; 
  
  //defining indices for extracted values
  var docDate = 2; 
  var postDate = 3;
  var clearing = 5; 
  var netDue = 6; 
  var reference = 7; 
  var amountDocCurr = 8; 
  var docCurr = 9; 
  var amountLocCurr = 10; 
  var locCurr = 11; 
  var text = 12; 
  var docNo = 13; 
  var clearingDoc = 14; 
  var payType = 15;
  var vendorId = 16; 
  var pc = 17; 
  var type = 18; 
  
  //for each row on the extracted sheet, grab the id of each and pull the associated properties from the object created in the previous steps
  for(var i = start_row; i < extractValues.length; i++) {
    var row = extractValues[i];
    var id = row[vendorId].toString(); 
    
    
    //as long as the vendorId is not undefined in the object, proceed
    if(vendorObj[id] != undefined) {
      
      //performing the rate derivation, which is the document currency / local currency
      //we will preform the check against the current rate later
      var rate = row[amountLocCurr] / row[amountDocCurr]; 
      
      //for each row, create an array (read: row) with the associated lookups and get it ready to push to 
      //a write array which will be "pasted" on a new sheet
      var formattedRow = [ 
        row[docDate], 
        row[postDate], 
        row[clearing], 
        row[netDue], 
        row[reference], 
        row[amountDocCurr], 
        row[docCurr],
        row[amountLocCurr], 
        row[locCurr], 
        row[text], 
        row[docNo].toString(), 
        row[clearingDoc], 
        row[payType],
        id,
        row[pc],
        row[type],
        vendorObj[id]["vendorName"], 
        vendorObj[id]["vendorCurr"],
        vendorObj[id]["vendorPm"], 
        vendorObj[id]["isNrt"], 
        vendorObj[id]["exception"],
        vendorObj[id]["instructions"], 
        rate
      ];
      
      writeArr.push(formattedRow); 
      
      //if the id is not found in the object, it means that the vendor does not appear on any of the tables defined in each of the sheets
    } else if (vendorObj[id] == undefined) {
      var exceptionRow = [ 
        row[docDate], 
        row[postDate], 
        row[clearing], 
        row[netDue], 
        row[reference], 
        row[amountDocCurr], 
        row[docCurr],
        row[amountLocCurr], 
        row[locCurr], 
        row[text], 
        row[docNo].toString(), 
        row[clearingDoc], 
        row[payType],
        id,
        row[pc],
        row[type],
        "vendorName not found", 
        "vendorCurr not found", 
        "vendorPm not found", 
        "vendorNrt not found", 
        "vendorException not found", 
        "vendorInstructions not found", 
        rate
      ]
      //push the row to the overall write array
      //push the exception row to the exception array
      writeArr.push(exceptionRow); 
      //      exceptionArr.push(exceptionRow); 
    }
  }
  //use one of the helper functions below to write the data to a "raw data" sheet
  
  writeData(writeArr); 
  
  var exceptionArray = analyse(writeArr); 
  
  
  //if there is more than 0 (non-inclusive) in the exception array, paste it to the exception sheet
  if(exceptionArray.length > 0) {
    var prompt = exceptionArray.length + " exceptions found!"
    SpreadsheetApp.getUi().alert(prompt); 
    exceptionArray.unshift(exceptionHeader)
    exceptionSheet.getRange(1, 1, exceptionArray.length, exceptionArray[0].length).setValues(exceptionArray); 
  } else {
    var prompt = "No exceptions found!" 
    SpreadsheetApp.getUi().alert(prompt);
  }
  
}

//checking if the sheet already exists
function sheetExist (sheetName) {
  var sheets = spreadsheet.getSheets(); 
  
  for(var i = 0; i < sheets.length; i++) {
    var currSheetName = sheets[i].getName(); 
    if(currSheetName == sheetName) {
      
      return true
      Logger.log("Sheet already exists"); 
    } else {
      return false; 
    }
  }
}

function writeData (dataArray) {
  var RAW_DATA = "raw Data";
  var writeSheet = ''; 
  
  writeSheet = spreadsheet.getSheetByName(RAW_DATA)
  dataArray.unshift(headerArray); 
  writeSheet.getRange(1,1, dataArray.length, dataArray[0].length).setValues(dataArray); 
} 



function analyse (array) {
  //  var sheet = spreadsheet.getSheetByName("raw Data"); 
  //  var values = sheet.getDataRange().getValues(); 
  
  var docDate = 0; 
  var postDate = 1; 
  var clearing = 2; 
  var netDue = 3; 
  var reference = 4; 
  var amountDocCurr = 5; 
  var docCurr = 6; 
  var amountLocCurr = 7; 
  var locCurr = 8; 
  var text = 9; 
  var docNumber = 10; 
  var clearingDoc = 11; 
  var payType = 12; 
  var vendorId = 13; 
  var pc = 14; 
  var type = 15; 
  var vendorName = 16; 
  var vendorCurr = 17; 
  var vendorPm = 18; 
  var isNrt = 19; 
  var exception = 20; 
  var instruction = 21; 
  
  var exceptionArray = []; 
  
  //loop through each row in the raw data (pasted from the vendorsetup functions above) and find the exceptions
  for(var i = 0; i < array.length; i++) {
    var row = array[i]; 
    
    
    //use a few helper functions to determine the cause of the exception
    if(row[vendorName] == "vendorName not found") {
      
      row.push("Vendor information not found"); 
      exceptionArray.push(row)
      
    } else if (isWithin(row[amountDocCurr]) && !(currMatch(row[docCurr], row[vendorCurr]))) {
      
      row.push("Document currency and vendor currency do not match")
      exceptionArray.push(row);
      
    } else if (isWithin(row[amountDocCurr]) && (!rateMatch(row[amountDocCurr], row[docCurr], row[amountLocCurr], row[locCurr]))) {
      
      row.push("Rate does not match current rate of " + currentRate); 
      exceptionArray.push(row)    
      
    } else if (specialInstructions(row[exception], row[clearing])) {
      
      row.push("vendor has special instructions"); 
      exceptionArray.push(row); 
      
    } 
  }   
  
  return exceptionArray; 
}

//checks if amount is within the threshold amounts in the global variables 
function isWithin (amount) {
  
  if(amount >= upperThreshold || amount <= lowerThreshold) {
    return true
  } else {
    return false
  }
}                  

//checks if the document currency matches the vendor currency
//note, if the vendorCurr is "CAD/USD", flag it either way, do not do any checking
function currMatch (docCurr, vendorCurr) {
  if(docCurr == vendorCurr) {
    return true
  } else {
    return false
  }
} 

//checks to make sure the rate derived from the document amount and the local amount is a reasonable rate. 
//defined as a global variable 
function rateMatch (docAmount, docCurr, locAmount, locCurr) {
  
  var rateCheck = '';
  
  if(docCurr == locCurr) {
    rateCheck = 1; 
  } else if (docCurr != locCurr) {
    rateCheck = currentRate;  
  }
  
  var deriveRate = locAmount / docAmount;
  var returnValue = '';
  
  if (deriveRate == rateCheck) {
    returnValue = true;
  } else {
    returnValue = false;
  }
  return returnValue
}

//requirement of flagging if an exception on the vendor exists and clearing date is null
function specialInstructions (exception, clearingDate) {
  
  if (exception == 'Y' && clearingDate == '') {
    return true
  } else {
    return false
  }
}

//this function will group rows by date as it's looping, and finding duplicates within that date specifically and highlighting them
function identifyDuplicates () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName("raw Data"); 
  var lastRow = sheet.getLastRow(); 
  var lastCol = sheet.getLastColumn(); 
  var range = sheet.getRange(2, 1, lastRow, lastCol);
  
  
  var docDateIndex = 0; 
  var referenceIndex = 4; 
  var docAmountIndex = 5; 
  var docCurrIndex = 6; 
  var textIndex = 9; 
  var vendorIndex = 13; 
  var docNumber = 14; 
  
  //defining the sorting method, needs to be smallest to largest
  var sortMethod = [
    {column: docDateIndex + 1, ascending: true}, 
    {column: docAmountIndex + 1, ascending: true}
  ];
  
  range.sort(sortMethod); //sort the sheet (easier to do within sheets than code)
  
  var values = range.getValues(); //getting the values anew after sorting the sheet
  var bucket = []; //init the bucket we're currently looking at
  var startRow = 2; //excluding header row
  var dateForCurrentBucket = values[0][docDateIndex].toLocaleString(); //intializing the current date as the date found in the first row, changing to locale string cause sheets does weird stuff with dates
  var possibleDuplicateArr = []; 
  var highThresholdDuplicate = []; 
  
  for(var i = 0; i < values.length; i++) {
    var rowAmount = values[i][docAmountIndex]; 
    var row = values[i]; 
    var previousRow = values[i - 1]; 
    //if the amount is over the upperTreshold (defined in global), do the below. ignore for under 
    //added 06-17-2020 - to ensure that the bp's are not one of these two, as they are mastercard bp's
    if(row[docNumber] != "1148974" || row[docNumber != "1148973"]) {
      if(Math.abs(rowAmount) > upperThreshold) {
        if(Math.abs(rowAmount) >= highAmountThreshold) {
          highThresholdDuplicate.push(row);
        } 
        //add new properties here if we wanted to compare more
        var item = {
          rowId: i + startRow, //creating an id for each row, which in this case is just the row number. since it's 0 indexed, we need to do i + start row to get the actual row number
          date: values[i][docDateIndex].toLocaleString(), 
          amount: values[i][docAmountIndex]
        }; 
        
        //if the date on our newly created object is equal to the date initialized in var dateForCurrentBucket, push it into a bucket that we will check the amounts on later
        if(item.date == dateForCurrentBucket) {
          bucket.push(item); 
        } else {
          
          //if the length of the bucket is > 1 (i.e. two same dates were found), highlight the row using a helper function (which checks if any two amounts are the same)
          if(bucket.length > 1) {
            highlightDuplicates(bucket) 
          };
          
          //if nothing fit into the bucket, that means we're onto a new date (since there were no consecutive dates)
          dateForCurrentBucket = item.date; //reset the bucket to the current date loop counter
          bucket = [item]; //bucket is now equal to first item in the new date
        }
      }
    }
  } 
  highThreshold(highThresholdDuplicate); 
}

function highThreshold (array) {
  var over25kSheet = spreadsheet.getSheetByName("Over25k");
  var referenceIndex = 4; 
  var amountIndex = 5; 
  var vendorIndex = 13
  
  array.sort(function (a, b) {
    return a[amountIndex] - b[amountIndex]
  })             
  
  array.unshift(headerArray); 
  over25kSheet.getRange(1, 1, array.length, array[0].length).setValues(array); 
  
  var rangeToHighlight = []; 
  
  var amountToCheck = array[0][amountIndex];
  var vendorToCheck = array[0][vendorIndex]; 
  var referenceToCheck = array[0][referenceIndex]; 
  
  for(var i = 1; i < array.length; i++) {
    var row = array[i]; 
    
    
    var item = {
      rowId: i + 1,
      reference: row[referenceIndex], 
      amount: row[amountIndex],
      vendor: row[vendorIndex]     
    }
    
    if((item.amount == amountToCheck && item.vendor == vendorToCheck) || (item.amount == amountToCheck && item.reference == referenceToCheck)) {
      rangeToHighlight.push(over25kSheet.getRange(item.rowId - 1, 1, 1, array[0].length).getA1Notation()); 
      rangeToHighlight.push(over25kSheet.getRange(item.rowId, 1, 1, array[0].length).getA1Notation()); 
    }
    
    amountToCheck = array[i][amountIndex];
    vendorToCheck = array[i][vendorIndex]; 
    referenceToCheck = array[i][referenceIndex]; 
    
  }
  over25kSheet.getRangeList(rangeToHighlight).setBackground("#ffff00"); 
}


function testHighlights () {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName("Sheet16"); 
  var ranges = sheet.getRangeList(["A1:W1", "A15:W15"]).setBackground('#ffff00')
  
  }

var rawDataSheet = spreadsheet.getSheetByName("raw Data"); 
function highlightDuplicates (bucket) { 
  //within each bucket, sort the amounts from largest to smallest
  bucket.sort(function (a, b) {
    if(a.amount < b.amount) return -1; 
    if(a.amount > b.amount) return 1; 
  }); 
  
  var amountForCurrentItem = bucket[0].amount; //same concept as the date check above, set the initial amount as the first item in the bucket, and check the one against it. 
  //if it's not the same, move on. if is the same, highlight the row and the one preceding
  var rangeToHighlight = []; 
  for(var i = 1; i < bucket.length; i++) {
    if(bucket[i].amount == amountForCurrentItem) {
      rangeToHighlight.push(rawDataSheet.getRange(bucket[i].rowId, 1, 1, rawDataSheet.getLastColumn()).getA1Notation()); 
      rangeToHighlight.push(rawDataSheet.getRange(bucket[i - 1].rowId, 1, 1, rawDataSheet.getLastColumn()).getA1Notation()); 
      //      highlightRow(bucket[i].rowId);
      //      highlightRow(bucket[i - 1].rowId);
      
    }
    amountForCurrentItem = bucket[i].amount; 
    
    
  }
  if(rangeToHighlight.length > 1) {
    rawDataSheet.getRangeList(rangeToHighlight).setBackground("#ffff00");
  }
  Logger.log(rangeToHighlight); 
}


//grabbing the rowId of the function that is fed to it by the "identifyDuplicates" function and highlighting the index of the row on the raw Data sheet
//function highlightRow(rowId) {
//  
//  var rowsToHighlight = []; 
//  var range = rowsToHighlight.push(rawDataSheet.getRange(rowId, 1, 1, rawDataSheet.getLastColumn()).getA1Notation());
//  
//  rawDataSheet.getRangeList(rowsToHighlight).setBackground("#ffff00"); 
//}

function highlightRows(array) {
  var rangeList = rawDataSheet.getRangeList(array).setBackground("#ffff00"); 
}

function ignorePrevious (newDocs, oldDocs) {
  var start_row = 11;
  var docNumIndex = 13; 
  var newDocNumSheet = spreadsheet.getSheetByName(EXTRACT_SHEET);
  var newDocRange = newDocNumSheet.getRange(start_row, 14, newDocNumSheet.getLastRow()).getValues().flat(); 
  var oldDocNum = spreadsheet.getSheetByName("Previous Document Numbers").getDataRange().getValues().flat(); 
  
  var newSet = new Set(newDocRange); 
  var oldSet = new Set(oldDocNum); 
  
  //  Logger.log([...difference(newSet, oldSet)].length);   
  Logger.log(newSet.size); 
  
}

function difference(setA, setB) {
  let _difference = new Set(setA)
  for (let elem of setB) {
    _difference.delete(elem)
  }
  return _difference
}