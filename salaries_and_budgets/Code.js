const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
const sourceSheet = spreadsheet.getSheetByName("People Moves"); 
const transferTemplate = spreadsheet.getSheetByName("Budget Transfer"); 
const lookupValues = spreadsheet.getSheetByName("S&B Inputs");  
const venaExport = spreadsheet.getSheetByName("Vena Export"); 

var proRateValues = buildProRateTable(); 
var salaryTable = salaryGradeTable(); 
var inputRate = inputRates(); 
var discretionBudget = discretionaryBudget (); 

function getPeopleMoves () {
  
  var requestReportFor = sourceSheet.getRange(2, 4).getDisplayValue(); 
  var ui = SpreadsheetApp.getUi()
  var response = ui.alert("Generate report for " + requestReportFor + "? Note: This will clear the Vena Export tab. Proceed?", ui.ButtonSet.OK_CANCEL)
  
  if(response == ui.Button.OK) {
    venaExport.clear();
   spreadsheet.toast("Calculating...", "Please wait", 300)
  } else {
    ui.alert("Operation cancelled");   
  }
  
  var completedMonthIndex = 0;
  var effectiveMonthIndex = 1; 
  var fromCCIndex = 6; 
  var toCCIndex = 8; 
  var fte = 10; 
  var payGradeIndex = 11; 
  var annualSalaryIndex = 12;
  var vacancyAllowanceIndex = 13; 
  var stipIndex = 14;  
  var ltipIndex = 15; 
  var BenefitsIndex = 16; 
  var travelIndex = 19; 
  var staffDevIndex = 20; 
  var entertainmentIndex = 21; 
  var teleComIndex = 22; 
  var commentIndex = 23; 
  var startRow = 4; 
  
  var range = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()); 
  var values = range.getValues();

  
  var filtered = values.filter(row => row[0] == requestReportFor); 

  var glRow = values[2]
  .filter(gl => gl != "")
  .map(gl => gl.toString()); 

  var templateArray = []; 
  templateArray.push(["CC", "GL", "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"])
  
  
  for(var i = 0; i < filtered.length; i++) {
    var row = filtered[i];
    var effectiveMonth = row[effectiveMonthIndex]; 
    var completedMonth = row[completedMonthIndex]; 
    var fteRate = row[fte]
 
    var fromCc = row[fromCCIndex]; 
    var toCc = row[toCCIndex]; 
    var payGrade = row[payGradeIndex]; 
    
    var salaryIndex = 0; 
    var stipIndex = 1; 
    var salary = getSalary(payGrade)[salaryIndex] * fteRate; 
    var vacAllow = vacancyAllowance(salary); 
    var stip = stipOnDC(payGrade) * fteRate; 
    var ltip = row[ltipIndex] || 0
    var benefits = benefit(salary); 
    var travel = discretionBudget["Travel"];
    var staffDev = discretionBudget["Staff Development"]; 
    var entertainment = discretionBudget["Entertainment"]; 
    var telecom = discretionBudget["Telecom"]; 
    
    var calculatedRow = 
        [
          salary, 
          vacAllow, 
          stip, 
          ltip, 
          benefits, 
          travel, 
          staffDev, 
          entertainment, 
          telecom
        ]
    
    calculatedRow.forEach(function (item, index) {
      var from = monthlyProRate(effectiveMonth, completedMonth, glRow[index], item * -1);
      from.unshift(fromCc, glRow[index]);
      templateArray.push(from); 
    })
    
    calculatedRow.forEach(function (item, index) {
      var to = monthlyProRate (effectiveMonth, completedMonth, glRow[index], item); 
      to.unshift(toCc, glRow[index]);
      templateArray.push(to); 
    })
    
    }
  venaExport.getRange(1, 1, templateArray.length, templateArray[0].length).setValues(templateArray); 
  spreadsheet.toast("Complete!", "", -1); 
}

function buildProRateTable () {
  
  // indexing rows and columns for range to work with
  var proRateStart = 24; 
  var proRateEnd = 36; 
  var noColumns = 8;
  var numRows = proRateEnd - proRateStart; 
  
  var range = lookupValues.getRange(proRateStart, 1, numRows, noColumns); 
  var proRateValues = range.getValues(); 
  
  var monthIndex = 0; 
  var workDaysIndex = 1; 
  var workDaysAmortIndex = 2; 
  
  var workDaysRemainingIndex = 3; 
  var workDaysRemainingPercentIndex = 4; 
  
  var straightLineMonthIndex = 5; 
  var straightLineRemainingIndex = 6; 
  var straightLinePercentIndex = 7; 
  
  var proRateMap = new Map(); 
  
  for(var i = 0; i < numRows; i++) {
    var row = proRateValues[i]; 
    var month = row[monthIndex]; 
    var workDays = row[workDaysIndex]; 
    var workDaysAmort = row[workDaysAmortIndex]; 
    
    var workDaysRemaining = row[workDaysRemainingIndex]; 
    var workDaysRemainingPercent = row[workDaysRemainingPercentIndex]; 
    
    var straightLineMonth = row[straightLineMonthIndex]; 
    var straightLineRemaining = row[straightLineRemainingIndex]; 
    var straightLinePercent = row[straightLinePercentIndex]; 
    
    proRateMap.set(month, {
      "workDays": Number(workDays), 
      "workDaysAmort": Number(workDaysAmort), 
      "workDaysRemaining": workDaysRemaining,
      "workDaysRemainingPercent": workDaysRemainingPercent, 
      "straightLineMonth": straightLineMonth, 
      "straightLineMonthRemaining": straightLineRemaining, 
      "straightLine": Number(straightLinePercent)     
    }); 
    
    
  }  
  return proRateMap;
}

function getSalary (payGrade) {
  var get = salaryTable.get(payGrade);
  var salary = get.salary; 
  var stip = salary * get.stipRate
  
  return [salary, stip]; 
}

function vacancyAllowance (salary) {
  return (salary * inputRate["Vacancy Rate"]) * -1;
} 

function stipOnDC (payGrade) {
  var getInfo = getSalary(payGrade);  
  var salaryInd = 0; 
  var stipInd = 1; 
  
  return getInfo[stipInd] + (getInfo[stipInd] * inputRate["DC on STIP"]); 
}

function benefit (salary) {
  return salary * inputRate["Benefits Rate"]; 
}

function salaryGradeTable () {
  var grade = 0; 
  var salary = 1; 
  var stip = 2; 
  var hourly = 3; 
  var billing = 4; 
  
  var salaryStartRange = 9; 
  var salaryEndRange = 20; 
  var salaryNumCol = 5; 
  
  var range = lookupValues.getRange(salaryStartRange, 1, salaryEndRange - salaryStartRange, salaryNumCol); 
  var salaryValues = range.getValues(); 

  var salaryMap = new Map (); 
  
  for(var i = 0; i < salaryValues.length; i++) {
    var payGrade = salaryValues[i][grade]; 
    var salaryRate = salaryValues[i][salary]; 
    var stipRate = salaryValues[i][stip]; 
    var hourlyCost = salaryValues[i][hourly]; 
    var billingRate = salaryValues[i][billing]; 

    salaryMap.set(payGrade, {
      "salary": salaryRate,
      "stipRate": stipRate, 
      "hourlyCost": hourlyCost, 
      "billingRate": billingRate    
    })
  }
  
  return salaryMap; 
}

function inputRates () {
  var values = lookupValues.getDataRange().getValues(); 
  var inputStart = 0; 
  var inputEnd = 7; 
  var inputHeaderLength = 2; 
 
  var obj = {};
  
  for(var currentRow = inputStart; currentRow < inputEnd; currentRow++) {
    obj[values[currentRow][0]] = values[currentRow][1]; 
  }
  
  return obj; 
  
}

function discretionaryBudget () {
  var values = lookupValues.getDataRange().getValues(); 
  var discStart = 40; 
  var discEnd = 44; 
  
  var descriptionIndex = 0; 
  var valueIndex = 1; 
  
  var headerLength = 2; 
  var discretionaryObj = {}; 
  
  for(var currentRow = discStart; currentRow < discEnd; currentRow++) {
    discretionaryObj[values[currentRow][descriptionIndex]] = values[currentRow][valueIndex]; 
  }
  
  return discretionaryObj

}


function glTable () {
  var values = lookupValues.getDataRange().getValues(); 
  var glStart = 46; 
  var glEnd = 55; 
  var glObj = {}; 
  var glIndex = 0; 
  var descriptionIndex = 1; 
  var amortMethodIndex = 2; 
  
  for(var i = glStart; i < glEnd; i++) {
    var row =  values[i];   
    
    glObj[row[glIndex].toString()] = {
      "description": row[descriptionIndex],
      "amortMethod": row[amortMethodIndex]    
    }
  }
  
  return glObj; 

}

function getProRate (gl, month) {
//  gl = 74111000
//  month = "June"
  var glLookup = glTable(); 
  var amortMethod = glLookup[gl]["amortMethod"]; 
  var rate = 0; 
  
  if(amortMethod == "Workdays") {
    rate = proRateValues.get(month)["workDaysAmort"]
  } else {
    rate = proRateValues.get(month)["straightLine"]
  }

  return rate; 
  
}

function monthlyProRate (effectiveMonth, completedMonth, gl, value) {
//  var value = 5857.575; 
//  var completedMonth = "June"; 
//  var effectiveMonth = "May"; 
//  var gl = "74111300"; 
  
  //init a months object to loop through
  var months = {
    "April": 0,
    "May": 0, 
    "June": 0,
    "July": 0,
    "August": 0,
    "September": 0, 
    "October": 0,
    "November": 0, 
    "December": 0, 
    "January": 0,
    "February": 0, 
    "March": 0, 
  } 
  
  //array of keys [april, may, june etc.]
  var keys = Object.keys(months); 
  
  //index of the completed and effective month in the array above. this will determine where to start and end looping
  var completedMonthIndex = keys.indexOf(completedMonth); 
  var effectiveMonthIndex = keys.indexOf(effectiveMonth); 
  var rate = 0; 
  
  for (var i = completedMonthIndex; i < keys.length; i++) {
    var month = keys[i]
    var rate = getProRate(gl, month); 
    if((month == completedMonth) && (completedMonth != effectiveMonth)) {
      for(var k = effectiveMonthIndex; k < completedMonthIndex; k++) {
        months[keys[k]] = 0; 
        rate += getProRate(gl, keys[k]); 
      }
      months[month] = value * rate; 
    } else {
      months[month] = value * rate; 
    }
    
  }
  
  var results = [];   
  
  for (var i in months) {
    results.push(months[i]);   
    
  }
  return results; 
  
}

function getKeyByValue(object, value) {
  return Object.keys(object).find(key => object[key] === value);
}
