/** 
  * Retrieves the FTE for cost center and calculates salary values
  * 
  * @param {costCenter} costCenter The cost center to lookup
  * @param {fteArray} fteArray Range of ALL FTE values <INCLUDE HEADER>
  * @param {salaryArray} salaryArray Range of salary values. Only select the job grade row and the salary row. 
  * @customfunction 
  */

function proRateLookup (costCenter, fteArray, salaryArray) { 
  //defining row index of job grade values are held
  //result of ["Job Grade", "A-OTH", "A-PFS", "B-OTH", "B-PFS", "C-OTH", "C-PFS", ...]
  var salaryGradeRow = salaryArray[0]; 
  
  
  //defining row index of actual salary values
  var salaryRow = salaryArray[1];
  
  //the fteArray parameter will have a 2d array of cost centers
  //we are only interested in the cost center, so we extract it and push it into this init array
  var ccResult = []; 
  for(var i = 0; i < fteArray.length; i++) {
    ccResult.push(fteArray[i][1])
  }
  
  //finding the index of input parameter (i.e. cost center) in the empty array we initialized
  var rowIndex = ccResult.indexOf(costCenter);
  
  //finding the row (i.e. index) of the cost center in the fteArray parameter using above index
  //will give a value of ["CFO Portoflio", "WAA0200", 0, 0, 1.0, etc...]
  var row = fteArray[rowIndex];  
  
  //defining where the paygrades are in the fteArray 
  //will give a result of ["C-OTH", ... , "L-OTH"] etc. 
  var payGradeRow = fteArray[0];
  
  //starting the loop on the row array
  var calculatedSalary = row.map(function (element, index) {
    
    //we want to exclude blank values and values that are not numbers (the cost center, any titles, etc.)
    //element = each item in the array, one at a time
    if(element != '' && !isNaN(element)) {
      
      //each element in the array is the actual fte value found
      //returns the numeric value of the fte
      var fte = element; 
      
      //the index of the element above on the paygrade row; 
      //will return a single pay grade e.g. "C-OTH"
      //index = where in the array it's iterating over. first element = 0, second = 1, etc...
      var payGrade = payGradeRow[index]; 
      var payGradeSalary = salaryRow[index]; 
      
      //takes the fte value and multiplies by whatever salary is retrieved
      return fte * payGradeSalary     
    } 
  })
  
  //if the fte value returns a undefined, it will still try to multiply it by the paygrade returning a "#N/A", so we filter out anything that is undefined; 
  var reduced = calculatedSalary.filter(function (element) {
    return element != undefined;   
  })
  
  //if the length of the array is not zero (i.e. there's something in it and it didn't just return all null values), reduce whatever is in the array
  if(reduced.length != 0) {
    return reduced.reduce(function (acc, val) {
      return acc + val; 
    }) 
  } else {
    return 0; 
  }
}





