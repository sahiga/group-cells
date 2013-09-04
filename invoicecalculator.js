function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Allocation Calculator", functionName: "invoiceCalculator"}];
  ss.addMenu("Calculator", menuEntries);
}

function invoiceCalculator() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Browser.msgBox("Allocation Calculator", "Welcome to the Allocation Calculator. We'll start by asking you about the cell ranges you want to use for your calculations. Enter your ranges in spreadsheet notation, e.g., F9:F37. Press OK to begin.", Browser.Buttons.OK);
  var projectInput = Browser.inputBox("Allocation Calculator", "Enter the range of cells under Projects:", Browser.Buttons.OK_CANCEL);
  var hoursInput = Browser.inputBox("Allocation Calculator", "Enter the range of cells under Hours:", Browser.Buttons.OK_CANCEL);
  var alloInput = Browser.inputBox("Allocation Calculator", "Enter the range of cells under Allocations that lists your projects:", Browser.Buttons.OK_CANCEL);
  var percentInput = Browser.inputBox("Allocation Calculator", "Enter the range of cells where you want your percentages:", Browser.Buttons.OK_CANCEL);
  var sumCell = Browser.inputBox("Allocation Calculator", "Enter the cell containing the number of hours you've worked:", Browser.Buttons.OK_CANCEL);
                            
  // getRange("F9:F37").getValues();
  // ("G9:G37").getValues();
  // ("F42:F48").getValues();
  // ("G42:G48");
  // ("G38").getValues();
  
  var projects = ss.getRange(projectInput).getValues();
  var hours = ss.getRange(hoursInput).getValues();
  var allocations = ss.getRange(alloInput).getValues();
  var percentages = ss.getRange(percentInput);
  var hoursSum = ss.getRange(sumCell).getValues();
  
  var projectsLen = ss.getRange(projectInput).getNumRows();
  var alloLen = ss.getRange(alloInput).getNumRows();
  
  var sum = new Array(alloLen);
  
/*  Browser.msgBox(projects.length + " & " +  allocations.length);
  
  
  Browser.msgBox(allocations);
  
  // setcell
 
  Browser.msgBox(projects[1][0] + " " + allocations[0][0]);
  
  Browser.msgBox(sum); */
  
  var n = 0
  
  for (n = 0; n < sum.length; n++) {
    sum[n] = new Array(1);
    sum[n][0] = 0;
  }
    
  /* Browser.msgBox(sum); */
  
  for (var i = 0; i < projectsLen; i++) {
    for (var j = 0; j < alloLen; j++) {
      if (projects[i][0] == allocations[j][0]) {
        
        /* percentages[j] += hours[i];
        percentages[j] = percentages[j]/hoursSum; 
        Browser.msgBox("percentages: " + percentages + " hours: " + hoursSum); */
        sum[j][0] = sum[j][0] + hours[i][0];
      }
      else { // do nothing
      }
    }
  }
  
  /* Browser.msgBox(sum); */
  
  for (n = 0; n < sum.length; n++) {
    sum[n][0] = (sum[n][0]/hoursSum[0][0]) * 100 + "%";
  }
  
  /* Browser.msgBox(sum); */
  
  percentages.setValues(sum);
  
  
}
  
  // ss.setNamedRange("Project", projectRange);
  
  /* .getRangeByName("Project");
  var hoursRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Hours"); 
  
  if (projects != null) {
    Browser.msgBox("number of columns: " + projects.getNumColumns() + " number of rows: " + projectsLen);
  }
  else {
    Browser.msgBox("nuthin'")
  }
  
  if (hours !=  null) {
    Browser.msgBox("sum: " + sum)
  }
  else {
    Browser.inputBox("???")
  }
  
/*  var prevCell;
  var projects = ['FixUp', 'Authorly', 'BrandReporter', 'HangPay', 'ShowKit', 'Curious Minds', 'StartupMinds']
  var i = 0;
  
  for (i = 0; i < projects.length; i++) {
    if cell == project[i] {
      
    }
*/     