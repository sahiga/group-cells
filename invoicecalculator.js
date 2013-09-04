// Get active sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();

// Add calculator to menu when user opens spreadsheet
function onOpen() {
  var menuEntries = [{name: 'Invoice Allocation Calculator', functionName: 'runCalculator'}];
  ss.addMenu('Calculator', menuEntries);
}

function runCalculator() {
  // Create new User Interface application
  var app = UiApp.createApplication().setTitle('Invoice Allocation Calculator');
  
  // Create simple HTML introduction
  var intro = app.createHTML("<strong>Welcome to the Invoice Allocation Calculator.</strong> Welcome to the Allocation Calculator. We'll start by asking you about the cell ranges you want to use for your calculations. Enter your ranges in spreadsheet notation, e.g., F9:F37.");
  
  // Create grid of input elements: this is where users will enter cell ranges
  var grid = app.createGrid(5, 2);
  grid.setWidget(0, 0, app.createLabel('Enter the range of cells under Projects:'));
  grid.setWidget(0, 1, app.createTextBox().setName('projectsRange'));
  grid.setWidget(1, 0, app.createLabel('Enter the range of cells under Hours:'));
  grid.setWidget(1, 1, app.createTextBox().setName('hoursRange'));
  grid.setWidget(2, 0, app.createLabel('Enter the range of cells under Allocations that lists your projects:'));
  grid.setWidget(2, 1, app.createTextBox().setName('allocationsRange'));
  grid.setWidget(3, 0, app.createLabel('Enter the range of cells where you want your percentages:'));
  grid.setWidget(3, 1, app.createTextBox().setName('percentagesRange'));
  grid.setWidget(4, 0, app.createLabel("Enter the cell containing the number of hours you've worked:"));
  grid.setWidget(4, 1, app.createTextBox().setName('totalHoursCell'));
  
  // Create submit button and click handler
  // grid is the callback element
  // function calculatePercentages() is the click handler
  var handler = app.createServerHandler('calculatePercentages');
  handler.addCallbackElement(grid);
  
  var submit = app.createButton('Calculate allocation percentages');
  submit.addClickHandler(handler)
  // Define styles for submit button
  .setStyleAttributes({
    background: 'rgb(255, 140, 134)',
    border: '1px solid rgb(255, 144, 92)',
    color: '#fff',
    padding: '10px 40px',
    margin: '20px auto',
    display: 'block',
    fontWeight: '800',
    borderRadius: '20px'
  });
  
  // Create panel
  var panel = app.createVerticalPanel();

  // Define styles for panel
  panel.setStyleAttributes({
    background: 'rgb(226, 234, 255)',
    padding: '10px'
  })
  // Add introduction, grid, and submit button to panel
  .add(intro).add(grid).add(submit);
  
  // Add panel to app
  app.add(panel);
  
  // Show app in current sheet
  ss.show(app);
}
  
function calculatePercentages(e) {
  // Grab user input from named cells within the UiApp grid 
  var projectsRange = e.parameter.projectsRange;
  var hoursRange = e.parameter.hoursRange;
  var allocationsRange = e.parameter.allocationsRange;
  var percentagesRange = e.parameter.percentagesRange;
  var totalHoursCell = e.parameter.totalHoursCell;
  
  // Translate user input into ranges in the current sheet, and grab values from these ranges
  var projects = ss.getRange(projectsRange).getValues();
  var hours = ss.getRange(hoursRange).getValues();
  var allocations = ss.getRange(allocationsRange).getValues();
  var percentages = ss.getRange(percentagesRange);
  var totalHours = ss.getRange(totalHoursCell).getValues();
  
  // Count lengths of projects and allocations
  var numProjects = ss.getRange(projectsRange).getNumRows();
  var numAllocations = ss.getRange(allocationsRange).getNumRows();
  
  // Initialize array to hold number of hours worked on each project
  var sum = new Array(numAllocations);
  
  // Initialize counter for sum array
  var count = 0;
  
  // Set the number of hours worked on each project to 0
  // getValues(), which we used to grab the ranges above, returns a 2D array of values indexed by row and then column
  // See https://developers.google.com/apps-script/reference/spreadsheet/range?hl=en#getValues()
  // Therefore, to set the number of hours on each project to 0, we must turn the sum array into a 2D array
  for (count = 0; count < sum.length; count++) {
    sum[count] = new Array(1);
    sum[count][0] = 0;
  }
  
  // For each line in projects
  for (var i = 0; i < numProjects; i++) {
    // Loop through each line in allocations
    for (var j = 0; j < numAllocations; j++) {
      // If the current project name is the same as the current name listed under allocations
      if (projects[i][0] === allocations[j][0]) {
        // Add the corresponding number of hours to the sum for that project
        sum[j][0] = sum[j][0] + hours[i][0];
      }
    }
  }
  
  // For each member of the sum array 
  for (count = 0; count < sum.length; count++) {
    // Divide it by the number of hours worked and turn it into a percentage
    sum[count][0] = (sum[count][0]/totalHours[0][0]) * 100 + "%";
  }
  
  // Print percentages in the appropriate range of the sheet
  percentages.setValues(sum);
}