function onOpen() {
  // allow user to run groupCompanies() sort method from spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = {name: "Sort (Group Companies)", functionName: "groupCompanies"}
  ss.addMenu("Scripts", menuEntries);
}

function groupCompanies() {
  // get current spreadsheet and current sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // grab all cells and their values; skip top row
  var readRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var readRangeCells = readRange.getValues();
  
  // initialize array to hold rows
  var rows = new Array(readRangeCells.length);
  
  // initialize array to hold sorted rows
  var sortedRows = new Array(0);
  
  // initialize temp and counter variables
  var prevCompany = "";
  var date = "";
  var i = 0;
  var num = 0;
  
  // turn "rows" into a 2d array
  for (i = 0; i < rows.length; i++) {
    rows[i] = new Array(0);
  }
  
  // sort rows in readRangeCells by column A (company)
  readRangeCells.sort();
  
  // turn "rows" into a 3d array
  // loop through rows in readRangeCells array
  for (i = 0; i < readRangeCells.length; i++) {
    // if column A is the same as previous row's column A
    // or if this is the first row in readRangeCells
    if (readRangeCells[i][0] == prevCompany || readRangeCells.indexOf(i) == 0) {
      // add the row to current member array of rows array
      rows[num].push(readRangeCells[i]);
    } else {
      // else add the row to the next member array of rows array
      num++;
      rows[num].push(readRangeCells[i]);
    }
    
    // if column A of current row is not empty
    if (readRangeCells[i][0] != "") {
      // save the contents of column A in prevCompany temp variable
      prevCompany = readRangeCells[i][0];
    } else {
      continue;
    }
  }
  
  // loop through each row array in "rows" (2d level)
  for (i = 0; i < rows.length; i++) {
    // loop through each row in row array (3d level)
    for (num = 0; num < rows[i].length; num++) {
      // if column E is not empty
      if (rows[i][num][4] != "") {
        // translate Google Spreadsheet date object into formatted string
        // this is the row's hidden label
        date = Utilities.formatDate(rows[i][num][4], "PST", "yyyy-MM-dd");
      // else default the row's hidden label to a number
      // that will appear after all specified years in the sort
      } else {
        date = 999999;
      }
      // add the hidden date label to the beginning of the row
      rows[i][num].unshift(date);
    }
    // sort each row array in "rows" by hidden date label
    rows[i].sort();
  }

  // sort "rows" by hidden date label
  rows.sort();

  // turn sortedRows into a 2d array
  // loop through each row array in "rows" (2d level)
  for (i = 0; i < rows.length; i++) {
    // loop through each row in row array (3d level)
    for (num = 0; num < rows[i].length; num++) {
      // remove hidden date label now that the sort is finished
      rows[i][num].shift(rows[i][num][0]);
      // if row is not empty
      if (rows[i][num] != ",,,,,,") {
        // add row to sortedRows
        sortedRows.push(rows[i][num]);
      } else {
        continue;
      }
    }
  }
  
  // clear content and formatting of readRange
  readRange.clear();
  
  // grab new range based on size of sortedRows
  var writeRange = sheet.getRange(2, 1, sortedRows.length, sortedRows[0].length);
  var writeRangeCells = writeRange.getValues();
  
  // write sortedRows to new range
  writeRange.setValues(sortedRows);
  
  // initialize color variables; reset prevCompany temp variable
  var colors = ["#ffffcc", "#bdf0f0"];
  var prevColor = colors[0];
  prevCompany = "";
  
  // loop through rows in writeRangeCells array
  for (i = 0; i < writeRangeCells.length; i++) {
    // if column A is the same as previous row's column A
    // or if this is the first row in writeRangeCells
    if (sortedRows[i][0] == prevCompany || sortedRows.indexOf(i) == 0) {
      // color the row with the current background color
      sheet.getRange(i+2, 1, 1, sortedRows[0].length).setBackgroundColor(prevColor);
    } else {
      // else switch the current background color
      if (prevColor == colors[0]) {
        prevColor = colors[1];
      } else {
        prevColor = colors[0];
      }
      // color the row with the new background color
      sheet.getRange(i+2, 1, 1, sortedRows[0].length).setBackgroundColor(prevColor);
    }
    
    // save contents of column A
    prevCompany = sortedRows[i][0];
  }
}