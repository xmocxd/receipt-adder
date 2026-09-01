const spreadsheetId = ""; 
const sheetName = "Weekly";
const DATE_ROW = 2;         // Row index containing the Sunday dates
const DATA_START_ROW = 9;   // First row where numeric data entries begin


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Week Data Entry')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Function called by the front-end form submission
function processForm(formObject) {
  const enteredNumber = parseFloat(formObject.numberInput);
  
  if (isNaN(enteredNumber)) {
    throw new Error("Please enter a valid number.");
  }
  
  // Extract category and format rule: 'other' translates to blank ""
  let categoryValue = formObject.category;
  if (!categoryValue || categoryValue === "other") {
    categoryValue = ""; 
  }
  
  appendToCurrentWeekColumn(enteredNumber, categoryValue);
  return "Success! Added entry.";
}

function testAppend() {
  appendToCurrentWeekColumn(42, "Grocery"); // Test function
}

function appendToCurrentWeekColumn(enteredNumber, categoryValue) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName); 
  
  const tz = ss.getSpreadsheetTimeZone();
  
  // 1. Calculate the Sunday date for the current week and format as YYYY-MM-DD
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const currentSunday = new Date(today);
  currentSunday.setDate(today.getDate() - today.getDay());
  const currentSundayStr = Utilities.formatDate(currentSunday, tz, "yyyy-MM-dd");
  
  // 2. Search the date row
  const lastCol = sheet.getLastColumn();
  const dateRowValues = sheet.getRange(DATE_ROW, 1, 1, lastCol).getValues()[0];
  
  let targetColumn = -1;
  
  for (let i = 0; i < dateRowValues.length; i++) {
    const cellValue = dateRowValues[i];
    
    if (cellValue !== "" && cellValue !== null) {
      let cellDateStr = "";
      
      if (cellValue instanceof Date) {
        cellDateStr = Utilities.formatDate(cellValue, tz, "yyyy-MM-dd");
      } else {
        const parsedDate = new Date(cellValue);
        if (!isNaN(parsedDate.getTime())) {
          cellDateStr = Utilities.formatDate(parsedDate, tz, "yyyy-MM-dd");
        }
      }
      
      if (cellDateStr === currentSundayStr) {
        targetColumn = i + 1; 
        break;
      }
    }
  }
  
  if (targetColumn === -1) {
    throw new Error("Could not find a column for Sunday: " + currentSundayStr);
  }
  
  // 3. Find the last filled cell in that column and append below it
  const lastRow = sheet.getLastRow();
  const numRowsToCheck = Math.max(1, lastRow - DATA_START_ROW + 1);
  const columnValues = sheet.getRange(DATA_START_ROW, targetColumn, numRowsToCheck, 1).getValues();
  
  let targetRow = DATA_START_ROW;
  
  for (let j = columnValues.length - 1; j >= 0; j--) {
    if (columnValues[j][0] !== "" && columnValues[j][0] !== null) {
      targetRow = DATA_START_ROW + j + 1;
      break;
    }
  }
  
  // 4. Write number into target column, and category into the cell immediately to the right
  sheet.getRange(targetRow, targetColumn).setValue(enteredNumber);
  sheet.getRange(targetRow, targetColumn + 1).setValue(categoryValue);
}
