const spreadsheetId = ""; 
const sheetName = "Weekly";
const DATE_ROW = 2;         // Row index containing the Sunday dates
const DATA_START_ROW = 9;   // First row where numeric data entries begin
const SPEND_PERCENT_ROW = 3;
const WEEK_PERCENT_ROW = 4;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Week Data Entry')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Function called by the front-end form submission
function processForm(formObject) {
  const enteredNumber = parseFloat(formObject.numberInput);
  
  if (isNaN(enteredNumber)) {
    throw new Error("Please enter a valid number.");
  }
  
  let categoryValue = formObject.category;
  if (!categoryValue || categoryValue === "other") {
    categoryValue = ""; 
  }

  const metrics = appendToCurrentWeekColumn(enteredNumber, categoryValue);
  
  return {
    status: "Success! Added entry.",
    spendPercent: metrics.spendPercent,
    weekPercent: metrics.weekPercent
  };
}

// Helper function to find the target column for the current Sunday date
function getCurrentWeekColumn(sheet, ss) {
  const tz = ss.getSpreadsheetTimeZone();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const currentSunday = new Date(today);
  currentSunday.setDate(today.getDate() - today.getDay());
  const currentSundayStr = Utilities.formatDate(currentSunday, tz, "yyyy-MM-dd");
  
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  
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
  
  return targetColumn;
}

function getMetrics() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName); 
  
  const targetColumn = getCurrentWeekColumn(sheet, ss);
  
  if (targetColumn === -1) {
    return { "spendPercent": 0, "weekPercent": 0 };
  }
  
  const weekday = new Date().getDay() + 1; // 1 for Sunday, 7 for Saturday
  const weekPercent = Math.trunc((weekday / 7) * 1000) / 1000;

  return {
    "spendPercent": sheet.getRange(SPEND_PERCENT_ROW, targetColumn + 1).getValue(),
    "weekPercent": weekPercent
  };
}

function appendToCurrentWeekColumn(enteredNumber, categoryValue) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName); 
  
  const targetColumn = getCurrentWeekColumn(sheet, ss);
  
  if (targetColumn === -1) {
    const tz = ss.getSpreadsheetTimeZone();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const currentSunday = new Date(today);
    currentSunday.setDate(today.getDate() - today.getDay());
    const currentSundayStr = Utilities.formatDate(currentSunday, tz, "yyyy-MM-dd");
    
    throw new Error("Could not find a column for Sunday: " + currentSundayStr);
  }
  
  // Find the last filled cell in that column and append below it
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
  
  // Write number into target column, and category into the cell immediately to the right
  sheet.getRange(targetRow, targetColumn).setValue(enteredNumber);
  sheet.getRange(targetRow, targetColumn + 1).setValue(categoryValue);

  return getMetrics();
}
