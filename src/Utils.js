/**
 * Utils.js - helper functions
 */

function convertToMinutes(value, unit) {
  value = parseInt(value, 10);
  switch (unit) {
    case "minutes": return value;
    case "hours": return value * 60;
    case "days": return value * 24 * 60;
    case "weeks": return value * 7 * 24 * 60;
    default: return 0;
  }
}

function getCellValue(cellRef, sheetName = null, spreadsheetId = null) {
  try {
    let sheet;
    if (spreadsheetId && sheetName) {
      sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    } else if (sheetName) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    } else {
      sheet = SpreadsheetApp.getActiveSheet();
    }
    if (!sheet) return "";
    return sheet.getRange(cellRef).getValue();
  } catch (error) {
    console.error("Error getting cell value:", error);
    return "";
  }
}
