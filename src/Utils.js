/**
 * Utils.js - helper functions
 */

function convertToMinutes(value, unit) {
  value = parseInt(value, 10);
  switch (unit) {
    case "minutes":
      return value;
    case "hours":
      return value * 60;
    case "days":
      return value * 24 * 60;
    case "weeks":
      return value * 7 * 24 * 60;
    default:
      return 0;
  }
}

function getCellValue(cellRef, sheetName = null, spreadsheetId = null) {
  try {
    let sheet;

    if (spreadsheetId && sheetName) {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      sheet = spreadsheet.getSheetByName(sheetName);
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

function deleteEvent(eventId) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      return { success: true };
    }
    return { success: false, error: "Event not found" };
  } catch (error) {
    console.error("Error deleting event:", error);
    return { success: false, error: error.toString() };
  }
}

function deleteReminderFromStorage(eventKey) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const events = JSON.parse(props.getProperty("events") || "{}");

    if (events[eventKey]) {
      delete events[eventKey];
      props.setProperty("events", JSON.stringify(events));
      return { success: true };
    }
    return { success: false, error: "Reminder not found in storage" };
  } catch (error) {
    console.error("Error deleting reminder from storage:", error);
    return { success: false, error: error.toString() };
  }
}
