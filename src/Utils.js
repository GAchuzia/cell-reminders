function validateDate(dateString) {
  if (!dateString) {
    return { isValid: false, error: "Date is required" };
  }

  const date = new Date(dateString);
  const now = new Date();

  if (isNaN(date.getTime())) {
    return { isValid: false, error: "Invalid date format" };
  }

  if (date <= now) {
    return { isValid: false, error: "Due date must be in the future" };
  }

  return { isValid: true, date: date };
}

function formatDate(date) {
  const d = new Date(date);
  return (
    d.toLocaleDateString() +
    " " +
    d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
  );
}

function validateCellReference(cellRef) {
  if (!cellRef) return false;
  const regex = /^[A-Z]+[0-9]+$/;
  return regex.test(cellRef);
}

function getCellValue(cellRef) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange(cellRef);
    return range.getValue();
  } catch (error) {
    console.error("Error getting cell value:", error);
    return "";
  }
}

function getAllReminders() {
  const props = PropertiesService.getDocumentProperties();
  const remindersJson = props.getProperty("reminders");
  return remindersJson ? JSON.parse(remindersJson) : {};
}

function saveReminders(reminders) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("reminders", JSON.stringify(reminders));
}

function generateReminderId() {
  return (
    "reminder_" + Date.now() + "_" + Math.random().toString(36).substr(2, 9)
  );
}

function isCalendarAvailable() {
  try {
    CalendarApp.getDefaultCalendar();
    return true;
  } catch (error) {
    console.error("Calendar not available:", error);
    return false;
  }
}

function createCalendarEvent(title, startTime, endTime, description = "") {
  try {
    if (!isCalendarAvailable()) {
      return { success: false, error: "Google Calendar is not available" };
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.createEvent(title, startTime, endTime, {
      description: description,
    });

    return { success: true, eventId: event.getId() };
  } catch (error) {
    console.error("Error creating calendar event:", error);
    return { success: false, error: error.toString() };
  }
}

function deleteCalendarEvent(eventId) {
  try {
    if (!isCalendarAvailable()) {
      return false;
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      return true;
    }
    return false;
  } catch (error) {
    console.error("Error deleting calendar event:", error);
    return false;
  }
}

function showToast(message, title = "Cell Reminders") {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
}
