function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Reminders")
    .addItem("Add Reminder", "showReminderSidebar")
    .addItem("View Reminders", "listReminders")
    .addToUi();
}

// Get currently selected cell
function getActiveCellA1() {
  const range = SpreadsheetApp.getActiveRange();
  return range ? range.getA1Notation() : "";
}

// Show sidebar
function showReminderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ReminderForm")
    .setTitle("Set Reminder")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Create a reminder and calendar event
function createReminder(cellRef, dueDate, message) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cal = CalendarApp.getDefaultCalendar();

  const start = new Date(dueDate);
  const end = new Date(start.getTime() + 30 * 60 * 1000); // 30 mins
  const event = cal.createEvent(message, start, end);

  const props = PropertiesService.getDocumentProperties();
  const reminders = JSON.parse(props.getProperty("reminders") || "{}");
  reminders[cellRef] = { message, dueDate, eventId: event.getId() };
  props.setProperty("reminders", JSON.stringify(reminders));
}

// List all reminders
function listReminders() {
  const props = PropertiesService.getDocumentProperties();
  const reminders = JSON.parse(props.getProperty("reminders") || "{}");

  let html = "<h3>Existing Reminders</h3><ul>";
  for (const cell in reminders) {
    const r = reminders[cell];
    html += `<li><b>${cell}</b>: ${r.message} (due: ${r.dueDate})</li>`;
  }
  html += "</ul>";

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(output, "Reminders");
}
