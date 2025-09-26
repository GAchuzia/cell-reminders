function showReminderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ReminderForm")
    .setTitle("Set Reminder");
  SpreadsheetApp.getUi().showSidebar(html);
}


function listReminders() {
  const props = PropertiesService.getDocumentProperties();
  const reminders = JSON.parse(props.getProperty("reminders") || "{}");
  let listHtml = "<h3>Existing Reminders</h3><ul>";

  for (const cell in reminders) {
    const r = reminders[cell];
    listHtml += `<li><b>${cell}</b>: ${r.message} (due: ${r.dueDate})</li>`;
  }
  listHtml += "</ul>";

  const html = HtmlService.createHtmlOutput(listHtml).setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, "Reminders");
}
