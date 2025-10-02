/**
 * Cell Reminders - Code.js
 * Main entry points and Calendar integration
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Cell Reminders")
    .addItem("Add Cell Event", "showReminderSidebar")
    .addItem("View Cell Events", "listReminders")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();
}

function getActiveCellA1() {
  const range = SpreadsheetApp.getActiveRange();
  const sheet = SpreadsheetApp.getActiveSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (range) {
    return {
      cellRef: range.getA1Notation(),
      sheetName: sheet.getName(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetName: spreadsheet.getName(),
    };
  }
  return null;
}

function showReminderSidebar() {
  const html = HtmlService.createTemplateFromFile("ReminderForm").evaluate()
    .setTitle("Cell Reminders")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function createReminder(cellInfo, dueDate, message, isAllDay, repeatType, notification) {
  try {
    const eventResult = createCalendarEvent(
      message,
      dueDate,
      isAllDay,
      repeatType,
      cellInfo,
      notification
    );
    if (!eventResult.success) return { success: false, error: eventResult.error };

    const props = PropertiesService.getDocumentProperties();
    const events = JSON.parse(props.getProperty("events") || "{}");

    const eventKey = `${cellInfo.spreadsheetId}_${cellInfo.sheetName}_${cellInfo.cellRef}`;
    events[eventKey] = {
      message,
      dueDate,
      isAllDay,
      repeatType,
      notification,
      eventId: eventResult.eventId,
      cellInfo,
      createdAt: new Date().toISOString(),
    };

    props.setProperty("events", JSON.stringify(events));
    return { success: true, eventId: eventResult.eventId };
  } catch (error) {
    console.error("Error creating event:", error);
    return { success: false, error: error.toString() };
  }
}

function listReminders() {
  const props = PropertiesService.getDocumentProperties();
  const events = JSON.parse(props.getProperty("events") || "{}");

  let html = "<h3>Existing Events</h3>";
  if (Object.keys(events).length === 0) {
    html += "<p>No events found.</p>";
  } else {
    html += "<ul style='list-style:none;padding:0;'>";
    for (const key in events) {
      const r = events[key];
      const cellDisplay = `${r.cellInfo.sheetName}!${r.cellInfo.cellRef}`;
      const dueDisplay = r.isAllDay
        ? new Date(r.dueDate).toLocaleDateString()
        : new Date(r.dueDate).toLocaleString();
      const repeatDisplay = r.repeatType !== "none" ? ` (${r.repeatType})` : "";
      const notifDisplay = r.notification
        ? `<br><small>Notify: ${r.notification.value} ${r.notification.unit} before</small>`
        : "";

      html += `<li style="margin-bottom:10px;padding:8px;border:1px solid #ddd;border-radius:4px;">
        <strong>${cellDisplay}</strong><br>
        ${r.message}<br>
        <small>Due: ${dueDisplay}${repeatDisplay}</small>
        ${notifDisplay}
      </li>`;
    }
    html += "</ul>";
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300),
    "Events"
  );
}

function showHelp() {
  const helpHtml = `
    <h3>Cell Reminders Help</h3>
    <ol>
      <li>Select a cell in your spreadsheet</li>
      <li>Go to "Cell Reminders" > "Add Event"</li>
      <li>Fill in the event message (defaults to cell content)</li>
      <li>Choose all-day or time-based</li>
      <li>Set repeat & notifications</li>
    </ol>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(helpHtml).setWidth(400).setHeight(350),
    "Help"
  );
}

/**
 * Calendar creation
 */
function createCalendarEvent(title, dueDate, isAllDay, repeatType, cellInfo, notification) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    let start, end, event;

    if (isAllDay) {
      const date = new Date(dueDate);
      start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      end = new Date(start);
      end.setDate(end.getDate() + 1);
    } else {
      start = new Date(dueDate);
      end = new Date(start);
      end.setMinutes(end.getMinutes() + 30);
    }

    let description = `Created from ${cellInfo.spreadsheetName} - ${cellInfo.sheetName}!${cellInfo.cellRef}`;
    if (repeatType !== "none") description += `\nRepeat: ${repeatType}`;

    if (repeatType === "none") {
      event = isAllDay
        ? calendar.createAllDayEvent(title, start, { description })
        : calendar.createEvent(title, start, end, { description });
    } else {
      let recurrence;
      switch (repeatType) {
        case "daily": recurrence = CalendarApp.newRecurrence().addDailyRule().times(100); break;
        case "weekly": recurrence = CalendarApp.newRecurrence().addWeeklyRule().times(100); break;
        case "monthly": recurrence = CalendarApp.newRecurrence().addMonthlyRule().times(100); break;
        case "yearly": recurrence = CalendarApp.newRecurrence().addYearlyRule().times(100); break;
      }
      event = isAllDay
        ? calendar.createAllDayEventSeries(title, start, recurrence, { description })
        : calendar.createEventSeries(title, start, end, recurrence, { description });
    }

    if (notification && notification.value && notification.unit) {
      const minutesBefore = convertToMinutes(notification.value, notification.unit);
      if (minutesBefore > 0) event.addPopupReminder(minutesBefore);
    }

    return { success: true, eventId: event.getId() };
  } catch (error) {
    console.error("Calendar error:", error);
    return { success: false, error: error.toString() };
  }
}
