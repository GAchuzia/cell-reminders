/**
 * Cell Reminders - Code.js
 * Main entry points and Calendar integration
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Cell Reminders")
    .addItem("Add Event", "showReminderSidebar")
    .addItem("View Events", "listReminders")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();
}

// Core

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
  const html = HtmlService.createHtmlOutput(getReminderFormHtml())
    .setTitle("Cell Reminders")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function createReminder(
  cellInfo,
  dueDate,
  message,
  isAllDay = false,
  repeatType = "none",
  notification = null
) {
  try {
    const eventResult = createCalendarEvent(
      message,
      dueDate,
      isAllDay,
      repeatType,
      cellInfo,
      notification
    );

    if (!eventResult.success)
      return { success: false, error: eventResult.error };

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
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div style="flex:1;">
            <strong>${cellDisplay}</strong><br>
            ${r.message}<br>
            <small>Due: ${dueDisplay}${repeatDisplay}</small>
            ${notifDisplay}
          </div>
          <button onclick="google.script.run.deleteEventFromList('${key}', '${r.eventId}')" style="background:#dc3545;color:white;border:none;padding:4px 8px;border-radius:3px;cursor:pointer;font-size:11px;">Delete</button>
        </div>
      </li>`;
    }
    html += "</ul>";
  }

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Events");
}

function showHelp() {
  const helpHtml = `
    <h3>Cell Reminders Help</h3>
    <p><strong>How to use:</strong></p>
    <ol>
      <li>Select a cell in your spreadsheet</li>
      <li>Go to "Cell Reminders" > "Add Event"</li>
      <li>Fill in the event message (defaults to cell content)</li>
      <li>Choose if it's an all-day event or specific time</li>
      <li>Set due date/time</li>
      <li>Optionally set it to repeat</li>
      <li>Optionally set a notification reminder</li>
      <li>Click "Add Event"</li>
    </ol>
    <p><strong>Features:</strong></p>
    <ul>
      <li>Creates events in Google Calendar</li>
      <li>Works with any Google Sheet</li>
      <li>Supports all-day and timed events</li>
      <li>Repeat options: daily, weekly, monthly, yearly</li>
      <li>Custom notifications</li>
      <li>Delete events from the list</li>
    </ul>
  `;

  const output = HtmlService.createHtmlOutput(helpHtml)
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Help");
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

// Event Management

function deleteEventFromList(eventKey, eventId) {
  try {
    // Delete from Google Calendar
    const deleteResult = deleteEvent(eventId);
    if (!deleteResult.success) {
      console.error("Failed to delete calendar event:", deleteResult.error);
    }

    // Delete from storage
    const storageResult = deleteReminderFromStorage(eventKey);
    if (!storageResult.success) {
      console.error("Failed to delete from storage:", storageResult.error);
    }

    // Show success message and refresh the list
    showSuccessMessage("Event deleted!");
    setTimeout(() => {
      listReminders();
    }, 1500);

    return { success: true };
  } catch (error) {
    console.error("Error deleting event:", error);
    return { success: false, error: error.toString() };
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

function showSuccessMessage(message) {
  const html = `
    <div style="text-align: center; padding: 20px;">
      <div style="color: #388e3c; font-size: 18px; font-weight: bold; margin-bottom: 10px;">
        ✓ ${message}
      </div>
      <div style="color: #666; font-size: 14px;">
        Feel free to exit out of this message.
      </div>
    </div>
  `;

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(output, "Success");
}

// Google Calendar API Integration

function createCalendarEvent(
  title,
  dueDate,
  isAllDay,
  repeatType,
  cellInfo,
  notification
) {
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

    let eventOptions = { description };

    if (repeatType === "none") {
      event = isAllDay
        ? calendar.createAllDayEvent(title, start, eventOptions)
        : calendar.createEvent(title, start, end, eventOptions);
    } else {
      let recurrence;
      switch (repeatType) {
        case "daily":
          recurrence = CalendarApp.newRecurrence().addDailyRule().times(100);
          break;
        case "weekly":
          recurrence = CalendarApp.newRecurrence().addWeeklyRule().times(100);
          break;
        case "monthly":
          recurrence = CalendarApp.newRecurrence().addMonthlyRule().times(100);
          break;
        case "yearly":
          recurrence = CalendarApp.newRecurrence().addYearlyRule().times(100);
          break;
      }
      event = isAllDay
        ? calendar.createAllDayEventSeries(
            title,
            start,
            recurrence,
            eventOptions
          )
        : calendar.createEventSeries(
            title,
            start,
            end,
            recurrence,
            eventOptions
          );
    }

    if (notification && notification.value && notification.unit) {
      const minutesBefore = convertToMinutes(
        notification.value,
        notification.unit
      );
      if (minutesBefore > 0) {
        event.addPopupReminder(minutesBefore);
      }
    }

    return { success: true, eventId: event.getId() };
  } catch (error) {
    console.error("Error creating Calendar event:", error);
    return { success: false, error: error.toString() };
  }
}

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