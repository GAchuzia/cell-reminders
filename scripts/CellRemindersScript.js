/**
 * Cell Reminders - Developer Script Version
 * Author: Grant Achuzia
 *
 * This is a standalone Google Apps Script that can be deployed immediately
 * without going through the Google Workspace Marketplace.
 *
 * Instructions:
 * 1. Go to script.google.com
 * 2. Create a new project
 * 3. Replace the default Code.gs content with this entire file
 * 4. Save and authorize the script
 * 5. Open any Google Sheet and refresh - you'll see the "Cell Reminders" menu
 *
 * Required OAuth Scopes (will be requested automatically):
 * - https://www.googleapis.com/auth/spreadsheets
 * - https://www.googleapis.com/auth/calendar
 * - https://www.googleapis.com/auth/script.container.ui
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Cell Reminders")
    .addItem("Add Reminder", "showReminderSidebar")
    .addItem("View Events", "listReminders")
    .addItem("View Tasks", "listTasks")
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
      let repeatDisplay = "";
      if (r.repeatType !== "none") {
        if (
          typeof r.repeatType === "object" &&
          r.repeatType.type === "custom"
        ) {
          repeatDisplay = ` (Every ${r.repeatType.interval} ${r.repeatType.frequency}`;
          if (r.repeatType.end.type === "after") {
            repeatDisplay += `, ${r.repeatType.end.count} times)`;
          } else if (r.repeatType.end.type === "on") {
            repeatDisplay += `, until ${r.repeatType.end.date})`;
          } else {
            repeatDisplay += ")";
          }
        } else {
          repeatDisplay = ` (${r.repeatType})`;
        }
      }
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
    let repeatDisplay = "none";
    if (repeatType !== "none") {
      if (typeof repeatType === "object" && repeatType.type === "custom") {
        repeatDisplay = `Every ${repeatType.interval} ${repeatType.frequency}`;
        if (repeatType.end.type === "after") {
          repeatDisplay += ` (${repeatType.end.count} times)`;
        } else if (repeatType.end.type === "on") {
          repeatDisplay += ` (until ${repeatType.end.date})`;
        }
      } else {
        repeatDisplay = repeatType;
      }
      description += `\nRepeat: ${repeatDisplay}`;
    }

    let eventOptions = { description };

    if (repeatType === "none") {
      event = isAllDay
        ? calendar.createAllDayEvent(title, start, eventOptions)
        : calendar.createEvent(title, start, end, eventOptions);
    } else {
      let recurrence = CalendarApp.newRecurrence();

      // Handle custom repeat
      if (typeof repeatType === "object" && repeatType.type === "custom") {
        const frequency = repeatType.frequency;
        const interval = repeatType.interval || 1;

        // Google Apps Script recurrence - intervals are set using the rule's interval method
        let rule;
        switch (frequency) {
          case "daily":
            rule = recurrence.addDailyRule();
            if (interval > 1) {
              rule.interval(interval);
            }
            break;
          case "weekly":
            rule = recurrence.addWeeklyRule();
            if (interval > 1) {
              rule.interval(interval);
            }
            break;
          case "monthly":
            rule = recurrence.addMonthlyRule();
            if (interval > 1) {
              rule.interval(interval);
            }
            break;
          case "yearly":
            rule = recurrence.addYearlyRule();
            if (interval > 1) {
              rule.interval(interval);
            }
            break;
        }

        // Handle end conditions
        if (repeatType.end.type === "after" && repeatType.end.count) {
          recurrence.times(repeatType.end.count);
        } else if (repeatType.end.type === "on" && repeatType.end.date) {
          const endDate = new Date(repeatType.end.date);
          endDate.setHours(23, 59, 59, 999); // End of day
          recurrence.until(endDate);
        } else {
          // Default to 100 occurrences if no end specified
          recurrence.times(100);
        }
      } else {
        // Handle simple repeat types
        switch (repeatType) {
          case "daily":
            recurrence.addDailyRule().times(100);
            break;
          case "weekly":
            recurrence.addWeeklyRule().times(100);
            break;
          case "monthly":
            recurrence.addMonthlyRule().times(100);
            break;
          case "yearly":
            recurrence.addYearlyRule().times(100);
            break;
        }
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

// Task Management

function createTask(
  cellInfo,
  dueDate,
  message,
  repeatType = "none",
  notification = null
) {
  try {
    // Create calendar event (tasks are all-day events)
    const eventResult = createCalendarEvent(
      message,
      dueDate,
      true, // isAllDay = true for tasks
      repeatType,
      cellInfo,
      notification
    );

    if (!eventResult.success)
      return { success: false, error: eventResult.error };

    const props = PropertiesService.getDocumentProperties();
    const tasks = JSON.parse(props.getProperty("tasks") || "{}");

    const taskKey = `${cellInfo.spreadsheetId}_${cellInfo.sheetName}_${cellInfo.cellRef}`;
    tasks[taskKey] = {
      message,
      dueDate,
      repeatType,
      notification,
      eventId: eventResult.eventId,
      cellInfo,
      createdAt: new Date().toISOString(),
    };

    props.setProperty("tasks", JSON.stringify(tasks));

    return { success: true, taskKey, eventId: eventResult.eventId };
  } catch (error) {
    console.error("Error creating task:", error);
    return { success: false, error: error.toString() };
  }
}

function listTasks() {
  const props = PropertiesService.getDocumentProperties();
  const tasks = JSON.parse(props.getProperty("tasks") || "{}");

  let html = "<h3>Existing Tasks</h3>";

  if (Object.keys(tasks).length === 0) {
    html += "<p>No tasks found.</p>";
  } else {
    html += "<ul style='list-style:none;padding:0;'>";
    for (const key in tasks) {
      const t = tasks[key];
      const cellDisplay = `${t.cellInfo.sheetName}!${t.cellInfo.cellRef}`;
      const dueDisplay = new Date(t.dueDate).toLocaleDateString();
      let repeatDisplay = "";
      if (t.repeatType !== "none") {
        if (
          typeof t.repeatType === "object" &&
          t.repeatType.type === "custom"
        ) {
          repeatDisplay = ` (Every ${t.repeatType.interval} ${t.repeatType.frequency}`;
          if (t.repeatType.end.type === "after") {
            repeatDisplay += `, ${t.repeatType.end.count} times)`;
          } else if (t.repeatType.end.type === "on") {
            repeatDisplay += `, until ${t.repeatType.end.date})`;
          } else {
            repeatDisplay += ")";
          }
        } else {
          repeatDisplay = ` (${t.repeatType})`;
        }
      }
      const notifDisplay = t.notification
        ? `<br><small>Notify: ${t.notification.value} ${t.notification.unit} before</small>`
        : "";

      html += `<li style="margin-bottom:10px;padding:8px;border:1px solid #ddd;border-radius:4px;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div style="flex:1;">
            <strong>${cellDisplay}</strong><br>
            ${t.message}<br>
            <small>Due: ${dueDisplay}${repeatDisplay}</small>
            ${notifDisplay}
          </div>
          <button onclick="google.script.run.deleteTaskFromList('${key}')" style="background:#dc3545;color:white;border:none;padding:4px 8px;border-radius:3px;cursor:pointer;font-size:11px;">Delete</button>
        </div>
      </li>`;
    }
    html += "</ul>";
  }

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Tasks");
}

function deleteTaskFromList(taskKey) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const tasks = JSON.parse(props.getProperty("tasks") || "{}");

    if (tasks[taskKey]) {
      // Delete from Google Calendar
      if (tasks[taskKey].eventId) {
        const deleteResult = deleteEvent(tasks[taskKey].eventId);
        if (!deleteResult.success) {
          console.error("Failed to delete calendar event:", deleteResult.error);
        }
      }

      // Delete from storage
      delete tasks[taskKey];
      props.setProperty("tasks", JSON.stringify(tasks));

      showSuccessMessage("Task deleted!");
      setTimeout(() => {
        listTasks();
      }, 1500);
      return { success: true };
    }
    return { success: false, error: "Task not found" };
  } catch (error) {
    console.error("Error deleting task:", error);
    return { success: false, error: error.toString() };
  }
}

// HTML Form

function getReminderFormHtml() {
  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
      .container { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
      .form-group { margin-bottom: 20px; }
      label { display: block; margin-bottom: 5px; font-weight: 600; color: #333; }
      input[type="text"], input[type="datetime-local"], input[type="date"], select { width: 100%; padding: 10px; border: 2px solid #e0e0e0; border-radius: 4px; font-size: 14px; box-sizing: border-box; transition: border-color 0.3s; }
      input[type="text"]:focus, input[type="datetime-local"]:focus, input[type="date"]:focus, select:focus { outline: none; border-color: #1976d2; }
      input[type="checkbox"] { margin-right: 8px; transform: scale(1.2); }
      .button-group { display: flex; gap: 10px; margin-top: 20px; }
      button { flex: 1; padding: 12px; border: none; border-radius: 4px; font-size: 14px; font-weight: 600; cursor: pointer; transition: background-color 0.3s; }
      .btn-primary { background-color: #1976d2; color: white; }
      .btn-primary:hover { background-color: #1565c0; }
      .btn-secondary { background-color: #757575; color: white; }
      .btn-secondary:hover { background-color: #616161; }
      .error { color: #d32f2f; font-size: 12px; margin-top: 5px; display: none; }
      .success { color: #388e3c; font-size: 12px; margin-top: 5px; display: none; }
      .loading { display: none; text-align: center; padding: 20px; }
      .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1976d2; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin: 0 auto 10px; }
      @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
      .help-text { font-size: 12px; color: #666; margin-top: 5px; }
      .cell-info { background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin-bottom: 15px; border-left: 4px solid #1976d2; }
      .tabs { display: flex; gap: 0px; margin-bottom: 20px; border-bottom: 2px solid #e0e0e0; }
      .tab { flex: 1; padding: 12px; background: #f5f5f5; border: none; border-bottom: 3px solid transparent; cursor: pointer; font-size: 14px; font-weight: 600; color: #666; transition: all 0.3s; }
      .tab:hover { background: #eeeeee; color: #333; }
      .tab.active { background: white; color: #1976d2; border-bottom-color: #1976d2; }
      .tab-content { display: none; }
      .tab-content.active { display: block; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="tabs">
        <button class="tab active" onclick="switchTab('events')">Events</button>
        <button class="tab" onclick="switchTab('tasks')">Tasks</button>
      </div>

      <!-- Events Tab -->
      <div id="eventsTab" class="tab-content active">
        <h3>Set Event</h3>
        <div class="form-group cell-info"><strong>Selected Cell:</strong> <span id="cellDisplay">Loading...</span><input type="hidden" id="cellInfo" /></div>
        <form id="reminderForm">
        <div class="form-group">
          <label for="msg">Event Message *</label>
          <input type="text" id="msg" required />
        </div>

        <div class="form-group">
          <label><input type="checkbox" id="allDay" onchange="toggleDateTimeInput()" /> All Day Event</label>
          <div class="help-text">Check if this is an all-day event</div>
        </div>

        <div class="form-group">
          <label for="due">Due Date & Time *</label>
          <input type="datetime-local" id="due" required />
          <input type="date" id="dueDate" style="display:none;" />
        </div>
        <div class="form-group">
          <label for="repeat">Repeat</label>
          <select id="repeat" onchange="toggleCustomRepeat('event')">
            <option value="none">No Repeat</option>
            <option value="daily">Daily</option>
            <option value="weekly">Weekly</option>
            <option value="monthly">Monthly</option>
            <option value="yearly">Yearly</option>
            <option value="custom">Custom</option>
          </select>
        </div>
        <div id="customRepeatEvent" class="form-group" style="display:none; background:#f9f9f9; padding:15px; border-radius:4px; margin-top:-10px;">
          <label style="margin-bottom:10px;">Custom Repeat Pattern</label>
          <div style="display:flex; gap:10px; margin-bottom:10px; align-items:center;">
            <span>Every</span>
            <input type="number" id="repeatIntervalEvent" min="1" value="1" style="width:60px; padding:5px;" />
            <select id="repeatFrequencyEvent" style="flex:1;">
              <option value="daily">day(s)</option>
              <option value="weekly">week(s)</option>
              <option value="monthly">month(s)</option>
              <option value="yearly">year(s)</option>
            </select>
          </div>
          <div>
            <label style="margin-bottom:5px;">Ends</label>
            <select id="repeatEndEvent" onchange="toggleRepeatEnd('event')" style="width:100%;">
              <option value="never">Never</option>
              <option value="after">After</option>
              <option value="on">On</option>
            </select>
            <div id="repeatEndAfterEvent" style="display:none; margin-top:10px;">
              <input type="number" id="repeatEndCountEvent" min="1" value="10" placeholder="10" style="width:100px; padding:5px;" />
              <span>occurrence(s)</span>
            </div>
            <div id="repeatEndOnEvent" style="display:none; margin-top:10px;">
              <input type="date" id="repeatEndDateEvent" style="width:100%; padding:5px;" />
            </div>
          </div>
        </div>
        <div class="form-group">
          <label>Notification</label>
          <div style="display:flex; gap:5px;">
            <input type="number" id="notifValue" min="1" placeholder="10" style="flex:1;" />
            <select id="notifUnit" style="flex:1;">
              <option value="">None</option>
              <option value="minutes">Minutes</option>
              <option value="hours">Hours</option>
              <option value="days">Days</option>
              <option value="weeks">Weeks</option>
            </select>
          </div>
          <small>Set how long before the event you want a popup reminder.</small>
        </div>
        <div class="button-group">
          <button type="button" class="btn-secondary" onclick="cancelForm()">Cancel</button>
          <button type="submit" class="btn-primary">Add Event</button>
        </div>
      </form>
      <div id="successMessage" style="display:none; background-color:#d4edda; border:1px solid #c3e6cb; color:#155724; padding:10px; border-radius:4px; margin-top:10px;">
        ✓ Event has been added successfully!
      </div>
      </div>

      <!-- Tasks Tab -->
      <div id="tasksTab" class="tab-content">
        <h3>Set Task</h3>
        <div class="form-group cell-info"><strong>Selected Cell:</strong> <span id="cellDisplayTask">Loading...</span><input type="hidden" id="cellInfoTask" /></div>
        <form id="taskForm">
          <div class="form-group">
            <label for="taskMsg">Task Message *</label>
            <input type="text" id="taskMsg" required />
          </div>

          <div class="form-group">
            <label for="taskDue">Due Date *</label>
            <input type="date" id="taskDue" required />
          </div>

          <div class="form-group">
            <label for="taskRepeat">Repeat</label>
            <select id="taskRepeat" onchange="toggleCustomRepeat('task')">
              <option value="none">No Repeat</option>
              <option value="daily">Daily</option>
              <option value="weekly">Weekly</option>
              <option value="monthly">Monthly</option>
              <option value="yearly">Yearly</option>
              <option value="custom">Custom</option>
            </select>
          </div>
          <div id="customRepeatTask" class="form-group" style="display:none; background:#f9f9f9; padding:15px; border-radius:4px; margin-top:-10px;">
            <label style="margin-bottom:10px;">Custom Repeat Pattern</label>
            <div style="display:flex; gap:10px; margin-bottom:10px; align-items:center;">
              <span>Every</span>
              <input type="number" id="repeatIntervalTask" min="1" value="1" style="width:60px; padding:5px;" />
              <select id="repeatFrequencyTask" style="flex:1;">
                <option value="daily">day(s)</option>
                <option value="weekly">week(s)</option>
                <option value="monthly">month(s)</option>
                <option value="yearly">year(s)</option>
              </select>
            </div>
            <div>
              <label style="margin-bottom:5px;">Ends</label>
              <select id="repeatEndTask" onchange="toggleRepeatEnd('task')" style="width:100%;">
                <option value="never">Never</option>
                <option value="after">After</option>
                <option value="on">On</option>
              </select>
              <div id="repeatEndAfterTask" style="display:none; margin-top:10px;">
                <input type="number" id="repeatEndCountTask" min="1" value="10" placeholder="10" style="width:100px; padding:5px;" />
                <span>occurrence(s)</span>
              </div>
              <div id="repeatEndOnTask" style="display:none; margin-top:10px;">
                <input type="date" id="repeatEndDateTask" style="width:100%; padding:5px;" />
              </div>
            </div>
          </div>

          <div class="form-group">
            <label>Notification</label>
            <div style="display:flex; gap:5px;">
              <input type="number" id="taskNotifValue" min="1" placeholder="10" style="flex:1;" />
              <select id="taskNotifUnit" style="flex:1;">
                <option value="">None</option>
                <option value="minutes">Minutes</option>
                <option value="hours">Hours</option>
                <option value="days">Days</option>
                <option value="weeks">Weeks</option>
              </select>
            </div>
            <small>Set how long before the task you want a popup reminder.</small>
          </div>

          <div class="button-group">
            <button type="button" class="btn-secondary" onclick="cancelForm()">Cancel</button>
            <button type="submit" class="btn-primary">Add Task</button>
          </div>
        </form>
        <div id="taskSuccessMessage" style="display:none; background-color:#d4edda; border:1px solid #c3e6cb; color:#155724; padding:10px; border-radius:4px; margin-top:10px;">
          ✓ Task has been added successfully!
        </div>
      </div>
    </div>
    <script>
      document.addEventListener("DOMContentLoaded", function(){ initializeForm(); setInterval(refreshSelectedCell, 2000); });
      
      function initializeForm(){
        google.script.run.withSuccessHandler(handleCellResponse).getActiveCellA1();
        const now=new Date();
        const oneHourLater=new Date(now.getTime()+60*60*1000);
        document.getElementById("due").value=new Date(oneHourLater.getTime()-oneHourLater.getTimezoneOffset()*60000).toISOString().slice(0,16);
      }

      function handleCellResponse(cellInfo){
        if(cellInfo){ 
          document.getElementById("cellDisplay").textContent=cellInfo.sheetName+"!"+cellInfo.cellRef; 
          document.getElementById("cellInfo").value=JSON.stringify(cellInfo); 
          google.script.run.withSuccessHandler(val=>{
            const msg=document.getElementById("msg");
            if(!msg.value || msg.value===msg.getAttribute("data-lastCellValue")) {
              msg.value=val;
              msg.setAttribute("data-lastCellValue",val);
            }
          }).getCellValue(cellInfo.cellRef,cellInfo.sheetName,cellInfo.spreadsheetId); 
        }
      }

      function refreshSelectedCell(){
        const activeTab = document.querySelector('.tab-content.active');
        if(activeTab && activeTab.id === 'eventsTab') {
          google.script.run.withSuccessHandler(handleCellResponse).getActiveCellA1();
        } else if(activeTab && activeTab.id === 'tasksTab') {
          google.script.run.withSuccessHandler(handleTaskCellResponse).getActiveCellA1();
        }
      }

      function switchTab(tabName) {
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
        
        if(tabName === 'events') {
          document.getElementById('eventsTab').classList.add('active');
          document.querySelectorAll('.tab')[0].classList.add('active');
        } else {
          document.getElementById('tasksTab').classList.add('active');
          document.querySelectorAll('.tab')[1].classList.add('active');
          if(!document.getElementById('cellDisplayTask').hasAttribute('data-initialized')) {
            initializeTaskForm();
            document.getElementById('cellDisplayTask').setAttribute('data-initialized', 'true');
          }
        }
      }

      function initializeTaskForm(){
        google.script.run.withSuccessHandler(handleTaskCellResponse).getActiveCellA1();
        const now = new Date();
        const tomorrow = new Date(now);
        tomorrow.setDate(tomorrow.getDate() + 1);
        document.getElementById("taskDue").value = tomorrow.toISOString().slice(0, 10);
      }

      function handleTaskCellResponse(cellInfo){
        if(cellInfo){ 
          document.getElementById("cellDisplayTask").textContent=cellInfo.sheetName+"!"+cellInfo.cellRef; 
          document.getElementById("cellInfoTask").value=JSON.stringify(cellInfo); 
          google.script.run.withSuccessHandler(val=>{
            const msg=document.getElementById("taskMsg");
            if(!msg.value || msg.value===msg.getAttribute("data-lastCellValue")) {
              msg.value=val;
              msg.setAttribute("data-lastCellValue",val);
            }
          }).getCellValue(cellInfo.cellRef,cellInfo.sheetName,cellInfo.spreadsheetId); 
        }
      }

      function toggleDateTimeInput(){
        const allDay=document.getElementById("allDay").checked;
        document.getElementById("due").style.display=allDay?"none":"block";
        document.getElementById("dueDate").style.display=allDay?"block":"none";
      }

      function toggleCustomRepeat(type) {
        const prefix = type === "event" ? "Event" : "Task";
        const repeatSelect = document.getElementById(type === "event" ? "repeat" : "taskRepeat");
        const customDiv = document.getElementById("customRepeat" + prefix);
        
        if (repeatSelect.value === "custom") {
          customDiv.style.display = "block";
        } else {
          customDiv.style.display = "none";
        }
      }

      function toggleRepeatEnd(type) {
        const prefix = type === "event" ? "Event" : "Task";
        const endSelect = document.getElementById("repeatEnd" + prefix);
        const afterDiv = document.getElementById("repeatEndAfter" + prefix);
        const onDiv = document.getElementById("repeatEndOn" + prefix);
        
        afterDiv.style.display = endSelect.value === "after" ? "block" : "none";
        onDiv.style.display = endSelect.value === "on" ? "block" : "none";
      }

      document.getElementById("reminderForm").addEventListener("submit",function(e){
        e.preventDefault();
        const msg=document.getElementById("msg").value.trim();
        const allDay=document.getElementById("allDay").checked;
        let repeat=document.getElementById("repeat").value;
        const cellInfo=JSON.parse(document.getElementById("cellInfo").value);
        let due=allDay?document.getElementById("dueDate").value:document.getElementById("due").value;
        
        // Handle custom repeat
        if(repeat === "custom") {
          const interval=parseInt(document.getElementById("repeatIntervalEvent").value)||1;
          const frequency=document.getElementById("repeatFrequencyEvent").value;
          const endType=document.getElementById("repeatEndEvent").value;
          const repeatEnd={ type:endType };
          
          if(endType === "after") {
            repeatEnd.count=parseInt(document.getElementById("repeatEndCountEvent").value)||10;
          } else if(endType === "on") {
            repeatEnd.date=document.getElementById("repeatEndDateEvent").value;
          }
          
          repeat={
            type:"custom",
            frequency:frequency,
            interval:interval,
            end:repeatEnd
          };
        }
        
        let notification=null;
        const notifValue=document.getElementById("notifValue").value;
        const notifUnit=document.getElementById("notifUnit").value;
        if(notifValue && notifUnit){ notification={ value:notifValue, unit:notifUnit }; }
        google.script.run.withSuccessHandler(()=>{
          showSuccessMessage("Event added!");
          setTimeout(() => {
            google.script.run.showReminderSidebar();
          }, 1500);
        }).createReminder(cellInfo,due,msg,allDay,repeat,notification);
      });

      function showReminderSidebarAgain(){
        google.script.run.showReminderSidebar();
      }

      function showSuccessMessage(message) {
        const successBox = document.getElementById("successMessage");
        successBox.style.display = "block";
        setTimeout(() => {
          successBox.style.display = "none";
        }, 3000);
      }

      document.getElementById("taskForm").addEventListener("submit",function(e){
        e.preventDefault();
        const msg=document.getElementById("taskMsg").value.trim();
        const due=document.getElementById("taskDue").value;
        let repeat=document.getElementById("taskRepeat").value;
        const cellInfo=JSON.parse(document.getElementById("cellInfoTask").value);
        
        // Handle custom repeat
        if(repeat === "custom") {
          const interval=parseInt(document.getElementById("repeatIntervalTask").value)||1;
          const frequency=document.getElementById("repeatFrequencyTask").value;
          const endType=document.getElementById("repeatEndTask").value;
          const repeatEnd={ type:endType };
          
          if(endType === "after") {
            repeatEnd.count=parseInt(document.getElementById("repeatEndCountTask").value)||10;
          } else if(endType === "on") {
            repeatEnd.date=document.getElementById("repeatEndDateTask").value;
          }
          
          repeat={
            type:"custom",
            frequency:frequency,
            interval:interval,
            end:repeatEnd
          };
        }
        
        let notification=null;
        const notifValue=document.getElementById("taskNotifValue").value;
        const notifUnit=document.getElementById("taskNotifUnit").value;
        if(notifValue && notifUnit){ notification={ value:notifValue, unit:notifUnit }; }

        google.script.run.withSuccessHandler(()=>{
          showTaskSuccessMessage("Task added!");
          setTimeout(() => {
            google.script.run.showReminderSidebar();
          }, 1500);
        }).createTask(cellInfo,due,msg,repeat,notification);
      });

      function showTaskSuccessMessage(message) {
        const successBox = document.getElementById("taskSuccessMessage");
        successBox.style.display = "block";
        setTimeout(() => {
          successBox.style.display = "none";
        }, 3000);
      }

      function cancelForm(){ google.script.host.close(); }
    </script>
  </body>
</html>`;
}
