/**
 * Cell Reminders - Developer Script Version
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
    .addItem("Add Cell Event", "showReminderSidebar")
    .addItem("View Cell Events", "listReminders")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();
}

// ============================================================================
// CORE FUNCTIONS
// ============================================================================

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
      <li>Choose a color for your event</li>
      <li>Click "Add Event"</li>
    </ol>
    <p><strong>Features:</strong></p>
    <ul>
      <li>Creates events in Google Calendar</li>
      <li>Works with any Google Sheet</li>
      <li>Supports all-day and timed events</li>
      <li>Repeat options: daily, weekly, monthly, yearly</li>
      <li>Custom notifications</li>
      <li>Event colors</li>
      <li>Delete events from the list</li>
    </ul>
  `;

  const output = HtmlService.createHtmlOutput(helpHtml)
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Help");
}

// ============================================================================
// EVENT MANAGEMENT FUNCTIONS
// ============================================================================

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

function getCalendarColors() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    return {
      1: "#a4bdfc", // Lavender
      2: "#7ae7bf", // Sage
      3: "#dbadff", // Grape
      4: "#ff887c", // Flamingo
      5: "#fbd75b", // Banana
      6: "#ffb878", // Tangerine
      7: "#46d6db", // Peacock
      8: "#e1e1e1", // Graphite
      9: "#5484ed", // Blueberry
      10: "#51b749", // Basil
      11: "#dc2127", // Tomato
    };
  } catch (error) {
    console.error("Error getting calendar colors:", error);
    return {};
  }
}

function showSuccessMessage(message) {
  const html = `
    <div style="text-align: center; padding: 20px;">
      <div style="color: #388e3c; font-size: 18px; font-weight: bold; margin-bottom: 10px;">
        ✓ ${message}
      </div>
      <div style="color: #666; font-size: 14px;">
        Please wait while we update the list...
      </div>
    </div>
  `;

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(output, "Success");
}

// ============================================================================
// GOOGLE CALENDAR API FUNCTIONS
// ============================================================================

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

// ============================================================================
// CELL HELPER
// ============================================================================

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

// ============================================================================
// HTML FORM WITH LIVE CELL UPDATES
// ============================================================================

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
    </style>
  </head>
  <body>
    <div class="container">
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
          <select id="repeat">
            <option value="none">No Repeat</option>
            <option value="daily">Daily</option>
            <option value="weekly">Weekly</option>
            <option value="monthly">Monthly</option>
            <option value="yearly">Yearly</option>
          </select>
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
        google.script.run.withSuccessHandler(handleCellResponse).getActiveCellA1();
      }

      function toggleDateTimeInput(){
        const allDay=document.getElementById("allDay").checked;
        document.getElementById("due").style.display=allDay?"none":"block";
        document.getElementById("dueDate").style.display=allDay?"block":"none";
      }

      document.getElementById("reminderForm").addEventListener("submit",function(e){
        e.preventDefault();
        const msg=document.getElementById("msg").value.trim();
        const allDay=document.getElementById("allDay").checked;
        const repeat=document.getElementById("repeat").value;
        const cellInfo=JSON.parse(document.getElementById("cellInfo").value);
        let due=allDay?document.getElementById("dueDate").value:document.getElementById("due").value;
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

      function cancelForm(){ google.script.host.close(); }
    </script>
  </body>
</html>`;
}
