function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Cell Reminders")
    .addItem("Add Reminder", "showReminderSidebar")
    .addItem("View Reminders", "listReminders")
    .addItem("Clear All", "clearAllReminders")
    .addToUi();
}

function getActiveCellA1() {
  const range = SpreadsheetApp.getActiveRange();
  return range ? range.getA1Notation() : "";
}

function showReminderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ReminderForm")
    .setTitle("Cell Reminders")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function createReminder(cellRef, dueDate, message) {
  try {
    if (!cellRef || !validateCellReference(cellRef)) {
      return { success: false, error: "Invalid cell reference" };
    }

    if (!message || message.trim().length === 0) {
      return { success: false, error: "Message is required" };
    }

    const dateValidation = validateDate(dueDate);
    if (!dateValidation.isValid) {
      return { success: false, error: dateValidation.error };
    }

    const startTime = dateValidation.date;
    const endTime = new Date(startTime.getTime() + 30 * 60 * 1000);

    const eventResult = createCalendarEvent(
      `Reminder: ${message}`,
      startTime,
      endTime,
      `Cell ${cellRef} reminder: ${message}`
    );

    if (!eventResult.success) {
      return { success: false, error: eventResult.error };
    }

    const reminders = getAllReminders();
    const reminderId = generateReminderId();

    reminders[cellRef] = {
      id: reminderId,
      message: message.trim(),
      dueDate: dueDate,
      eventId: eventResult.eventId,
      createdAt: new Date().toISOString(),
      cellValue: getCellValue(cellRef),
    };

    saveReminders(reminders);
    showToast(`Reminder created for cell ${cellRef}`, "Success");

    return { success: true };
  } catch (error) {
    console.error("Error creating reminder:", error);
    return {
      success: false,
      error: "Failed to create reminder: " + error.toString(),
    };
  }
}

function listReminders() {
  try {
    const reminders = getAllReminders();

    if (Object.keys(reminders).length === 0) {
      const html = HtmlService.createHtmlOutput(
        `
        <div style="padding: 20px; text-align: center;">
          <h3>No Reminders</h3>
          <p>You don't have any reminders set yet.</p>
          <p>Select a cell and click "Add Reminder" to get started!</p>
        </div>
      `
      )
        .setWidth(400)
        .setHeight(200);

      SpreadsheetApp.getUi().showModalDialog(html, "Reminders");
      return;
    }

    let html = `
      <div style="padding: 20px;">
        <h3>Your Reminders</h3>
        <div style="max-height: 400px; overflow-y: auto;">
    `;

    for (const cellRef in reminders) {
      const reminder = reminders[cellRef];
      const formattedDate = formatDate(reminder.dueDate);
      const isOverdue = new Date(reminder.dueDate) < new Date();

      html += `
        <div style="border: 1px solid #ddd; margin: 10px 0; padding: 15px; border-radius: 5px; ${
          isOverdue ? "background-color: #ffebee;" : ""
        }">
          <div style="font-weight: bold; color: #1976d2;">Cell ${cellRef}</div>
          <div style="margin: 5px 0;">${reminder.message}</div>
          <div style="color: ${
            isOverdue ? "#d32f2f" : "#666"
          }; font-size: 0.9em;">
            Due: ${formattedDate} ${isOverdue ? "OVERDUE" : ""}
          </div>
          <div style="margin-top: 10px;">
            <button onclick="deleteReminder('${cellRef}')" style="background: #f44336; color: white; border: none; padding: 5px 10px; border-radius: 3px; cursor: pointer;">
              Delete
            </button>
          </div>
        </div>
      `;
    }

    html += `
        </div>
        <div style="margin-top: 20px; text-align: center;">
          <button onclick="google.script.host.close()" style="background: #1976d2; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer;">
            Close
          </button>
        </div>
      </div>
    `;

    const output = HtmlService.createHtmlOutput(html)
      .setWidth(500)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(output, "Reminders");
  } catch (error) {
    console.error("Error listing reminders:", error);
    showToast("Error loading reminders", "Error");
  }
}

function deleteReminder(cellRef) {
  try {
    const reminders = getAllReminders();

    if (!reminders[cellRef]) {
      return { success: false, error: "Reminder not found" };
    }

    const reminder = reminders[cellRef];

    if (reminder.eventId) {
      deleteCalendarEvent(reminder.eventId);
    }

    delete reminders[cellRef];
    saveReminders(reminders);

    showToast(`Reminder for cell ${cellRef} deleted`, "Success");
    return { success: true };
  } catch (error) {
    console.error("Error deleting reminder:", error);
    return {
      success: false,
      error: "Failed to delete reminder: " + error.toString(),
    };
  }
}

function clearAllReminders() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Clear All Reminders",
      "Are you sure you want to delete all reminders? This action cannot be undone.",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      const reminders = getAllReminders();

      for (const cellRef in reminders) {
        const reminder = reminders[cellRef];
        if (reminder.eventId) {
          deleteCalendarEvent(reminder.eventId);
        }
      }

      saveReminders({});
      showToast("All reminders cleared", "Success");
    }
  } catch (error) {
    console.error("Error clearing reminders:", error);
    showToast("Error clearing reminders", "Error");
  }
}
