# Cell Reminders - User Guide

## Overview

Cell Reminders allows you to create Google Calendar events directly from any cell in your Google Sheets. Events are linked to specific cells and can be set as all-day or timed reminders with repeat options.

## Getting Started

### Installation (Developer Script - Immediate Use)

1. **Create New Apps Script Project**
   - Go to [script.google.com](https://script.google.com)
   - Click "New Project"
   - Name it "Cell Reminders"

2. **Add the Code**
   - Delete the default code in Code.gs
   - Copy and paste the entire contents from [`script/CellRemindersScript.js`](scripts\CellRemindersScript.js)
   - Replace the `appsscript.json` content with the version from `script/appsscript.json`

3. **Enable Calendar API**
   - Click the "Services" button (+ icon) in the left sidebar
   - Search for "Calendar API"
   - Click "Add"

4. **Save and Authorize**
   - Save the project (Ctrl+S)
   - Click "Run" to authorize the script
   - Grant all requested permissions

5. **Test in Google Sheets**
   - Open any Google Sheet
   - Refresh the page
   - You should see "Cell Reminders" in the menu bar

## How to Use

### Creating a Reminder

1. **Select a Cell**
   - Click on any cell in your Google Sheet
   - The cell content will be used as the default event message

2. **Open Reminder Form**
   - Go to "Cell Reminders" > "Add Cell Reminder"
   - A sidebar will open on the right

3. **Fill Out the Form**
   - **Message**: Edit or enter your event description
   - **All Day Event**: Check if this is an all-day event (no specific time)
   - **Due Date & Time**:
     - For timed events: Select date and time
     - For all-day events: Select date only
   - **Repeat**: Choose if the event should repeat (daily, weekly, monthly, yearly)

4. **Create the Event**
   - Click "Add Reminder"
   - The event will be created in Google Calendar
   - A success message will appear

### Viewing Your Reminders

1. **Open Reminders List**
   - Go to "Cell Reminders" > "View Cell Reminders"
   - A dialog will show all your existing reminders

2. **Reminder Information Displayed**
   - Cell location (Sheet name and cell reference)
   - Event message
   - Due date/time
   - Repeat setting (if any)

## Troubleshooting

### Common Issues

#### "No cell selected" Error

**Problem**: Trying to create a reminder without selecting a cell
**Solution**: Click on a cell first, then open the reminder form

#### "Authorization required" Error

**Problem**: Script needs permission to access Google services
**Solution**:

1. Run the script from the Apps Script editor
2. Grant all requested permissions
3. Try again in Google Sheets

#### Events not appearing in Google Calendar

**Problem**: Calendar API might not be enabled
**Solution**:

1. Go back to the Apps Script editor
2. Check that Calendar API is enabled in Services
3. Re-save and try again

#### Form won't open

**Problem**: Script might not be properly installed
**Solution**:

1. Refresh the Google Sheet
2. Check that the menu appears
3. If not, re-install the script

### Getting Help

1. **Check the Browser Console**
   - Press F12 in your browser
   - Look for error messages in the Console tab

2. **View Apps Script Logs**
   - Go to the Apps Script editor
   - Click "Executions" to see recent runs and errors

3. **Test with Simple Cases**
   - Try creating an event with minimal information
   - Use a simple cell with basic text

## Privacy and Security

### Data Handling

- Event data is stored in Google Calendar (your Google account)
- Reminder metadata is stored in the Google Sheet's properties
- No data is sent to external servers
- All operations use Google's secure APIs

### Permissions Required

- **Google Sheets**: To read cell content and show the add-on menu
- **Google Calendar**: To create and manage events
- **Script Container UI**: To show the sidebar and dialogs

### Data Retention

- Events remain in Google Calendar until you delete them
- Reminder metadata stays with the Google Sheet
- Copying or sharing sheets includes the reminder data
