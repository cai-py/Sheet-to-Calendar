function syncSheetToCalendar() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'your-calendar@group.calendar.google.com'; // Change this to your calendar ID

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found");
      return;
    }

    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("Calendar not found");
      return;
    }

    // gather all data from the sheet
    var lastRow = sheet.getLastRow(); // finds the last row that has content
    var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());// sets the range to (row 2, col 1; last row, last column)
    var data = dataRange.getValues();
    var idRange = sheet.getRange(2, 8, lastRow - 1, 1); // Column H for Event ID
    var ids = idRange.getValues();

    // Loop through each row 
    for (var i = 0; i < data.length; i++) {
      // Gather row Data
      var row = data[i];
      var eventTitle = row[0];
      var eventStartDate = row[1];
      var eventEndDate = row[2];
      var eventStartTime = row[3];
      var eventEndTime = row[4];
      var eventDescription = row[5];
      var eventLocation = row[6];
      var eventId = ids[i][0];

      Logger.log("Processing row " + (i + 2) + ": " + eventTitle + ", " + eventStartDate + ", " + eventStartTime + ", " + eventEndTime);

      var startDateTime, endDateTime;

      if (eventStartTime && eventEndTime) {
        var startDateTimeStr = Utilities.formatDate(new Date(eventStartDate + 'T' + eventStartTime), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        var endDateTimeStr = Utilities.formatDate(new Date(eventEndDate + 'T' + eventEndTime), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        startDateTime = new Date(startDateTimeStr);
        endDateTime = new Date(endDateTimeStr);
      } else {
        // Create all-day event
        startDateTime = new Date(eventStartDate);
        endDateTime = new Date(eventEndDate);
        endDateTime.setDate(endDateTime.getDate() + 1); // End date needs to be the next day for all-day events
      }

      if (isNaN(startDateTime.getTime()) || isNaN(endDateTime.getTime())) {
        Logger.log("Invalid date/time: " + startDateTime + " or " + endDateTime);
        continue;
      }

      var event;
      if (eventId) {
        try {
          event = calendar.getEventById(eventId);
          if (event) {
            if (eventStartTime && eventEndTime) {
              event.setTitle(eventTitle);
              event.setTime(startDateTime, endDateTime);
            } else {
              event.setTitle(eventTitle);
              event.setAllDayDate(startDateTime, endDateTime);
            }
            event.setDescription(eventDescription);
            event.setLocation(eventLocation);
            Logger.log("Event updated: " + eventTitle);
          }
        } catch (e) {
          Logger.log("Event not found: " + eventId);
        }
      } else {
        if (eventStartTime && eventEndTime) {
          event = calendar.createEvent(eventTitle, startDateTime, endDateTime, {
            description: eventDescription,
            location: eventLocation
          });
        } else {
          event = calendar.createAllDayEvent(eventTitle, startDateTime, endDateTime, {
            description: eventDescription,
            location: eventLocation
          });
        }
        ids[i][0] = event.getId();
        Logger.log("Event created: " + eventTitle);
      }
    }

    idRange.setValues(ids); // Update the Event ID column in the sheet
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

function syncCalendarToSheet() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'your-calendar@group.calendar.google.com'; // Change this to your calendar ID

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found");
      return;
    }

    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("Calendar not found");
      return;
    }

    var events = calendar.getEvents(new Date('January 1, 2020'), new Date('December 31, 2024')); // Adjust the date range accordingly
    var data = [];
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var startTime = event.getStartTime();
      var endTime = event.getEndTime();
      var eventId = event.getId();
      var eventTitle = event.getTitle();
      var eventDescription = event.getDescription();
      var eventLocation = event.getLocation();
      
      // Combine date and time strings
      var eventStartDate = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var eventEndDate = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var startTimeStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm:ss");
      var endTimeStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss");
      
      data.push([eventTitle, eventStartDate, eventEndDate, startTimeStr, endTimeStr, eventDescription, eventLocation, eventId]);
    }
    
    // Clear existing data (excluding headers)
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    
    // Write new data to the sheet
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }

    Logger.log("Sheet updated from calendar.");
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

function removeDeletedCalendarEventsFromSheet() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'your-calendar@group.calendar.google.com'; // Change this to your calendar ID

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found");
      return;
    }

    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("Calendar not found");
      return;
    }

    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    var data = dataRange.getValues();
    var idRange = sheet.getRange(2, 8, lastRow - 1, 1); // Column H for Event ID
    var ids = idRange.getValues();

    var calendarEventIds = calendar.getEvents(new Date('January 1, 2020'), new Date('December 31, 2024'))
      .map(event => event.getId());

    var rowsToDelete = [];
    for (var i = 0; i < ids.length; i++) {
      var eventId = ids[i][0];
      if (eventId && !calendarEventIds.includes(eventId)) {
        rowsToDelete.push(i + 2); // Adjust for header row
      }
    }

    // Delete rows from bottom to top to avoid index shift issues
    rowsToDelete.reverse().forEach(row => sheet.deleteRow(row));

    Logger.log("Deleted events that no longer exist in calendar.");
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

function deleteAllCalendarEvents() {
  try {
    var calendarId = 'your-calendar@group.calendar.google.com'; // Change this to your calendar ID
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("Calendar not found");
      return;
    }

    var events = calendar.getEvents(new Date('January 1, 2020'), new Date('December 31, 2024')); // Adjust the date range accordingly
    for (var i = 0; i < events.length; i++) {
      events[i].deleteEvent();
      Logger.log("Event deleted: " + events[i].getTitle());
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

function deleteCalendarEventsNotInSheet() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'your-calendar@group.calendar.google.com'; // Change this to your calendar ID

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found");
      return;
    }

    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log("Calendar not found");
      return;
    }

    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    var data = dataRange.getValues();
    var idRange = sheet.getRange(2, 8, lastRow - 1, 1); // Column H for Event ID
    var sheetEventIds = idRange.getValues().flat();

    var calendarEvents = calendar.getEvents(new Date('January 1, 2020'), new Date('December 31, 2024'));
    var calendarEventIds = calendarEvents.map(event => event.getId());

    // Find events in calendar not in the sheet
    var eventsToDelete = calendarEventIds.filter(eventId => !sheetEventIds.includes(eventId));

    eventsToDelete.forEach(eventId => {
      try {
        var event = calendar.getEventById(eventId);
        event.deleteEvent();
        Logger.log("Event deleted: " + eventId);
      } catch (e) {
        Logger.log("Failed to delete event with ID: " + eventId + " - " + e.message);
      }
    });

    Logger.log("Deleted events that are no longer present in the sheet.");
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

// Run Manually to sync Calendar and Sheet
function manualFullSync() {
  removeDeletedEventsFromSheet();
  deleteEventsNotInSheet();
  updateSheetFromCalendar();
  createOrUpdateCalendarEvent();
}
