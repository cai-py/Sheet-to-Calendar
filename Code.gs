function parseDate(dateString, timeString) {
  Logger.log("Original dateString: " + dateString + ", timeString: " + timeString);

  if (dateString instanceof Date) {
    dateString = Utilities.formatDate(dateString, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  Logger.log("Formatted dateString: " + dateString);

  if (timeString) {
    // Ensure timeString is a string in "HH:mm:ss" format
    if (timeString instanceof Date) {
      timeString = Utilities.formatDate(timeString, Session.getScriptTimeZone(), "HH:mm:ss");
    }
    var dateTimeString = `${dateString}T${timeString}`;
    //ISO 8601 format: `YYYY-MM-DDTHH:MM:SSZ` or `YYYY-MM-DDTHH:MM:SS±HH:MM`
    var parsedDate = new Date(dateTimeString);
    Logger.log("Parsed dateTimeString: " + dateTimeString + ", Parsed Date: " + parsedDate);
    return isNaN(parsedDate.getTime()) ? new Date(dateString) : parsedDate;
  } else {
    var parsedDate = new Date(dateString);
    Logger.log("Parsed Date: " + parsedDate);
    return parsedDate;
  }
}

function syncSheetToCalendar() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'c77d39b85831a771d41e2ffc594f67263ab7c2d0a76369e3b4d9c77fabc98ad4@group.calendar.google.com'; 

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
    var idRange = sheet.getRange(2, 8, lastRow - 1, 1);
    var ids = idRange.getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var eventTitle = row[0];
      var eventStartDate = row[1];
      var eventEndDate = row[2];
      var eventStartTime = row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "HH:mm:ss") : null;
      var eventEndTime = row[4] ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "HH:mm:ss") : null;
      var eventDescription = row[5];
      var eventLocation = row[6];
      var eventId = ids[i][0];

      Logger.log("Processing row " + (i + 2) + ": " + eventTitle + ", " + eventStartDate + ", " + eventStartTime + ", " + eventEndDate + ", " + eventEndTime);

      var startDateTime, endDateTime;

      if (eventStartTime && eventEndTime) {
        startDateTime = parseDate(eventStartDate, eventStartTime);
        endDateTime = parseDate(eventEndDate, eventEndTime);
      } else {
        startDateTime = new Date(eventStartDate);
        endDateTime = new Date(eventEndDate);
        endDateTime.setDate(endDateTime.getDate() + 1); // End date needs to be the next day for all-day events
      }

      Logger.log("Parsed startDateTime: " + startDateTime);
      Logger.log("Parsed endDateTime: " + endDateTime);

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
          event = calendar.createAllDayEvent(eventTitle, startDateTime, {
            description: eventDescription,
            location: eventLocation
          });
        }
        ids[i][0] = event.getId();
        Logger.log("Event created: " + eventTitle);
      }
    }

    idRange.setValues(ids);
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

function syncCalendarToSheet() {
  try {
    var sheetName = 'Fall 2024';  
    var calendarId = 'your-calendar@group.calendar.google.com'; 

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

    var events = calendar.getEvents(new Date('January 1, 2020'), new Date('December 31, 2024'));
    var data = [];

    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var startTime = event.getStartTime();
      var endTime = event.getEndTime();
      var eventId = event.getId();
      var eventTitle = event.getTitle();
      var eventDescription = event.getDescription();
      var eventLocation = event.getLocation();

      var eventStartDate = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var eventEndDate = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var startTimeStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm:ss");
      var endTimeStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss");

      data.push([eventTitle, eventStartDate, eventEndDate, startTimeStr, endTimeStr, eventDescription, eventLocation, eventId]);
    }

    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }

    Logger.log("Sheet updated from calendar.");
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

// Run Manually to sync Calendar and Sheet
function manualFullSync() {
  syncCalendarToSheet();
  syncSheetToCalendar();
}

