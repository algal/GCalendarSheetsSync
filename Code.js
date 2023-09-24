function updateEventsFromSheetNew(sheetName, calendarName, startDate, endDate) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
        Logger.log('No sheet found with the name: ' + sheetName);
        return;
    }

    // Get the calendar by its name
    var calendars = Calendar.CalendarList.list();
    var calendarId = null;
    for (var i = 0; i < calendars.items.length; i++) {
        if (calendars.items[i].summary === calendarName) {
            calendarId = calendars.items[i].id;
            break;
        }
    }

    if (!calendarId) {
        Logger.log('No calendar found with the name: ' + calendarName);
        return;
    }

    // Fetch data from the sheet
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // Map column headers to their index
    var colMap = {};
    headers.forEach(function(header, index) {
        colMap[header] = index;
    });

    // Verify presence of EventId
    if (colMap["EventId"] === undefined) {
        Logger.log("The column 'EventId' is missing in the sheet.");
        return;
    }

    // Loop through rows and process events based on ACTION
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var eventId = row[colMap["EventId"]];
        var action = colMap["ACTION"] !== undefined ? row[colMap["ACTION"]] : null;

        if (action === "CREATE" && !eventId) {
            var newEvent = {
                summary: colMap["Event Name"] !== undefined ? row[colMap["Event Name"]] : "",
                start: {
                    'dateTime': colMap["Event Start"] !== undefined ? new Date(row[colMap["Event Start"]]).toISOString() : new Date().toISOString()
                },
                end: {
                    'dateTime': colMap["Event End"] !== undefined ? new Date(row[colMap["Event End"]]).toISOString() : new Date().toISOString()
                },
                description: colMap["Description"] !== undefined ? row[colMap["Description"]] : ""
            };
            
            var createdEvent = Calendar.Events.insert(newEvent, calendarId);
            sheet.getRange(i + 1, colMap["EventId"] + 1).setValue(createdEvent.id); // Populate EventId in the sheet

        } else if (action === "DELETE") {
            try {
                Calendar.Events.remove(calendarId, eventId);
            } catch (e) {
                Logger.log('Error deleting the event with EventID: ' + eventId + '. Error: ' + e.toString());
            }
            continue;

        } else {
            // Check if the event exists in the calendar
            var event;
            try {
                event = Calendar.Events.get(calendarId, eventId);
            } catch (e) {
                Logger.log('No event found in the calendar with the EventID: ' + eventId);
                continue;
            }

            // Update the event based on present columns
            if (colMap["Event Name"] !== undefined && row[colMap["Event Name"]]) {
                event.summary = row[colMap["Event Name"]];
            }

            if (colMap["Event Start"] !== undefined && row[colMap["Event Start"]]) {
                event.start = {
                    'dateTime': new Date(row[colMap["Event Start"]]).toISOString()
                };
            }

            if (colMap["Event End"] !== undefined && row[colMap["Event End"]]) {
                event.end = {
                    'dateTime': new Date(row[colMap["Event End"]]).toISOString()
                };
            }

            if (colMap["Description"] !== undefined && row[colMap["Description"]]) {
                event.description = row[colMap["Description"]];
            }

            // Update the event in the calendar
            Calendar.Events.update(event, calendarId, eventId);
        }
    }

    Logger.log('Process completed.');
}

function fancyUpdate() {
  updateEventsFromSheetNew("Events2","Cine Club 2023-2024")
}


function exportEventsToSheetNew(sheetName, calendarName, startDate, endDate) {
    // Access the current spreadsheet and the specified sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    // If the sheet doesn't exist, create it
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
    }
  
    // Set headers
//    sheet.clear();
    sheet.appendRow(['EventId', 'Event Name', 'Event Start', 'Event End', 'Description']);
  
    // Get the calendar by its name
    var calendars = Calendar.CalendarList.list();
    var calendarId = null;
    for (var i = 0; i < calendars.items.length; i++) {
        if (calendars.items[i].summary === calendarName) {
            calendarId = calendars.items[i].id;
            break;
        }
    }

    if (!calendarId) {
        Logger.log('No calendar found with the name: ' + calendarName);
        return;
    }
  
    var options = {
        singleEvents: true,
        orderBy: 'startTime'
    };
  
    if (startDate) {
        options.timeMin = new Date(startDate).toISOString();
    }
  
    if (endDate) {
        options.timeMax = new Date(endDate).toISOString();
    }
    
    var events = Calendar.Events.list(calendarId, options);
    var count = 0;
  
    while (events.items && events.items.length > 0) {
        for (var i = 0; i < events.items.length; i++) {
            var event = events.items[i];
            var eventId = event.id;
            var eventName = event.summary;
            var eventStart = event.start.dateTime || event.start.date;  // Handle all-day events
            var eventEnd = event.end.dateTime || event.end.date;
            var description = event.description || "";
            
            sheet.appendRow([eventId, eventName, eventStart, eventEnd, description]);
            count++;
        }
      
        // If there's a nextPageToken, use it to fetch the next page of results
        if (events.nextPageToken) {
            options.pageToken = events.nextPageToken;
            events = Calendar.Events.list(calendarId, options);
        } else {
            break;  // No more events
        }
    }

    // Log results
    Logger.log('Number of events exported: ' + count);
    Logger.log('Name of the calendar accessed: ' + calendarName);
    Logger.log('Name of the spreadsheet: ' + spreadsheet.getName());
    Logger.log('Name of the sheet: ' + sheetName);
}

function fancyExport() {
  exportEventsToSheet("Events2","Cine Club 2023-2024")
}



function populateSheetWithEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");

  // Specify the calendar by name
  var calendar = CalendarApp.getCalendarsByName("Cine Club 2023-2024")[0];

  // Set headers
  sheet.appendRow(["Event ID", "Event Name", "Event Start", "Event End"]);

  // Get events for the next 2 years
  var now = new Date();
  var twoYearsFromNow = new Date(now);
  twoYearsFromNow.setFullYear(now.getFullYear() + 2);
  var events = calendar.getEvents(now, twoYearsFromNow);

  // Loop through events and append to sheet
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle().startsWith("Cine Club")) {
      var eventId = events[i].getId();
      var eventTitle = events[i].getTitle();
      var eventStart = events[i].getStartTime();
      var eventEnd = events[i].getEndTime();

      sheet.appendRow([eventId, eventTitle, eventStart, eventEnd]);
    }
  }
}

function updateEventsFromSheet() {
  Logger.log("beg: updateEventsFromSheet");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var rows = sheet.getRange("A2:D" + sheet.getLastRow()).getValues();

  // Specify the calendar by name
  var calendar = CalendarApp.getCalendarsByName("Cine Club 2023-2024")[0];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var eventId = row[0];
    var eventTitle = row[1];
    var eventStart = new Date(row[2]);
    var eventEnd = new Date(row[3]);

    // Fetch the event from Google Calendar using ID
    var eventsMatchingTime = calendar.getEvents(eventStart, new Date(eventEnd.getTime() + 1));
    var eventsMatchingId = eventsMatchingTime.filter(function (e) { return e.getId() == eventId });
    if (eventsMatchingId.length == 0) {
      Logger.log("Event not found with ID: " + eventId)
    }
    else if (eventsMatchingId.length > 1) {
      Logger.log("Multiple events found with ID. Not updating: " + eventId)
    }
    else {
      var event = eventsMatchingId[0];
      if (event.getTitle() == eventTitle && event.getStartTime() == eventStart && event.getEndTime() == eventEnd) {
        Logger.log("Not updating event since no value changed for eventID: " + eventId);
      }
      else {
        Logger.log("updating event %s with new title %s: ", eventId, eventTitle)
        Logger.log("old eventTit");
        event.setTitle(eventTitle);
        event.setTime(eventStart, eventEnd);
      }
    }
  }
}

function testLogging() {
  Logger.log("start: testLogging()");
}
