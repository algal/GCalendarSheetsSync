/**
 * Updates Google Calendar events based on the
 * information from a specified Google Sheet.  
 *
 * The function reads each row of the sheet and performs an action
 * 'CREATE', 'DELETE', or 'UPDATE' on the Calendar, based on the
 * 'ACTION' column in the Sheet.
 *
 * It dynamically recognizes the column headers, allowing for flexible
 * column ordering and optional columns. The column names it
 * recognizes are "EventId", "Event Name", "Event Start", "Event End",
 * "Description", and "Location".
 * 
 * - If ACTION is "CREATE" and EventId is empty, creates a new event and populates EventId in the sheet.
 * 
 * - If ACTION is "DELETE", deletes the event with the corresponding EventId from the calendar.
 * 
 * - For any other value of ACTION, when EventId describes an existing event, it updates the value of any non-empty field.
 * 
 * @function
 * @name updateEventsFromSheet
 * @param {string} sheetName - The name of the sheet/tab within the spreadsheet to read data from.
 * @param {string} calendarName - The name of the Google Calendar to update events in.
 * @param {string} [startDate] - Optional. The start date to process events after (inclusive). Must be in a recognizable date string format, e.g., "YYYY-MM-DD".
 * @param {string} [endDate] - Optional. The end date to process events before (inclusive). Must be in a recognizable date string format, e.g., "YYYY-MM-DD".
 * 
 * @example
 * updateEventsFromSheet("Events", "My Calendar", "2023-01-01", "2023-12-31");
 */
function updateEventsFromSheet(sheetName, calendarName, startDate, endDate) {
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
                description: colMap["Description"] !== undefined ? row[colMap["Description"]] : "",
                location: colMap["Location"] !== undefined ? row[colMap["Location"]] : ""
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

            if (colMap["Location"] !== undefined && row[colMap["Location"]]) {
                event.location = row[colMap["Location"]];
            }

            // Update the event in the calendar
            Calendar.Events.update(event, calendarId, eventId);
        }
    }

    Logger.log('Process completed.');
}

function fancyUpdate() {
  updateEventsFromSheet("Events2","Cine Club 2023-2024")
}


function exportEventsToSheet(calendarName, startDate, endDate) {
    // Access the current spreadsheet and the specified sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(calendarName);
    
    // If the sheet doesn't exist, create it
    if (!sheet) {
        sheet = spreadsheet.insertSheet(calendarName);
    } else {
        // If a sheet with calendarName already exists, find a unique name and create a new sheet
        var count = 2;
        while(sheet) {
            var newName = calendarName + ' (' + count + ')';
            sheet = spreadsheet.getSheetByName(newName);
            if(!sheet) {
                sheet = spreadsheet.insertSheet(newName);
                break;
            }
            count++;
        }
    }
  
    // Set headers
    sheet.clear();
    sheet.appendRow(['EventId', 'Event Name', 'Event Start', 'Event End', 'Description', 'Location']);
  
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
            var location = event.location || "";
            
            sheet.appendRow([eventId, eventName, eventStart, eventEnd, description, location]);
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
    Logger.log('Name of the sheet: ' + sheet.name);
}

function exportCineClub() {
  exportEventsToSheet("Cine Club 2023-2024")
}

function exportMyEvents() {
    exportEventsToSheet("events","2023-11-24","2023-12-25")
}
