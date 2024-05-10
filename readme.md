# Google Calendar Sheet Sync

This repo contains a couple Apps Script functions that can be used to export, import, and update Google Calendars, using Google Sheets.

The function `exportEventsToSheet` goes from Calendar to Sheets. It dumps events from a Google Calendar into a new sheet of the current Google Sheets document.

The function `updateEventsFromSheet` goes from Sheets to Calendar. It updates a Google Calendar, based on events in a sheet of the current Sheets document. It can be used to edit events, to delete events, or to create new events. It is aware of the following event fields: title, description, starting, ending, description, and location. It tries to avoid accidentally clearing a field.

Both functions take optional arguments for processing events only within a certain date range.

## How to setup

In Google Sheets, select the menu item Extensions / App Script.

Within the Apps Script editor, in the left sidebar, select "Services +", and enable "Google Calendar API" version 3. (This is also known as Calendar Advances Services.) 

Within the Apps Script editor, in Project Settings, ensure that "Enable Chrome V8 runtime" is enabled.

Within the Apps Script editor, in the Editor component, copy the code into the Code.js file.

Alternatively, instead of copy/pasting, you can use a dedicated tool like `clasp` to manage AppsScript code from the command line. iirc, that also requires enabling the Google APps Script API.

## How to use

For whatever task you want, write your own entry function which calls the two utility functions. For instance the following function will dump the calendar "Cine Club 2023-2024" into a new sheet:

```js
function exportCineClub() {
  exportEventsToSheet("Cine Club 2023-2024")
}
```

After exporting the events to a sheet, you could then edit the sheet to edit the events, or to create new events or mark events for deletion. To create or delete events, add action ACTION column and set its values to "CREATE" or "DELETE" as appropriate.

Then this function will update events in that calendar from the sheet  "Cine Club 2023-2024 (2)":

```js
function fancyUpdate() {
  updateEventsFromSheet("Cine Club 2023-2024 (2)","Cine Club 2023-2024")
}
```

These scripts only support certain event fields, which are reflected in the required names for the column headers. They are as follows:  "EventId", "Event Name", "Event Start", "Event End", "Description", and "Location".


## Differences from native CSV import functionality

As of 2023-09-24T1912, Google Calendar has built-in functionality to import a CSV file. This functionality requires the following column headers: "Subject", "Start Date", "Start Time", "End Time", "Description", "Location".

Because of the difference in column headers and date and time representation, the native import format is incompatible with this system's sheets format.
