
// Cache for ConvertCalendar configuration
var _convertCalendarConfigCache = null;

/**
 * Loads the ConvertCalendar configuration from the "Mapping-ConvertCalendar" sheet.
 * The sheet is expected to have "Key" in the first column and "TargetCalendarId" in the second.
 * @return {Object} A map where keys are strings from the "Key" column and 
 *                  values are target calendar IDs.
 */
function LoadConvertCalendarConfiguration() {
  if (_convertCalendarConfigCache !== null) {
    return _convertCalendarConfigCache;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Mapping-ConvertCalendar";
  const sheet = spreadsheet.getSheetByName(sheetName);

  _convertCalendarConfigCache = {}; // Initialize cache, even if sheet is not found

  if (!sheet) {
    console.error("Sheet '" + sheetName + "' not found. ConvertCalendar functionality will not work without this mapping.");
    return _convertCalendarConfigCache;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) { // Need at least header + 1 data row
    console.warn("Sheet '" + sheetName + "' has no data rows (or only a header). Ensure header (e.g., 'Key,TargetCalendarId') exists and there is data.");
    return _convertCalendarConfigCache;
  }

  values.shift(); // Remove header row

  const configMap = {};
  values.forEach(function (row, index) {
    if (row.length >= 2) {
      const key = row[0] ? row[0].toString().trim().toLowerCase() : "";
      const targetCalendarId = row[1] ? row[1].toString().trim() : "";
      if (key && targetCalendarId) {
        if (configMap.hasOwnProperty(key)) {
          console.warn("In 'Mapping-ConvertCalendar', duplicate key '" + key + "' found at row " + (index + 2) + ". Previous value will be overwritten.");
        }
        configMap[key] = targetCalendarId;
      } else {
        if (key || targetCalendarId) { // Log if one is present but the other is missing, and it's not an entirely blank row
          console.warn("In 'Mapping-ConvertCalendar', row " + (index + 2) + " has incomplete data: Key='" + key + "', TargetCalendarId='" + targetCalendarId + "'. Skipping this row.");
        }
      }
    }
  });

  _convertCalendarConfigCache = configMap;
  // console.log("Loaded ConvertCalendar Configuration: ", _convertCalendarConfigCache); // Uncomment for debugging
  return _convertCalendarConfigCache;
}

/**
 * Checks if an event title starts with '#' and, if so, attempts to create a new event
 * in a target calendar based on configuration in "Mapping-ConvertCalendar".
 * The new event will have properties copied from the original, with the title modified
 * to remove the '#' prefix and key.
 * @param {CalendarEvent} event The calendar event to process.
 */


function processConversion(calendarId, start, end) {
  console.log("Hello starting convert calendar for calendarid:", calendarId)
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calendarName = calendar.getName();
  var events = calendar.getEvents(start, end);

  var entries = {};
  for (var e = 0; e < events.length; e++) {
    var event = events[e];
    ConvertCalendarEvent(event);
  }
}

function ConvertCalendarEvent(event) {
  const originalTitle = event.getTitle();
  if (!originalTitle || !originalTitle.startsWith('#')) {
    return; // Not an event to convert
  }

  const titleContentAfterHash = originalTitle.substring(1);
  const firstSpaceIndex = titleContentAfterHash.indexOf(' ');

  let key;
  let newEventTitle;

  if (firstSpaceIndex === -1) { // e.g., "#ProjectX"
    key = titleContentAfterHash;
    newEventTitle = key; // Use the key itself as the new title
  } else { // e.g., "#ProjectX Meeting"
    key = titleContentAfterHash.substring(0, firstSpaceIndex).toLowerCase();
    const restOfTitle = titleContentAfterHash.substring(firstSpaceIndex + 1).trim();
    newEventTitle = restOfTitle === "" ? key : restOfTitle; // If "#Key ", title becomes "Key"
  }

  if (!key) { // e.g. title was just "#"
    console.warn("Event title '" + originalTitle + "' starts with '#' but could not extract a valid key. Skipping conversion.");
    return;
  }

  const config = LoadConvertCalendarConfiguration();
  if (!config.hasOwnProperty(key)) {
    // console.log("No conversion rule found for key '" + key + "' from title '" + originalTitle + "'. Skipping conversion."); // Optional: for debugging
    return;
  }

  const targetCalendarName = config[key];
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];

  if (!targetCalendar) {
    console.error("Target calendar with ID '" + targetCalendarId + "' for key '" + key + "' not found or access denied. Skipping conversion for event: " + originalTitle);
    return;
  }

  const startTime = event.getStartTime();
  const endTime = event.getEndTime();
  const description = event.getDescription();
  const location = event.getLocation();

  const guestList = event.getGuestList();
  let guestsString = "";
  if (guestList && guestList.length > 0) {
    guestsString = guestList.map(function (g) { return g.getEmail(); }).join(',');
  }

  const options = {
    description: description,
    location: location
  };
  if (guestsString !== "") {
    options.guests = guestsString;
    options.sendInvites = false; // Crucial: set to false to avoid re-inviting guests
  }

  try {
    var newEvent;
    if (event.isAllDayEvent()) {
      newEvent = targetCalendar.createAllDayEvent(newEventTitle, startTime, endTime, options);
    } else {
      newEvent = targetCalendar.createEvent(newEventTitle, startTime, endTime, options);
    }
    console.log("Event '" + originalTitle + "' (new title: '" + newEventTitle + "') successfully created in calendar '" + targetCalendar.getName() + "' (ID: " + newEvent.getId() + ")");

    // const originalColor = event.getColor();
    // if (originalColor && newEvent.setColor) {
    //     try {
    //         newEvent.setColor(originalColor);
    //     } catch (e) {
    //         console.warn("Could not set color '" + originalColor + "' for new event '" + newEventTitle + "': " + e.toString());
    //     }
    // }

    // const attachments = event.getAttachments();
    // if (attachments && attachments.length > 0 && newEvent.addAttachment && newEvent.addDriveAttachment) {
    //     attachments.forEach(function(attachment) {
    //         try {
    //             if (typeof MimeType !== 'undefined' && attachment.getMimeType() === MimeType.GOOGLE_DRIVE) {
    //                 const fileId = attachment.getFileId();
    //                 if (fileId) {
    //                     newEvent.addDriveAttachment(fileId);
    //                 } else {
    //                     console.warn("Drive attachment '" + attachment.getTitle() + "' has no file ID, cannot copy to new event '" + newEventTitle + "'.");
    //                 }
    //             } else {
    //                 newEvent.addAttachment(attachment.getBlob()); // This might fail for some attachment types or if blob access is restricted.
    //             }
    //         } catch (e) {
    //             console.warn("Could not copy attachment '" + attachment.getTitle() + "' to new event '" + newEventTitle + "': " + e.toString());
    //         }
    //     });
    // }
    // Note: The original event is NOT deleted or modified by this function.
    // If you want to "move" the event, you would call event.deleteEvent() here after successful creation.
    event.deleteEvent();
  } catch (e) {
    console.error("Failed to create event '" + newEventTitle + "' in calendar '" + targetCalendar.getName() + "'. Original title: '" + originalTitle + "'. Error: " + e.toString());
  }
}