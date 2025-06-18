function executeForToday() {
  execute(0)
}
function executeForYesterday() {
  execute(-1)
}

function executeForLast7Days() {
  for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    execute(day)
  }
}

function executeForLast100Days() {
  for (var e = 100; e >= 0; e--) {
    var day = 0 - e;
    execute(day)
  }
}

Date.prototype.getWeekNumber = function () {
  var d = new Date(Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weeknumber = Math.ceil((((d - yearStart) / 86400000) + 1) / 7)
  var yearAndWeek = this.getFullYear() * 100 + weeknumber
  return yearAndWeek
};

// Cache for calendar configurations
var _calendarsConfigCache = null; // Consider initializing at top of script execution if needed
function GetCalendarsConfiguration() {
  if (_calendarsConfigCache !== null) {
    return _calendarsConfigCache;
  }
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Configuration-Calendars");
  if (!sheet) {
    console.error("Sheet 'Configuration-Calendars' not found.");
    return []; // Return empty array or throw error
  }
  // Get data starting from row 2, column 1, for 1 column, and all available rows
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  var values = range.getValues();
  var calendarIds = values.map(function (row) {
    return row[0];
  }).filter(function (id) { return id && id.toString().trim() !== ''; }); // Filter out empty/null/blank ids
  _calendarsConfigCache = calendarIds;
  return calendarIds;
}


function execute(daysOffsetStart) {
  //daysOffsetEnd = 0
  var MINUTE = 60 * 1000;
  var DAY = 24 * 60 * MINUTE;  // ms
  var NOW = new Date();
  NOW.setHours(0, 0, 0, 0);
  var START_DATE = new Date(NOW.getTime() + (daysOffsetStart) * DAY);
  var END_DATE = new Date(NOW.getTime() + (1 + daysOffsetStart) * DAY - MINUTE);


  var start = START_DATE;
  var end = END_DATE;
  clearToday(start, end);
  var caledarIds = GetCalendarsConfiguration();
  for (var e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    processCalendar(calendarId, start, end)
  }
}

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
      const key = row[0] ? row[0].toString().trim() : "";
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
function ConvertCalendar(event) {
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
    key = titleContentAfterHash.substring(0, firstSpaceIndex);
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
    guestsString = guestList.map(function(g) { return g.getEmail(); }).join(',');
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

function processCalendar(calendarId, start, end) {
  clearToday(start, end);
  console.log("Hello")
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calendarName = calendar.getName();
  var events = calendar.getEvents(start, end);

  var entries = {};
  for (var e = 0; e < events.length; e++) {
    var event = events[e];
    var status = event.getMyStatus().toString();
    var type = event.getEventType().toString();
    var start = event.getStartTime();
    var end = event.getEndTime();

    var title = event.getTitle();
    console.log(title);



    var duration = (end - start) / 3600000;
    var color = event.getColor();
    var day = Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM-dd');
    var weeknumber = start.getWeekNumber()
    var month = Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM');
    var dayLog = { start: start, end: end, day: day, weeknumber: weeknumber, month: month, duration: duration, title: title, calendarName: calendarName, status: status, type: type, color: color }
    //console.log(dayLog);
    var category = getCategory(dayLog)

      if (dayLog.calendarName === "DailyLog" || dayLog.calendarName === "DataPoints") {
        ReplaceTitleWithCategory(event,category)
    }

    var value = getValue(dayLog)
    var dayLog = { ...dayLog, category: category, value: value }

    //collor=1 - do not count
    var ownerAcceptedValue = ownerAccepted(event, calendarName, 'pwujczyk@google.com')
    var myStatus = event.getMyStatus() == CalendarApp.GuestStatus.YES;
    var myStatus1 = event.getMyStatus() == CalendarApp.GuestStatus.NO;
    var myStatus2 = event.getMyStatus() == CalendarApp.GuestStatus.OWNER;


    var x2 = event.isOwnedByMe();
    console.log("ownerAcceptedValue", ownerAcceptedValue)
    if (type != "WORKING_LOCATION" && type != "OUT_OF_OFFICE" && status != "INVITED" && status != "NO" && color != 1 && ownerAcceptedValue) {
      SaveItem(dayLog)
    }


  }
  return entries;
}

function ownerAccepted(event, calendarName, owner) {

  if (calendarName != owner) {
    return true;
  }

  var guestList = event.getGuestList();
  //calendar item without anybody invited
  if (guestList.length == 0) {
    return true;
  }

  var pwujczykAccepted = event.getGuestList(true).some(g =>
    g.getEmail() == "pwujczyk@google.com" &&
    g.getGuestStatus() == CalendarApp.GuestStatus.YES);

  return pwujczykAccepted;
  //  var myStatus=event.getMyStatus();
  //  var mystatus=event.getMyStatus()== CalendarApp.GuestStatus.OWNER;
  //   if (event.getMyStatus()== CalendarApp.GuestStatus.OWNER)
  //   {
  //     return true
  //   }

  //     if (event.getMyStatus()== CalendarApp.GuestStatus.YES)
  //   {
  //     return true
  //   }


  //   var pwujczykAccepted = guestList.filter(function (guest) { return guest.getEmail() == 'pwujczyk@google.com' && guest.getGuestStatus() === CalendarApp.GuestStatus.YES; });
  //   if (pwujczykAccepted.length > 0) {
  //     return true;
  //   }


  //  for (var i = 0; i < guestList.length; i++) {
  //     var guest = guestList[i];

  //     var dsfa = guest.getEmail();
  //     var xxx=guest.getName();
  //     var x1 = guest.getGuestStatus() === CalendarApp.GuestStatus.YES;

  //     var guestStatus = guest.getGuestStatus()
  //     var yes = guestStatus.YES.toString();
  //     console.log("GuestStatus")
  //   }
}

// Cache for DailyLog configuration
var _dailyLogConfigCache = null;

function LoadDailyLogConfiguration() {
  if (_dailyLogConfigCache !== null) {
    return _dailyLogConfigCache;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Mapping-DailyLog";
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    console.error("Sheet '" + sheetName + "' not found.");
    _dailyLogConfigCache = {}; // Cache empty object to avoid re-checking
    return _dailyLogConfigCache;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) { // Need at least header + 1 data row
    console.warn("Sheet '" + sheetName + "' has no data rows (or only a header). Ensure header 'key,category' exists and there is data.");
    _dailyLogConfigCache = {};
    return _dailyLogConfigCache;
  }

  values.shift(); // Remove header row (e.g., "key", "category")

  const configMap = {};
  values.forEach(function (row) {
    if (row.length >= 2) {
      const key = row[0] ? row[0].toString().trim() : "";
      const category = row[1] ? row[1].toString().trim() : "";
      if (key) { // Only add if key is not empty
        configMap[key] = category;
      }
    }
  });

  _dailyLogConfigCache = configMap;
  return _dailyLogConfigCache;
}

function ReplaceTitleWithCategory(event, category) {
  if (!category) {
    return; // Don't do anything if no category was found.
  }

  const originalTitle = event.getTitle();
  const colonIndex = originalTitle.indexOf(':');

  let newTitle;
  if (colonIndex === -1) {
    // No colon, replace the whole title with the category.
    newTitle = category;
  } else {
    // Colon found, replace the part before it.
    const valuePart = originalTitle.substring(colonIndex); // e.g., ": 1 hour"
    newTitle = category + valuePart;
  }

  if (originalTitle !== newTitle) {
    event.setTitle(newTitle);
    console.log("Updated event title from '" + originalTitle + "' to '" + newTitle + "'.");
  }
}


function GetDailyLogCategory(dayLog) {
  const title = dayLog.title;
  const dailyLogConfig = LoadDailyLogConfiguration();
  if (title) {
    if (title.indexOf(':') === -1) {
      var r = dailyLogConfig.hasOwnProperty(title) ? dailyLogConfig[title] : null;
      return r;
    }
    const key = title.substring(0, title.indexOf(':')).trim();

    let result = dailyLogConfig.hasOwnProperty(key) ? dailyLogConfig[key] : null;
    return result;
  }
  return null;
}

function GetCaldendarsCategory(dayLog) {
  var configuration = LoadConfiguration();
  for (var e = 0; e < configuration.length; e++) {
    var conf = configuration[e];
    if (conf.column == "Title") {
      if (dayLog.title == conf.value) {
        var returnValue = conf.category
        return returnValue;
      }
    }


    if (conf.column == "Color") {
      if (dayLog.color == conf.value) {
        var returnValue = conf.category
        return returnValue;
      }
    }


    if (conf.column == "CalendarName") {
      if (dayLog.calendarName == conf.value) {
        var returnValue = conf.category
        return returnValue;
      }
    }
  }
}

function GetDailyLogValue(dayLog) {
  const title = dayLog.title;
  if (!title || title.indexOf(':') === -1) {
    // console.log("GetDailyLogCategory: Title '" + title + "' does not contain ':' or is empty.");
    return null; // No colon, so no key to extract
  }
  const parts = title.split(':');

  // If there are at least two parts (i.e., a delimiter was found and there's content after it)
  if (parts.length > 1) {
    // Return the second part (index 1)
    var r = parts[1].trim();
    return r;
  } else {
    // If no delimiter or no content after the delimiter, return null
    console.warn(`No second part found for string: "${inputString}".`);
    return null;
  }
}

function GetCalendarValue(dayLog) {

}


function getValue(dayLog) {
  if (dayLog.calendarName === "DailyLog" || dayLog.calendarName === "DataPoints") {
    var r = GetDailyLogValue(dayLog);
    return r;
  }
  else {
    var r = GetCalendarValue(dayLog);
    return r;

  }
}

function getCategory(dayLog) {
  if (dayLog.calendarName === "DailyLog" || dayLog.calendarName === "DataPoints") {
    var r = GetDailyLogCategory(dayLog);
    return r;
  }
  else {
    var r = GetCaldendarsCategory(dayLog);
    return r;

  }
}

// Cache for general configuration
var _generalConfigCache = null;
function LoadConfiguration() {
  if (_generalConfigCache !== null) {
    return _generalConfigCache;
  }
  //var configuration = SpreadsheetApp.openById("1-qb1wmRiDWJTq5n5T3ItkWJmHjU-BhGYFe9e439MFtc");
  var configuration = SpreadsheetApp.getActiveSpreadsheet();
  //var vacations = configuration.getSheetByName("Vacations");
  var conf = configuration.getSheetByName("Mapping-Caldendars");
  // TODO: Add check if conf sheet exists, similar to other config loaders

  var data = conf.getDataRange().getValues();
  data.shift();// Remove header 
  var items = [];

  data.forEach(function (row) {
    var element = { 'column': row[0], 'value': row[1], 'category': row[2] }
    items.push(element);
  });
  _generalConfigCache = items;
  return items;
}


function SaveItem(dayLog) {

  getSheet().appendRow([dayLog.start, dayLog.end, dayLog.day, dayLog.weeknumber, dayLog.month, dayLog.duration, dayLog.title, dayLog.calendarName, dayLog.status, dayLog.type, dayLog.color, dayLog.category, dayLog.value]);
}

function getSheet() {
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var daily = file.getSheetByName("Days");
  return daily;
}

function clearToday(start, end) {
  var sheet = getSheet()
  var data = sheet.getDataRange().getValues();
  for (i = data.length - 1; i > 0; i--) {
    var x = i;
    var x1 = data[i];
    var lineStart = data[i][0]
    var lineEnd = data[i][1]
    if (start <= lineStart && lineStart <= end) {
      console.log(start);
      sheet.deleteRow(i + 1)
    }
  }
}