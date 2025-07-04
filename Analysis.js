

Date.prototype.getWeekNumber = function () {
  var d = new Date(Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weeknumber = Math.ceil((((d - yearStart) / 86400000) + 1) / 7)
  var yearAndWeek = this.getFullYear() * 100 + weeknumber
  return yearAndWeek
};


function processAnalisys(calendarId, start, end) {
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
      ReplaceTitleWithCategory(event, category)
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

function GetDailyLogCategory(dayLog) {
  const title = dayLog.title;
  const dailyLogConfig = LoadDailyLogConfiguration();
  if (title) {
    if (title.indexOf(':') === -1) {
      var r = dailyLogConfig.hasOwnProperty(title) ? dailyLogConfig[title.toLowerCase()] : null;
      return r;
    }
    const key = title.substring(0, title.indexOf(':')).trim();

    let result = dailyLogConfig.hasOwnProperty(key) ? dailyLogConfig[key.toLowerCase()] : null;
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
    var r = parts[1];
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