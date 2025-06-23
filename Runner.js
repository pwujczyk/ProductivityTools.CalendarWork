function executeAnalysisForToday() {
  execute(0,processCalendar)
}

function executeAnalysisForYesterday() {
  execute(-1,processCalendar)
}

function executeAnalysisForLast7Days() {
  for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    execute(day,processCalendar)
  }
}

function executeAnalysisForLast100Days() {
  for (var e = 100; e >= 0; e--) {
    var day = 0 - e;
    execute(day,processCalendar)
  }
}


function executeConversionForToday() {
  execute(0,ConvertCalendar)
}

function executeConversionForLast7Days() {
    for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    execute(day,ConvertCalendar)
  }
}

function executeConversionForLast100Days() {
    for (var e = 100; e >= 0; e--) {
    var day = 0 - e;
    execute(day,ConvertCalendar)
  }
}


//common
function execute(daysOffsetStart, fn) {
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
    fn(calendarId, start, end)
    //processCalendar(calendarId, start, end)
  }
}

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

