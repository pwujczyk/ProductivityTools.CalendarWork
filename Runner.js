function executeAnalysisForToday() {
  executeAnalysis(0)
}

function executeAnalysisForYesterday() {
  executeAnalysis(-1)
}

function executeAnalysisForLast7Days() {
  for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    executeAnalysis(day)
  }
}

function executeAnalysisForLast100Days() {
  for (var e = 100; e >= 0; e--) {
    var day = 0 - e;
    executeAnalysis(day)
  }
}


function executeConversionForToday() {
  executeConversion(0)
}

function executeConversionForLast7Days() {
    for (var e = 7; e >= 0; e--) {
    var day = 0 - e;
    executeConversion(day)
  }
}

function executeConversionForLast100Days() {
    for (var e = 100; e >= 0; e--) {
    var day = 0 - e;
    executeConversion(day)
  }
}


function getDateRangeForOffset(daysOffsetStart) {
  const MINUTE_IN_MS = 60 * 1000;
  const DAY_IN_MS = 24 * 60 * MINUTE_IN_MS;
  const now = new Date();
  now.setHours(0, 0, 0, 0); // Set to the beginning of today
  const startDate = new Date(now.getTime() + (daysOffsetStart * DAY_IN_MS));
  const endDate = new Date(now.getTime() + ((1 + daysOffsetStart) * DAY_IN_MS) - MINUTE_IN_MS);
  return { start: startDate, end: endDate };
}

//common
function executeAnalysis(daysOffsetStart) {
  const { start, end } = getDateRangeForOffset(daysOffsetStart);
  clearToday(start, end);
  const caledarIds = GetCalendarsConfiguration();
  for (let e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    fn(calendarId, start, end)
  }
}

function executeConversion(daysOffsetStart) {
  const { start, end } = getDateRangeForOffset(daysOffsetStart);
  const caledarIds = GetCalendarsConfiguration();
  for (let e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    fn(calendarId, start, end)
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
