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

function executeAnalysisForLast50Days() {
  for (var e = 50; e >= 0; e--) {
    var day = 0 - e;
    executeAnalysis(day)
  }
}

function executeAnalysisForLastXDays() {
  for (var e = 50; e >= 0; e--) {
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
  const now = new Date();
  now.setHours(0, 0, 0, 0); // Ustawia na początek dzisiejszego dnia

  // Używamy setDate, która jest odporna na zmiany czasu
  const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  startDate.setDate(startDate.getDate() + daysOffsetStart);

  const endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 1);
  endDate.setMinutes(endDate.getMinutes() - 1); // Odjęcie jednej minuty, aby dostać koniec dnia

  return { start: startDate, end: endDate };
}

//common
function executeAnalysis(daysOffsetStart) {
  const { start, end } = getDateRangeForOffset(daysOffsetStart);
  clearToday(start, end);
  const caledarIds = GetCalendarsConfiguration();
  for (let e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    processAnalisys(calendarId, start, end)
  }
}

function executeConversion(daysOffsetStart) {
  const { start, end } = getDateRangeForOffset(daysOffsetStart);
  const caledarIds = GetCalendarsConfiguration();
  for (let e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    processConversion(calendarId, start, end)
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
