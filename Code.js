DAYS_PAST = 1
DAYS_FUTURE = 0
var DAY = 24 * 60 * 60 * 1000;  // ms
var NOW = new Date().setHours(0,0,0,0);;
var START_DATE = new Date(NOW.getTime() - DAYS_PAST * DAY);
var END_DATE = new Date(NOW.getTime() + DAYS_FUTURE * DAY);

var caledarIds = ['c_0f4cb3f8b97b7a808d0da14c2a98dee84e6612cef687984a70959059b1fa33b2@group.calendar.google.com', 'pwujczyk@google.com']

function process() {
  var start = START_DATE;
  var end = END_DATE;
  clearToday(start, end);
  for (var e = 0; e < caledarIds.length; e++) {
    var calendarId = caledarIds[e];
    processCalendar(calendarId, start, end)
  }
}

function processCalendar(calendarId, start, end) {
  console.log("Hello")
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calendarName = calendar.getName();
  var events = calendar.getEvents(start, end);

  var entries = {};
  for (var e = 0; e < events.length; e++) {
    var event = events[e];
    var status = event.getMyStatus().toString();
    //console.log(event.getTag())
    var type = event.getEventType().toString();
    //console.log(event.getColor())
    //console.log(event.getTitle())
    //console.log("---")
    var start = event.getStartTime();
    var end = event.getEndTime();
    var duration = (end - start) / 3600000;
    var title = event.getTitle();
    var color = event.getColor();
    var day = Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM-dd');
    var dayLog = { start: start, end: end, day: day, duration: duration, title: title, category: calendarName, status: status, type: type, color: color }
    //console.log(dayLog);
    SaveItem(dayLog)

  }
  return entries;
}


function SaveItem(dayLog) {

  getSheet().appendRow([dayLog.start, dayLog.end, dayLog.day, dayLog.duration, dayLog.title, dayLog.category, dayLog.status, dayLog.type, dayLog.color]);
}

function getSheet() {
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var daily = file.getSheetByName("Days");
  return daily;
}

function clearToday(sheet, start, end) {
  var sheet = getSheet()
  var data = sheet.getDataRange().getValues();
  for (i = data.length - 1; i > 0; i--) {
    var lineStart = data[i][0]
    var lineEnd = data[i][1]
    if (start < lineStart && lineStart < end) {
      console.log(start);
      sheet.deleteRow(i + 1)
    }
  }
}