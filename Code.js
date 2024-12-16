MONTHS_PAST=1
MONTHS_FUTURE=0
var DAY = 24 * 60 * 60 * 1000;  // ms
var MONTH = 10 * DAY;  // ms
var NOW = new Date();
var START_DATE = new Date(NOW.getTime() - MONTHS_PAST * MONTH);
var END_DATE = new Date(NOW.getTime() + MONTHS_FUTURE * MONTH);

var caledarIds=['c_0f4cb3f8b97b7a808d0da14c2a98dee84e6612cef687984a70959059b1fa33b2@group.calendar.google.com','pwujczyk@google.com']

function process(){
    for (var e = 0; e < caledarIds.length; e++) {
      var calendarId=caledarIds[e];
      processCalendar(calendarId)
    }
}

function processCalendar(calendarId) {
  console.log("Hello")
  var calendar=CalendarApp.getCalendarById(calendarId);
  var calendarName=calendar.getName();
  var events = calendar.getEvents(START_DATE, END_DATE);
  
  var entries = {};
  for (var e = 0; e < events.length; e++) {
     var event = events[e];
     console.log(event.getColor())
     console.log(event.getTitle())
     console.log("---")
     var start=event.getStartTime();
     var end=event.getEndTime();
     var duration=(end-start)/3600000;
     var title=event.getTitle();
     var day= Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM-dd');
     var dayLog={start:start,end:end,day:day, duration: duration, title:title, category:calendarName}
     console.log(dayLog);
    // var partner = get1on1Partner(event);
    // if (!partner) continue;
    // var guest = event.getGuestByEmail(toEmail(ME));
    // if (!guest || guest.getGuestStatus() ==
    //     CalendarApp.GuestStatus.NO) continue;
    // if (!entries[partner]) {
    //   entries[partner] = {
    //     name: partner,
    //     recent: null,
    //     upcoming: null,
    //     row: [],
    //   };
    // }
    
    // var endTime = event.getEndTime();
    // if (endTime > NOW) {
    //   var oldTime = entries[partner]['upcoming'];
    //   entries[partner]['upcoming'] = timeMax(endTime, oldTime, false);
    // } else {
    //   var oldTime = entries[partner]['recent'];
    //   entries[partner]['recent'] = timeMax(endTime, oldTime, true);
    // }
  }
  return entries;
}


function SaveItem(){
  var file = SpreadsheetApp.getActiveSpreadsheet();
   var daily = file.getSheetByName("Days");
   daily.appendRow("fdsafa");
}