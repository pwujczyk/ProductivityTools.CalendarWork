

var caledarIds = ['c_0f4cb3f8b97b7a808d0da14c2a98dee84e6612cef687984a70959059b1fa33b2@group.calendar.google.com'
  , 'c_61629aa73c878650c1e66cefeefa354bb85696005a69c45fc7b3f2bf2f8130c3@group.calendar.google.com'
  , 'c_c836fe4957a38740ddbd08a2b537cee5f9b630b97e5862f01d99dff0719b4ee5@group.calendar.google.com'
  , 'c_833af64a3dc6013a751692a881ef3f8b8e39f27751fb2fc2acb00b416505e0b3@group.calendar.google.com'
  , 'c_9ad969d46441f6da4e92934b0a43a4c395c6724d74df8f4fee38f2427699b891@group.calendar.google.com'
  , 'c_fd2e2e2a38f584641f3b838fbc32d70e450d665637ec63a13d0aac8393c4dbd6@group.calendar.google.com'
  , 'pwujczyk@google.com']

function executeForToday() {
  execute(0)
}
function executeForYesterday() {
  execute(-1)
}

Date.prototype.getWeekNumber = function(){
  var d = new Date(Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weeknumber= Math.ceil((((d - yearStart) / 86400000) + 1)/7)
  var yearAndWeek=this.getFullYear()*100+weeknumber
  return yearAndWeek
};


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
    var type = event.getEventType().toString();
    var start = event.getStartTime();
    var end = event.getEndTime();

    var title = event.getTitle();
    console.log(title);



    var duration = (end - start) / 3600000;
    var color = event.getColor();
    var day = Utilities.formatDate(start, 'Europe/Warsaw', 'yyyy-MM-dd');
    var weeknumber=start.getWeekNumber()
    var dayLog = { start: start, end: end, day: day, weeknumber: weeknumber, duration: duration,  title: title, calendarName: calendarName, status: status, type: type, color: color }
    //console.log(dayLog);
    var category = getCategory(dayLog)
    var dayLog = { ...dayLog, category: category }

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


function getCategory(dayLog) {
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

function LoadConfiguration() {
  //var configuration = SpreadsheetApp.openById("1-qb1wmRiDWJTq5n5T3ItkWJmHjU-BhGYFe9e439MFtc");
  var configuration = SpreadsheetApp.getActiveSpreadsheet();
  //var vacations = configuration.getSheetByName("Vacations");
  var conf = configuration.getSheetByName("Configuration");

  var data = conf.getDataRange().getValues();
  data.shift();// Remove header 
  var items = [];

  data.forEach(function (row) {
    var element = { 'column': row[0], 'value': row[1], 'category': row[2] }
    items.push(element);
  });
  return items;
}


function SaveItem(dayLog) {

  getSheet().appendRow([dayLog.start, dayLog.end, dayLog.day, dayLog.weeknumber, dayLog.duration,  dayLog.title, dayLog.calendarName, dayLog.status, dayLog.type, dayLog.color, dayLog.category]);
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
    var lineStart = data[i][0]
    var lineEnd = data[i][1]
    if (start < lineStart && lineStart < end) {
      console.log(start);
      sheet.deleteRow(i + 1)
    }
  }
}