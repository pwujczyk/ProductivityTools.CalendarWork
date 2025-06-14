function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PT.Calendar')
      .addItem('Clean & Execute for today', 'executeForToday')
      .addItem('Clean & Execute for yesterday', 'executeForYesterday')
      .addItem('Clean & Execute for Last 7 days', 'executeForLast7Days')
      .addToUi();
}
