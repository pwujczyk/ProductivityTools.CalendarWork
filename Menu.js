function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PT.Calendar')
      .addItem('Execute for today', 'executeForToday')
      .addItem('Execute for yesterday', 'executeForYesterday')
      .addToUi();
}
