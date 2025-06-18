function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PT.Calendar')
      .addItem('Clean & Execute analysis for today', 'executeAnalysisForToday')
      .addItem('Clean & Execute analysis for yesterday', 'executeAnalysisForYesterday')
      .addItem('Clean & Execute analysis for Last 7 days', 'executeAnalysisForLast7Days')
      .addItem('Clean & Execute analysis for Last 100 days', 'executeAnalysisForLast100Days')
      .addItem('Execute conversion for today', 'executeConversionForToday')
      .addItem('Execute conversion for Last 7 days', 'executeConversionForLast7Days')
      .addItem('Clean & Execute analysis for Last 100 days', 'executeConversionForLast100Days')

      
      .addToUi();
}
