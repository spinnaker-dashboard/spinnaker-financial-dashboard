function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Dashboard')
    .addItem('run validation...', 'runBigQueryValidationv3')
    .addToUi();
}
