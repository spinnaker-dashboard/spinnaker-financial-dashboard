function applyGrowthTextColorFormatting_v2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Rows to format
  var rows = [7, 8, 9, 10, 13, 14, 15];

  // Fixed columns for growth
  var rawCol = 'H'; // Raw Growth
  var bqCol = 'N';  // BigQuery Growth

  // Helper function to color cell text based on sign
  function setCellSignColor(range) {
    var value = range.getValue();
    if (typeof value === 'number' && !isNaN(value)) {
      if (value > 0) range.setFontColor('#008000');   // green
      else if (value < 0) range.setFontColor('#FF0000'); // red
      else range.setFontColor('#000000');              // zero -> black
    } else {
      range.setFontColor('#000000'); // non-numeric (blank, text, etc.)
    }
  }

  // Apply to each row for both columns
  rows.forEach(function(r) {
    setCellSignColor(sheet.getRange(rawCol + r));
    setCellSignColor(sheet.getRange(bqCol + r));
  });
}
