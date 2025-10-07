function applyGrowthTextColorFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Raw growth (H7:H15)
  var rawCells = ['H7','H8','H9','H10','H13','H14','H15'];

  // BigQuery growth (K7:K15)
  var bqCells = ['K7','K8','K9','K10','K13','K14','K15'];

  function setTextColor(cells) {
    cells.forEach(function(addr) {
      var cell = sheet.getRange(addr);
      var val = cell.getValue();
      if (typeof val === 'number') {
        if (val > 0) cell.setFontColor('green');
        else if (val < 0) cell.setFontColor('red');
        else cell.setFontColor('black');
      } else {
        cell.setFontColor('black');
      }
    });
  }

  setTextColor(rawCells);
  setTextColor(bqCells);
}
