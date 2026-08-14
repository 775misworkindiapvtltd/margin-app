function copyDarkRedRow19Formulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('QUOT');

  if (!sheet) {
    throw new Error('QUOT sheet nahi mili.');
  }

  var redColor = '#e60a18';
  var markerRow = 19;
  var sourceRow = 18;
  var firstTargetRow = 20;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  // Row 19 ke colors read karo
  var colors = sheet
    .getRange(markerRow, 1, 1, lastColumn)
    .getBackgrounds()[0];

  // Sirf row 19 ke exact #e60a18 color wale columns par kaam hoga
  for (var column = 1; column <= lastColumn; column++) {
    var color = String(colors[column - 1] || '')
      .trim()
      .toLowerCase();

    if (color !== redColor) {
      continue;
    }

    var sourceCell = sheet.getRange(sourceRow, column);

    // Row 18 ka formula row 20 se last row tak copy hoga.
    // Destination blank ho ya filled, dono cases mein overwrite hoga.
    for (var row = firstTargetRow; row <= lastRow; row++) {
      sourceCell.copyTo(
        sheet.getRange(row, column),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
    }
  }
}
