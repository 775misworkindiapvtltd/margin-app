function copyLightRedRow20AndClearBelow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('QUOT');

  if (!sheet) {
    throw new Error('QUOT sheet nahi mili.');
  }

  var targetColor = '#f4cccc';
  var colorRow = 19;
  var sourceRow = 18;
  var pasteRow = 20;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getMaxColumns();

  // Row 19 ke colors read karo
  var colors = sheet
    .getRange(colorRow, 1, 1, lastColumn)
    .getBackgrounds()[0];

  for (var column = 1; column <= lastColumn; column++) {
    var color = String(colors[column - 1] || '')
      .trim()
      .toLowerCase();

    // Sirf exact #f4cccc color wale columns par kaam hoga
    if (color !== targetColor) {
      continue;
    }

    // Row 18 ka formula sirf row 20 mein paste karo
    sheet
      .getRange(sourceRow, column)
      .copyTo(
        sheet.getRange(pasteRow, column),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );

    // Row 21 se last row tak data clear karo.
    // Formatting clear nahi hogi, sirf values/formulas clear honge.
    if (lastRow >= 21) {
      sheet
        .getRange(21, column, lastRow - 20, 1)
        .clearContent();
    }
  }
}
