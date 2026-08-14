/**
 * Copies formulas for every column marked dark red (#e60a18) in row 19.
 *
 * Row 18 is treated as the formula template row. The template formula from
 * each marked column is copied to row 20 through the sheet's last used row.
 * Row 19 is never overwritten because it is the marker/header row.
 *
 * The target sheet in this project is named QUOT. A sheet name can also be
 * supplied explicitly, for example:
 *   copyDarkRedRow19Formulas('QUOT');
 *
 * @param {string=} sheetName Optional target sheet name. QUOT is used when
 *   omitted.
 * @return {Object} A summary of the columns and cells updated.
 */
function copyDarkRedRow19Formulas(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet is available. Open the QUOT spreadsheet and try again.');
  }

  // QUOT is the target tab. A Sheet object is also accepted if another
  // function calls this helper with SpreadsheetApp.getActiveSheet().
  var target = sheetName || 'QUOT';
  var sheet = typeof target === 'object' && typeof target.getRange === 'function'
    ? target
    : ss.getSheetByName(String(target));
  if (!sheet || typeof sheet.getRange !== 'function') {
    throw new Error('Sheet not found: ' + String(target) + '. Expected the sheet tab named QUOT.');
  }

  var markerColor = '#e60a18';
  var markerRow = 19;
  var sourceRow = 18;
  var firstTargetRow = 20;
  var dataRange = sheet.getDataRange();
  var lastRow = dataRange.getRow() + dataRange.getNumRows() - 1;
  var lastColumn = dataRange.getColumn() + dataRange.getNumColumns() - 1;

  if (sheet.getMaxRows() < markerRow || lastColumn < 1) {
    return {
      status: 'nothing_to_copy',
      sheetName: getSheetName_(sheet),
      markerColor: markerColor,
      markerColumns: [],
      copiedCells: 0,
      lastRow: lastRow
    };
  }

  var markerBackgrounds = sheet
    .getRange(markerRow, 1, 1, lastColumn)
    .getBackgrounds()[0];
  var sourceFormulas = sheet
    .getRange(sourceRow, 1, 1, lastColumn)
    .getFormulas()[0];

  var markerColumns = [];
  var copiedColumns = [];
  var skippedColumns = [];
  var copiedCells = 0;

  for (var column = 1; column <= lastColumn; column++) {
    var background = String(markerBackgrounds[column - 1] || '').trim().toLowerCase();
    if (background !== markerColor) continue;

    markerColumns.push(columnToLetter_(column));
    var formula = String(sourceFormulas[column - 1] || '').trim();
    if (!formula) {
      skippedColumns.push({
        column: columnToLetter_(column),
        reason: 'No formula found in row ' + sourceRow
      });
      continue;
    }

    if (lastRow < firstTargetRow) {
      copiedColumns.push(columnToLetter_(column));
      continue;
    }

    var sourceCell = sheet.getRange(sourceRow, column);
    for (var row = firstTargetRow; row <= lastRow; row++) {
      // PASTE_FORMULA preserves relative references when the row-18 formula is
      // copied into each destination row and does not overwrite formatting.
      sourceCell.copyTo(
        sheet.getRange(row, column),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
      copiedCells++;
    }
    copiedColumns.push(columnToLetter_(column));
  }

  return {
    status: copiedColumns.length ? 'ok' : 'nothing_to_copy',
    sheetName: getSheetName_(sheet),
    markerColor: markerColor,
    markerRow: markerRow,
    sourceRow: sourceRow,
    firstTargetRow: firstTargetRow,
    lastRow: lastRow,
    markerColumns: markerColumns,
    copiedColumns: copiedColumns,
    skippedColumns: skippedColumns,
    copiedCells: copiedCells
  };
}

/**
 * Returns a display name without making the repair depend on getName().
 *
 * @param {Object} sheet Google Sheets Sheet object.
 * @return {string} Sheet name when available.
 */
function getSheetName_(sheet) {
  return typeof sheet.getName === 'function' ? sheet.getName() : 'QUOT';
}

/**
 * Converts a 1-based sheet column number to its A1 column letters.
 *
 * @param {number} columnNumber 1-based column number.
 * @return {string} A1 column letters.
 */
function columnToLetter_(columnNumber) {
  var letters = '';
  var number = columnNumber;
  while (number > 0) {
    var remainder = (number - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    number = Math.floor((number - 1) / 26);
  }
  return letters;
}
