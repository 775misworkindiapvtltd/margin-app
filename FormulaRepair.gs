/**
 * Copies formulas for every column marked dark red (#e60a18) in row 19.
 *
 * Row 18 is treated as the formula template row. The template formula from
 * each marked column is copied to row 20 through the sheet's last used row.
 * Row 19 is never overwritten because it is the marker/header row.
 *
 * Run copyDarkRedRow19Formulas() while the required quotation sheet is active.
 * A sheet name can also be supplied, for example:
 *   copyDarkRedRow19Formulas('QUOTATION');
 *
 * @param {string=} sheetName Optional target sheet name. Active sheet is used
 *   when omitted.
 * @return {Object} A summary of the columns and cells updated.
 */
function copyDarkRedRow19Formulas(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(String(sheetName)) : ss.getActiveSheet();
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  var markerColor = '#e60a18';
  var markerRow = 19;
  var sourceRow = 18;
  var firstTargetRow = 20;
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), sheet.getMaxColumns());

  if (sheet.getMaxRows() < markerRow || lastColumn < 1) {
    return {
      status: 'nothing_to_copy',
      sheetName: sheet.getName(),
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
    sheetName: sheet.getName(),
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
