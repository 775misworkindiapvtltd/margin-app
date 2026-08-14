/**
 * Finds the exact #e60a18 background in row 19 of the QUOT sheet and copies
 * the row-18 formula only in those matching columns.
 *
 * Row 19 is used only as the color marker row. Row 18 is always the formula
 * source, and the formula is copied to row 20 through the last used row.
 */
function copyDarkRedRow19Formulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet is available.');
  }

  // Important: this returns a Sheet object. Do not use var sheet = 'QUOT'.
  var sheet = ss.getSheetByName('QUOT');
  if (!sheet) {
    throw new Error('Sheet tab "QUOT" was not found.');
  }

  var markerColor = '#e60a18';
  var markerRow = 19;
  var preferredSourceRow = 18;
  var firstTargetRow = 20;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getMaxColumns();

  if (lastRow < firstTargetRow) {
    var noRowsMessage = 'No target rows found in QUOT.';
    ss.toast(noRowsMessage, 'Formula Repair', 8);
    Logger.log(noRowsMessage);
    return { status: 'nothing_to_copy', message: noRowsMessage };
  }

  var markerBackgrounds = sheet
    .getRange(markerRow, 1, 1, lastColumn)
    .getBackgrounds()[0];
  var formulasInRow18 = sheet
    .getRange(preferredSourceRow, 1, 1, lastColumn)
    .getFormulas()[0];

  var markerColumns = [];
  var copiedColumns = [];
  var skippedColumns = [];
  var copiedCells = 0;

  for (var column = 1; column <= lastColumn; column++) {
    var background = String(markerBackgrounds[column - 1] || '')
      .trim()
      .toLowerCase();

    // This is the required gate: only an exact row-19 #e60a18 match is used.
    if (background !== markerColor) continue;

    var columnName = columnToLetter_(column);
    markerColumns.push(columnName);
    var sourceFormula = String(formulasInRow18[column - 1] || '').trim();
    if (!sourceFormula) {
      skippedColumns.push({
        column: columnName,
        reason: 'No formula found in row 18'
      });
      continue;
    }

    var sourceCell = sheet.getRange(preferredSourceRow, column);
    for (var row = firstTargetRow; row <= lastRow; row++) {
      // The source is always row 18; row 20 is never used as a source.
      sourceCell.copyTo(
        sheet.getRange(row, column),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
      copiedCells++;
    }
    copiedColumns.push(columnName);
  }

  var status = copiedCells ? 'ok' : 'nothing_to_copy';
  var summary = {
    status: status,
    sheetName: 'QUOT',
    markerColor: markerColor,
    markerRow: markerRow,
    preferredSourceRow: preferredSourceRow,
    firstTargetRow: firstTargetRow,
    lastRow: lastRow,
    markerColumns: markerColumns,
    copiedColumns: copiedColumns,
    skippedColumns: skippedColumns,
    copiedCells: copiedCells
  };

  var message = copiedCells
    ? 'Copied ' + copiedCells + ' formulas in ' + copiedColumns.join(', ')
    : 'No formulas copied. Red columns: ' + (markerColumns.join(', ') || 'none');
  ss.toast(message, 'Formula Repair', 10);
  Logger.log(JSON.stringify(summary));
  return summary;
}

/**
 * Converts a 1-based sheet column number to A1 column letters.
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
