/**
 * Finds #e60a18 in row 19 of the QUOT sheet and copies formulas down.
 *
 * Row 18 is the preferred formula template. If row 18 is empty in a marked
 * column, the existing formula in row 20 is used as the fallback template.
 * Formulas are copied into row 20 through the last used row.
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
  var formulasInRow20 = sheet
    .getRange(firstTargetRow, 1, 1, lastColumn)
    .getFormulas()[0];

  var markerColumns = [];
  var copiedColumns = [];
  var skippedColumns = [];
  var copiedCells = 0;

  for (var column = 1; column <= lastColumn; column++) {
    var background = String(markerBackgrounds[column - 1] || '')
      .trim()
      .toLowerCase();
    if (background !== markerColor) continue;

    var columnName = columnToLetter_(column);
    markerColumns.push(columnName);

    var sourceRow = formulasInRow18[column - 1]
      ? preferredSourceRow
      : (formulasInRow20[column - 1] ? firstTargetRow : 0);
    if (!sourceRow) {
      skippedColumns.push({
        column: columnName,
        reason: 'No formula in row 18 or row 20'
      });
      continue;
    }

    var sourceCell = sheet.getRange(sourceRow, column);
    for (var row = firstTargetRow; row <= lastRow; row++) {
      // PASTE_FORMULA adjusts relative row references for each destination row.
      sourceCell.copyTo(
        sheet.getRange(row, column),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
      copiedCells++;
    }
    copiedColumns.push(columnName + ' (from row ' + sourceRow + ')');
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
