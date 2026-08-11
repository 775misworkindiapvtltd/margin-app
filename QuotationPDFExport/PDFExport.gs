/** ============================================================
 *  QUOTATION PDF EXPORT SCRIPT
 *  Paste this whole file into Extensions > Apps Script
 *  (Multi Metals India - "Copy of Sales Quotation Format")
 * ============================================================ */

/** ============================================================
 *  CONFIGURATION -- EDIT THESE VALUES BEFORE RUNNING
 * ============================================================ */

// 1) Drive folder where the generated PDF will be saved.
//    Open your Drive folder in the browser and copy the ID from the URL:
//    https://drive.google.com/drive/folders/<THIS_PART_IS_THE_FOLDER_ID>
var PDF_FOLDER_ID = 'PASTE_YOUR_DRIVE_FOLDER_ID_HERE';

// 2) Tab/sheet name to export.
var SHEET_NAME = 'PDF FINAL';

// 3) Header block (customer/quotation info at the top).
//    It ONLY repeats on later pages if the item rows themselves spill onto
//    those pages. If the item table fits on page 1 and only the
//    Technical/Commercial Terms block gets pushed to page 2, the header will
//    NOT be repeated there.
var HEADER_LAST_ROW = 25;

// 4) Item entry block -- blank rows inside this range get hidden.
var ITEM_FIRST_ROW = 26;
var ITEM_LAST_ROW = 67;
var ITEM_CHECK_COLUMN = 'D'; // column used to test "does this row have data" (Description of Goods)
var MIN_VISIBLE_ITEM_ROWS = 5; // always keep at least this many item rows visible

// 5) Terms block that must NEVER be split across a page break.
//    This block starts at the "TECHNICAL TERMS & CONDITIONS" heading and
//    runs all the way to the end of the sheet (Commercial Terms,
//    Declaration, "Thanks for your business" banner, etc.). If it doesn't
//    fully fit in the remaining space of the current page, the WHOLE block
//    is pushed to a fresh page instead of being cut mid-way.
//    The row is auto-detected by searching for TNC_HEADING_TEXT, so it
//    keeps working even if rows shift after future edits.
var TNC_HEADING_TEXT = 'TECHNICAL TERMS & CONDITIONS';
var FOOTER_FIRST_ROW_FALLBACK = 77; // used only if TNC_HEADING_TEXT can't be found

// 6) Page-break calibration (used by both the header-repeat decision and
//    the terms-block-no-split logic).
//    These describe, in the same pixel scale Sheets uses for row/column
//    sizes, how tall a printed page is relative to its width when
//    "Fit to width" is applied. Defaults below are a standard estimate
//    for A4 portrait with default margins.
//    -> Generate a PDF once; if pagination looks off, run
//       debugPageBreakEstimate() and nudge PAGE_HEIGHT_BASE_PX up/down.
var PAGE_WIDTH_BASE_PX = 800;
var PAGE_HEIGHT_BASE_PX = 1050;

// 7) Cell holding the Quotation Ref No, used to name the PDF file.
//    Adjust to match your actual sheet if this cell reference is wrong.
var QUOTATION_REF_CELL = 'N16';

/** ============================================================
 *  MENU (optional convenience)
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quotation Tools')
    .addItem('Generate Quotation PDF', 'generateQuotationPDF')
    .addToUi();
}

/** ============================================================
 *  MAIN ENTRY POINT
 * ============================================================ */
function generateQuotationPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found. Check the SHEET_NAME config.');
  }

  var termsFirstRow = findTermsHeadingRow_(sheet);
  var spacerRow = termsFirstRow - 1;
  var originalSpacerHeight = sheet.getRowHeight(spacerRow);
  var originalFrozenRows = sheet.getFrozenRows();

  try {
    var visibility = adjustItemRowVisibility_(sheet);
    console.log('Item rows -> filled: ' + visibility.filledCount +
      ', visible: ' + visibility.visibleCount +
      ', hidden: ' + visibility.hiddenCount);

    // Decide whether the header (rows 1-25) needs to repeat: only when the
    // item table itself overflows past page 1. Otherwise leave it unfrozen
    // for this export so it does NOT reappear on the page the Terms block
    // gets pushed to.
    var itemsOverflow = doItemsOverflowFirstPage_(sheet, termsFirstRow);
    console.log('Items overflow to next page: ' + itemsOverflow);
    setHeaderRepeat_(sheet, itemsOverflow);

    adjustTermsPageBreak_(sheet, termsFirstRow, spacerRow, originalSpacerHeight);

    SpreadsheetApp.flush();

    var pdfBlob = exportSheetAsPdf_(ss, sheet);

    var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
    var fileName = buildFileName_(sheet);
    var file = folder.createFile(pdfBlob.setName(fileName));

    var pdfUrl = file.getUrl();

    // Requested: print the PDF URL to the console/log.
    console.log('PDF URL: ' + pdfUrl);
    Logger.log('PDF URL: ' + pdfUrl);

    SpreadsheetApp.getActiveSpreadsheet().toast('PDF ready: ' + fileName, 'Quotation PDF', 8);
    return pdfUrl;
  } finally {
    // Always restore the sheet's original state so this export doesn't
    // leave the normal editing view altered.
    sheet.setRowHeight(spacerRow, originalSpacerHeight);
    sheet.setFrozenRows(originalFrozenRows);
    SpreadsheetApp.flush();
  }
}

/** ============================================================
 *  1) ITEM ROW VISIBILITY (rows 26-67)
 * ============================================================ */
function adjustItemRowVisibility_(sheet) {
  var totalRows = ITEM_LAST_ROW - ITEM_FIRST_ROW + 1;

  // Reset to a known state first so hidden state doesn't affect the read.
  sheet.showRows(ITEM_FIRST_ROW, totalRows);

  var range = sheet.getRange(ITEM_CHECK_COLUMN + ITEM_FIRST_ROW + ':' + ITEM_CHECK_COLUMN + ITEM_LAST_ROW);
  var values = range.getValues();

  var visibility = computeItemVisibility_(values.map(function (row) { return row[0]; }), MIN_VISIBLE_ITEM_ROWS);

  if (visibility.hiddenCount > 0) {
    sheet.hideRows(ITEM_FIRST_ROW + visibility.visibleCount, visibility.hiddenCount);
  }

  return visibility;
}

/**
 * Pure logic (no Sheets API calls) so it can be unit tested independently.
 * cellValues: array of raw cell values for ITEM_FIRST_ROW..ITEM_LAST_ROW (in order)
 * minVisible: minimum rows to keep visible even if fewer/no rows have data
 */
function computeItemVisibility_(cellValues, minVisible) {
  var totalRows = cellValues.length;

  var lastFilledOffset = -1;
  for (var i = 0; i < totalRows; i++) {
    var v = cellValues[i];
    if (v !== '' && v !== null && v !== undefined && String(v).trim() !== '') {
      lastFilledOffset = i;
    }
  }

  var filledCount = lastFilledOffset + 1;
  var visibleCount = Math.max(filledCount, minVisible);
  visibleCount = Math.min(visibleCount, totalRows);

  var hiddenCount = totalRows - visibleCount;

  return {
    filledCount: filledCount,
    visibleCount: visibleCount,
    hiddenCount: hiddenCount
  };
}

/** ============================================================
 *  2) HEADER ROWS 1-25 -- CONDITIONAL REPEAT
 *     Only repeats on later pages when the item table itself overflows
 *     past page 1. If items fit on page 1 and only the Terms block gets
 *     pushed to the next page, the header must NOT repeat there.
 * ============================================================ */
function doItemsOverflowFirstPage_(sheet, termsFirstRow) {
  var pageHeightPx = getEffectivePageHeightPx_(sheet);
  var headerHeightPx = sumVisibleRowHeights_(sheet, 1, HEADER_LAST_ROW);
  // Everything between the header and the Terms block (items + weight/
  // charges/total rows) -- this is the content whose overflow decides
  // whether the header needs to repeat.
  var contentHeightPx = sumVisibleRowHeights_(sheet, HEADER_LAST_ROW + 1, termsFirstRow - 1);

  var pagesNeeded = computePagesNeeded_(headerHeightPx, contentHeightPx, pageHeightPx);
  return pagesNeeded > 1;
}

/**
 * Pure logic (no Sheets API calls) so it can be unit tested independently.
 * Estimates how many pages the header + content would need if printed
 * back-to-back with no repeated header (i.e. does it overflow page 1 at all).
 */
function computePagesNeeded_(headerHeightPx, contentHeightPx, pageHeightPx) {
  if (pageHeightPx <= 0) {
    return 1;
  }
  var totalHeight = headerHeightPx + contentHeightPx;
  return Math.max(1, Math.ceil(totalHeight / pageHeightPx));
}

function setHeaderRepeat_(sheet, shouldRepeat) {
  var targetFrozenRows = shouldRepeat ? HEADER_LAST_ROW : 0;
  if (sheet.getFrozenRows() !== targetFrozenRows) {
    sheet.setFrozenRows(targetFrozenRows);
  }
}

/** ============================================================
 *  3) TERMS BLOCK (TECHNICAL TERMS & CONDITIONS onward) MUST NOT SPLIT
 * ============================================================ */
function adjustTermsPageBreak_(sheet, termsFirstRow, spacerRow, naturalSpacerHeight) {
  var termsLastRow = sheet.getLastRow();
  var pageHeightPx = getEffectivePageHeightPx_(sheet);
  var headerHeightPx = sumVisibleRowHeights_(sheet, 1, HEADER_LAST_ROW);
  var bodyPerPagePx = pageHeightPx - headerHeightPx;

  if (bodyPerPagePx <= 0) {
    // Header alone would exceed the estimated page height -- bail out
    // rather than apply a nonsensical adjustment.
    return;
  }

  var heightBeforeTerms = sumVisibleRowHeights_(sheet, HEADER_LAST_ROW + 1, spacerRow - 1);
  var termsHeight = sumVisibleRowHeights_(sheet, termsFirstRow, termsLastRow);

  var usedOnCurrentPage = heightBeforeTerms % bodyPerPagePx;
  var spaceLeftOnCurrentPage = bodyPerPagePx - usedOnCurrentPage;

  var extraNeeded = computeSpacerExtraPx_(spaceLeftOnCurrentPage, termsHeight);
  sheet.setRowHeight(spacerRow, naturalSpacerHeight + extraNeeded);
}

/**
 * Pure logic (no Sheets API calls) so it can be unit tested independently.
 * Returns extra pixels to add to the spacer row height so the terms block
 * either fits fully in the remaining space of the current page, or starts
 * cleanly at the top of the next page instead of being cut mid-way.
 */
function computeSpacerExtraPx_(spaceLeftOnCurrentPage, termsHeight) {
  if (termsHeight <= spaceLeftOnCurrentPage) {
    return 0; // already fits on the current page -- no push needed
  }
  var extra = Math.max(0, spaceLeftOnCurrentPage - 1);
  return Math.round(extra);
}

function sumVisibleRowHeights_(sheet, firstRow, lastRow) {
  var total = 0;
  for (var r = firstRow; r <= lastRow; r++) {
    if (!sheet.isRowHiddenByUser(r)) {
      total += sheet.getRowHeight(r);
    }
  }
  return total;
}

function getEffectivePageHeightPx_(sheet) {
  var lastCol = sheet.getLastColumn();
  var totalColWidth = 0;
  for (var c = 1; c <= lastCol; c++) {
    totalColWidth += sheet.getColumnWidth(c);
  }
  return PAGE_HEIGHT_BASE_PX * (totalColWidth / PAGE_WIDTH_BASE_PX);
}

/**
 * Finds the row containing the "TECHNICAL TERMS & CONDITIONS" heading by
 * searching the sheet. Falls back to FOOTER_FIRST_ROW_FALLBACK if not found
 * (e.g. heading text was edited).
 */
function findTermsHeadingRow_(sheet) {
  var finder = sheet.createTextFinder(TNC_HEADING_TEXT).matchCase(false).matchEntireCell(false);
  var found = finder.findNext();
  if (found) {
    return found.getRow();
  }
  console.log('Could not find "' + TNC_HEADING_TEXT + '" -- using fallback row ' + FOOTER_FIRST_ROW_FALLBACK);
  return FOOTER_FIRST_ROW_FALLBACK;
}

/** ============================================================
 *  4) PDF EXPORT
 * ============================================================ */
function exportSheetAsPdf_(ss, sheet) {
  var baseUrl = ss.getUrl().replace(/edit(\?.*)?$/, '');
  var url = baseUrl + 'export'
    + '?format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    + '&top_margin=0.40'
    + '&bottom_margin=0.40'
    + '&left_margin=0.40'
    + '&right_margin=0.40'
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenumbers=false'
    + '&gridlines=false'
    // fzr just tells the export to repeat whatever rows are currently
    // frozen on the sheet. setHeaderRepeat_() has already set frozen rows
    // to either 0 (no repeat) or HEADER_LAST_ROW (repeat), so fzr=true is
    // always safe here -- with 0 frozen rows it's a no-op.
    + '&fzr=true';

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed (' + response.getResponseCode() + '): ' + response.getContentText());
  }

  return response.getBlob().setContentType('application/pdf');
}

function buildFileName_(sheet) {
  var ref = '';
  try {
    ref = String(sheet.getRange(QUOTATION_REF_CELL).getValue() || '').trim();
  } catch (e) {
    // ignore, fall back to default name below
  }
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyyyy_HHmmss');
  var safeRef = ref ? ref.replace(/[\\/:*?"<>|]/g, '-') : 'Quotation';
  return safeRef + '_' + stamp + '.pdf';
}

/** ============================================================
 *  DEBUG HELPER -- run this manually to calibrate PAGE_HEIGHT_BASE_PX
 *  and to see the header-repeat / terms-block decisions before exporting.
 * ============================================================ */
function debugPageBreakEstimate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var termsFirstRow = findTermsHeadingRow_(sheet);
  var pageHeightPx = getEffectivePageHeightPx_(sheet);
  var headerHeightPx = sumVisibleRowHeights_(sheet, 1, HEADER_LAST_ROW);
  var contentHeightPx = sumVisibleRowHeights_(sheet, HEADER_LAST_ROW + 1, termsFirstRow - 1);
  var termsHeightPx = sumVisibleRowHeights_(sheet, termsFirstRow, sheet.getLastRow());

  console.log('Terms block starts at row: ' + termsFirstRow);
  console.log('Estimated page height (px): ' + pageHeightPx);
  console.log('Header height (px): ' + headerHeightPx);
  console.log('Items+charges content height (px): ' + contentHeightPx);
  console.log('Terms block height (px): ' + termsHeightPx);
  console.log('Items overflow to next page: ' + doItemsOverflowFirstPage_(sheet, termsFirstRow));
}
