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

// 3) Header block -- always visible, repeats on every PDF page.
var HEADER_LAST_ROW = 25;

// 4) Item entry block -- blank rows inside this range get hidden.
var ITEM_FIRST_ROW = 26;
var ITEM_LAST_ROW = 67;
var ITEM_CHECK_COLUMN = 'D'; // column used to test "does this row have data" (Description of Goods)
var MIN_VISIBLE_ITEM_ROWS = 5; // always keep at least this many item rows visible

// 5) Footer block (weight / charges / total / T&C / thanks banner)
//    -- kept intact, never split across a page break.
var FOOTER_FIRST_ROW = 77;
var FOOTER_LAST_ROW = 0; // 0 = auto-detect using sheet.getLastRow()

// 6) Page-break calibration (used only by the footer-no-split logic).
//    These describe, in the same pixel scale Sheets uses for row/column
//    sizes, how tall a printed page is relative to its width when
//    "Fit to width" is applied. Defaults below are a standard estimate
//    for A4 portrait with default margins.
//    -> Generate a PDF once, check where row 77's block actually breaks,
//       and nudge PAGE_HEIGHT_BASE_PX up/down if it is off.
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

  var spacerRow = FOOTER_FIRST_ROW - 1;
  var originalSpacerHeight = sheet.getRowHeight(spacerRow);

  try {
    var visibility = adjustItemRowVisibility_(sheet);
    console.log('Item rows -> filled: ' + visibility.filledCount +
      ', visible: ' + visibility.visibleCount +
      ', hidden: ' + visibility.hiddenCount);

    ensureHeaderRepeats_(sheet);
    adjustFooterPageBreak_(sheet, spacerRow, originalSpacerHeight);

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
    // Always restore the spacer row so the sheet's normal layout isn't left altered.
    sheet.setRowHeight(spacerRow, originalSpacerHeight);
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
 *  2) HEADER ROWS 1-25 REPEAT ON EVERY PAGE
 * ============================================================ */
function ensureHeaderRepeats_(sheet) {
  if (sheet.getFrozenRows() < HEADER_LAST_ROW) {
    sheet.setFrozenRows(HEADER_LAST_ROW);
  }
}

/** ============================================================
 *  3) FOOTER BLOCK (row 77 onward) MUST NOT BE CUT MID-PAGE
 * ============================================================ */
function adjustFooterPageBreak_(sheet, spacerRow, naturalSpacerHeight) {
  var footerLastRow = resolveFooterLastRow_(sheet);
  var pageHeightPx = getEffectivePageHeightPx_(sheet);
  var headerHeightPx = sumVisibleRowHeights_(sheet, 1, HEADER_LAST_ROW);
  var bodyPerPagePx = pageHeightPx - headerHeightPx;

  if (bodyPerPagePx <= 0) {
    // Header alone would exceed the estimated page height -- bail out
    // rather than apply a nonsensical adjustment.
    return;
  }

  var heightBeforeFooter = sumVisibleRowHeights_(sheet, HEADER_LAST_ROW + 1, spacerRow - 1);
  var footerHeight = sumVisibleRowHeights_(sheet, FOOTER_FIRST_ROW, footerLastRow);

  var usedOnCurrentPage = heightBeforeFooter % bodyPerPagePx;
  var spaceLeftOnCurrentPage = bodyPerPagePx - usedOnCurrentPage;

  var extraNeeded = computeSpacerExtraPx_(spaceLeftOnCurrentPage, footerHeight);
  sheet.setRowHeight(spacerRow, naturalSpacerHeight + extraNeeded);
}

/**
 * Pure logic (no Sheets API calls) so it can be unit tested independently.
 * Returns extra pixels to add to the spacer row height so the footer block
 * either fits fully in the remaining space of the current page, or starts
 * cleanly at the top of the next page instead of being cut mid-way.
 */
function computeSpacerExtraPx_(spaceLeftOnCurrentPage, footerHeight) {
  if (footerHeight <= spaceLeftOnCurrentPage) {
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

function resolveFooterLastRow_(sheet) {
  return FOOTER_LAST_ROW > 0 ? FOOTER_LAST_ROW : sheet.getLastRow();
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
    + '&fzr=true'; // repeat frozen rows (1-25) on every page

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
 * ============================================================ */
function debugPageBreakEstimate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var pageHeightPx = getEffectivePageHeightPx_(sheet);
  var headerHeightPx = sumVisibleRowHeights_(sheet, 1, HEADER_LAST_ROW);
  console.log('Estimated page height (px): ' + pageHeightPx);
  console.log('Header height (px): ' + headerHeightPx);
  console.log('Body capacity per page (px): ' + (pageHeightPx - headerHeightPx));
}
