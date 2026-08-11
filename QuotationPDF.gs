var PDF_FOLDER_ID = '1RKsuuzsjY5qwPqUAM2KNxuQHylC3i7QP';
var SHEET_NAME = 'PDF FINAL';
var HEADER_LAST_ROW = 25;
var ITEM_FIRST_ROW = 26;
var ITEM_LAST_ROW = 67;
var MIN_VISIBLE_ITEMS = 5;
var QUOTATION_REF_CELL = 'N16';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PDF Export')
    .addItem('Generate PDF', 'generateQuotationPDF')
    .addToUi();
}

function generateQuotationPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  var displayValues = fullRange.getDisplayValues();
  var backgrounds = fullRange.getBackgrounds();
  var fontColors = fullRange.getFontColors();
  var fontSizes = fullRange.getFontSizes();
  var fontWeights = fullRange.getFontWeights();
  var fontFamilies = fullRange.getFontFamilies();
  var fontStyles = fullRange.getFontStyles();
  var hAligns = fullRange.getHorizontalAlignments();
  var vAligns = fullRange.getVerticalAlignments();
  var merges = fullRange.getMergedRanges();

  var colWidths = [];
  for (var c = 1; c <= lastCol; c++) colWidths.push(sheet.getColumnWidth(c));

  var rowHeights = [];
  for (var r = 1; r <= lastRow; r++) rowHeights.push(sheet.getRowHeight(r));

  var mergeMap = buildMergeMap_(merges);

  // Find last filled item row - check column B onwards (skip Sr No column A)
  var lastFilledItem = ITEM_FIRST_ROW - 1;
  for (var r = ITEM_FIRST_ROW - 1; r < ITEM_LAST_ROW; r++) {
    var hasData = false;
    for (var c = 1; c < displayValues[r].length; c++) {
      if (displayValues[r][c] && String(displayValues[r][c]).trim() !== '') {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastFilledItem = r + 1;
    }
  }
  var visibleItemEnd = Math.max(lastFilledItem, ITEM_FIRST_ROW + MIN_VISIBLE_ITEMS - 1);
  visibleItemEnd = Math.min(visibleItemEnd, ITEM_LAST_ROW);
  Logger.log('Last filled item row: ' + lastFilledItem);
  Logger.log('Visible items: row ' + ITEM_FIRST_ROW + ' to ' + visibleItemEnd);
  Logger.log('Rows being skipped: ' + (visibleItemEnd + 1) + ' to ' + ITEM_LAST_ROW);

  // Find TECHNICAL TERMS row
  var termsRow = lastRow;
  for (var r = 0; r < displayValues.length; r++) {
    if (displayValues[r].join(' ').toUpperCase().indexOf('TECHNICAL TERMS') !== -1) {
      termsRow = r + 1;
      break;
    }
  }
  Logger.log('TECHNICAL TERMS at row: ' + termsRow);

  // Build visible rows list - SKIP blank item rows
  var visibleRows = [];
  for (var r = 1; r <= lastRow; r++) {
    // Skip rows between visibleItemEnd+1 and ITEM_LAST_ROW
    if (r > visibleItemEnd && r <= ITEM_LAST_ROW) {
      continue;
    }
    visibleRows.push(r);
  }
  Logger.log('Total visible rows: ' + visibleRows.length + ' out of ' + lastRow);

  // Check overflow
  var totalColWidth = 0;
  for (var i = 0; i < colWidths.length; i++) totalColWidth += colWidths[i];
  var scale = 710 / totalColWidth;

  var headerPx = 0;
  for (var r = 1; r <= HEADER_LAST_ROW; r++) headerPx += rowHeights[r - 1] * scale;
  var itemsPx = 0;
  for (var r = ITEM_FIRST_ROW; r <= visibleItemEnd; r++) itemsPx += rowHeights[r - 1] * scale;
  var middlePx = 0;
  for (var r = ITEM_LAST_ROW + 1; r < termsRow; r++) middlePx += rowHeights[r - 1] * scale;

  var itemsOverflow = (headerPx + itemsPx + middlePx) > 1000;
  Logger.log('Overflow: ' + itemsOverflow + ' (header=' + Math.round(headerPx) + ' items=' + Math.round(itemsPx) + ' middle=' + Math.round(middlePx) + ')');

  // Get images as base64
  var imageData = getImagesBase64_(sheet);

  // Build HTML
  var html = buildHtml_(displayValues, colWidths, rowHeights, backgrounds, fontColors, fontSizes, fontWeights, fontFamilies, fontStyles, hAligns, vAligns, mergeMap, visibleRows, termsRow, itemsOverflow, scale, imageData, lastCol);

  // Convert to PDF
  var pdfBlob = Utilities.newBlob(html, 'text/html', 'q.html').getAs('application/pdf');
  var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var fileName = buildFileName_(sheet);
  var file = folder.createFile(pdfBlob.setName(fileName));
  SpreadsheetApp.getActiveSpreadsheet().toast('PDF done! ' + fileName, 'Done', 10);
  Logger.log('PDF: ' + file.getUrl());
  return file.getUrl();
}

function buildHtml_(data, colWidths, rowHeights, bgs, fcs, fss, fws, ffs, fsts, hAs, vAs, mergeMap, visibleRows, termsRow, itemsOverflow, scale, imageData, lastCol) {
  var css = '@page{size:A4 portrait;margin:10mm 8mm 10mm 8mm;}';
  css += '*{box-sizing:border-box;}';
  css += 'body{margin:0;padding:0;font-family:Roboto,Arial,sans-serif;font-size:9pt;}';
  css += 'table{border-collapse:collapse;width:100%;table-layout:fixed;}';
  css += 'td,th{padding:1px 3px;border:0.5px solid #999;overflow:hidden;word-wrap:break-word;}';
  if (itemsOverflow) {
    css += 'thead{display:table-header-group;}';
  } else {
    css += 'thead{display:table-row-group;}';
  }
  css += '.terms-section{page-break-inside:avoid;}';
  css += '.terms-section table{border-collapse:collapse;width:100%;table-layout:fixed;}';

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>';

  var colgroup = '<colgroup>';
  for (var c = 0; c < lastCol; c++) {
    colgroup += '<col style="width:' + Math.round(colWidths[c] * scale) + 'px;">';
  }
  colgroup += '</colgroup>';

  html += '<table>' + colgroup;

  // THEAD rows 1-25
  html += '<thead>';
  for (var i = 0; i < visibleRows.length; i++) {
    var r = visibleRows[i];
    if (r > HEADER_LAST_ROW) break;
    html += makeRow_(r, data, rowHeights, bgs, fcs, fss, fws, ffs, fsts, hAs, vAs, mergeMap, scale, imageData, lastCol);
  }
  html += '</thead>';

  // TBODY - items and middle (before terms)
  html += '<tbody>';
  for (var i = 0; i < visibleRows.length; i++) {
    var r = visibleRows[i];
    if (r <= HEADER_LAST_ROW) continue;
    if (r >= termsRow) break;
    html += makeRow_(r, data, rowHeights, bgs, fcs, fss, fws, ffs, fsts, hAs, vAs, mergeMap, scale, imageData, lastCol);
  }
  html += '</tbody></table>';

  // TERMS BLOCK - never splits
  html += '<div class="terms-section"><table>' + colgroup + '<tbody>';
  for (var i = 0; i < visibleRows.length; i++) {
    var r = visibleRows[i];
    if (r < termsRow) continue;
    html += makeRow_(r, data, rowHeights, bgs, fcs, fss, fws, ffs, fsts, hAs, vAs, mergeMap, scale, imageData, lastCol);
  }
  html += '</tbody></table></div>';

  html += '</body></html>';
  return html;
}

function makeRow_(r, data, rowHeights, bgs, fcs, fss, fws, ffs, fsts, hAs, vAs, mergeMap, scale, imageData, lastCol) {
  var ri = r - 1;
  var rh = Math.max(Math.round(rowHeights[ri] * scale * 0.75), 12);
  var html = '<tr style="height:' + rh + 'px;">';

  var c = 0;
  while (c < lastCol) {
    var key = r + ',' + (c + 1);
    if (mergeMap.skip[key]) { c++; continue; }

    var colspan = 1, rowspan = 1;
    if (mergeMap.starts[key]) {
      colspan = mergeMap.starts[key].cols;
      rowspan = mergeMap.starts[key].rows;
    }

    var bg = bgs[ri][c];
    var fc = fcs[ri][c];
    var fs = fss[ri][c];
    var fw = fws[ri][c];
    var ff = ffs[ri][c];
    var fst = fsts[ri][c];
    var ha = hAs[ri][c];
    var va = vAs[ri][c];
    var val = data[ri][c] || '';

    // Check for image
    var imgTag = '';
    for (var x = 0; x < imageData.length; x++) {
      if (imageData[x].row === r && imageData[x].col === (c + 1)) {
        imgTag = '<img src="' + imageData[x].data + '" style="max-height:' + Math.round(rh * 1.2) + 'px;max-width:100%;">';
        break;
      }
    }

    var s = '';
    if (bg && bg !== '#ffffff') s += 'background:' + bg + ';';
    if (fc) s += 'color:' + fc + ';';
    s += 'font-size:' + Math.max(Math.round(fs * 0.85), 6) + 'pt;';
    if (fw === 'bold') s += 'font-weight:bold;';
    if (fst === 'italic') s += 'font-style:italic;';
    if (ff) s += 'font-family:' + ff + ',sans-serif;';
    s += 'text-align:' + (ha || 'left') + ';';
    s += 'vertical-align:' + (va || 'middle') + ';';

    var td = '<td style="' + s + '"';
    if (colspan > 1) td += ' colspan="' + colspan + '"';
    if (rowspan > 1) td += ' rowspan="' + rowspan + '"';
    td += '>';

    if (imgTag) {
      td += imgTag;
    } else {
      td += esc_(val);
    }
    td += '</td>';
    html += td;
    c += colspan;
  }
  html += '</tr>';
  return html;
}

function buildMergeMap_(merges) {
  var starts = {};
  var skip = {};
  for (var i = 0; i < merges.length; i++) {
    var rng = merges[i];
    var sr = rng.getRow(), sc = rng.getColumn();
    var nr = rng.getNumRows(), nc = rng.getNumColumns();
    starts[sr + ',' + sc] = {rows: nr, cols: nc};
    for (var mr = sr; mr < sr + nr; mr++) {
      for (var mc = sc; mc < sc + nc; mc++) {
        if (mr === sr && mc === sc) continue;
        skip[mr + ',' + mc] = true;
      }
    }
  }
  return {starts: starts, skip: skip};
}

function getImagesBase64_(sheet) {
  var imgs = [];
  try {
    var sheetImages = sheet.getImages();
    for (var i = 0; i < sheetImages.length; i++) {
      var img = sheetImages[i];
      var anchor = img.getAnchorCell();
      var blob = img.getBlob();
      var contentType = blob.getContentType();
      var b64 = Utilities.base64Encode(blob.getBytes());
      var dataUri = 'data:' + contentType + ';base64,' + b64;
      imgs.push({
        row: anchor.getRow(),
        col: anchor.getColumn(),
        data: dataUri
      });
    }
  } catch(e) {
    Logger.log('Image error: ' + e.message);
  }
  return imgs;
}

function esc_(s) {
  var str = String(s);
  str = str.replace(/&/g, String.fromCharCode(38) + 'amp;');
  str = str.replace(/</g, String.fromCharCode(38) + 'lt;');
  str = str.replace(/>/g, String.fromCharCode(38) + 'gt;');
  str = str.replace(/"/g, String.fromCharCode(38) + 'quot;');
  str = str.replace(/\n/g, '<br>');
  return str;
}

function buildFileName_(sheet) {
  var ref = '';
  try { ref = String(sheet.getRange(QUOTATION_REF_CELL).getValue() || '').trim(); } catch(e){}
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy_HHmm');
  var name = ref ? ref.replace(/[\\/:*?"<>|]/g, '-') : 'Quotation';
  return name + '_' + stamp + '.pdf';
}
