var PDF_FOLDER_ID = '1RKsuuzsjY5qwPqUAM2KNxuQHylC3i7QP';
var SHEET_NAME = 'PDF FINAL';
var HEADER_LAST_ROW = 25;
var ITEM_FIRST_ROW = 26;
var ITEM_LAST_ROW = 67;
var MIN_VISIBLE_ITEMS = 5;
var QUOTATION_REF_CELL = 'N16';

// These are layout units used only for planning the HTML pages. The output is
// still rendered as HTML and converted with Utilities.getAs('application/pdf').
var HTML_PAGE_HEIGHT = 1000;
var CONTINUATION_NOTE_HEIGHT = 14;
var HTML_WIDTH = 710;
var HTML_ROW_SCALE = 0.75;

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
  var mergeMap = buildMergeMap_(fullRange.getMergedRanges());

  var colWidths = [];
  var rowHeights = [];
  var totalColWidth = 0;
  var c;
  var r;
  for (c = 1; c <= lastCol; c++) {
    colWidths.push(sheet.getColumnWidth(c));
    totalColWidth += sheet.getColumnWidth(c);
  }
  for (r = 1; r <= lastRow; r++) rowHeights.push(sheet.getRowHeight(r));

  var lastFilledItem = findLastFilledItem_(displayValues, lastCol);
  var visibleItemEnd = Math.max(
    lastFilledItem,
    ITEM_FIRST_ROW + MIN_VISIBLE_ITEMS - 1
  );
  visibleItemEnd = Math.min(visibleItemEnd, ITEM_LAST_ROW, lastRow);

  var termsRow = findTermsRow_(displayValues);
  var termsExists = termsRow <= lastRow;
  var middleFirst = Math.min(ITEM_LAST_ROW + 1, lastRow + 1);
  var middleLast = termsExists ? termsRow - 1 : lastRow;

  var scale = HTML_WIDTH / Math.max(totalColWidth, 1);
  var imageData = getImagesBase64_(sheet);
  var plan = buildPagePlan_(
    rowHeights,
    mergeMap,
    visibleItemEnd,
    middleFirst,
    middleLast,
    termsExists ? termsRow : lastRow + 1,
    lastRow,
    scale
  );

  Logger.log('Last filled item row: ' + lastFilledItem);
  Logger.log('Visible item rows: ' + (visibleItemEnd - ITEM_FIRST_ROW + 1));
  Logger.log('TECHNICAL TERMS row: ' + (termsExists ? termsRow : 'not found'));
  Logger.log('Item pages: ' + plan.itemPageCount);
  Logger.log('Item overflow: ' + plan.itemOverflow);
  Logger.log('Page plan: ' + JSON.stringify(plan.summary));

  var html = buildHtml_(
    displayValues,
    colWidths,
    rowHeights,
    backgrounds,
    fontColors,
    fontSizes,
    fontWeights,
    fontFamilies,
    fontStyles,
    hAligns,
    vAligns,
    mergeMap,
    plan.pages,
    scale,
    imageData,
    lastCol
  );

  var pdfBlob = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf');
  var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var fileName = buildFileName_(sheet);
  var file = folder.createFile(pdfBlob.setName(fileName));
  ss.toast('PDF done! ' + fileName, 'Done', 10);
  Logger.log('PDF: ' + file.getUrl());
  return file.getUrl();
}

function findLastFilledItem_(data, lastCol) {
  var last = ITEM_FIRST_ROW - 1;
  var end = Math.min(ITEM_LAST_ROW, data.length);
  var r;
  var c;
  for (r = ITEM_FIRST_ROW; r <= end; r++) {
    for (c = 2; c <= lastCol; c++) {
      if (String(data[r - 1][c - 1] || '').trim() !== '') {
        last = r;
        break;
      }
    }
  }
  return last;
}

function findTermsRow_(data) {
  var needle = 'TECHNICAL TERMS';
  var r;
  for (r = 0; r < data.length; r++) {
    if (data[r].join(' ').toUpperCase().indexOf(needle) !== -1) return r + 1;
  }
  return data.length + 1;
}

function renderedRowHeight_(rowHeights, row, scale) {
  var sourceHeight = rowHeights[row - 1] || 21;
  return Math.max(Math.round(sourceHeight * scale * HTML_ROW_SCALE), 12);
}

function rowsForRange_(first, last, rowHeights, mergeMap, scale) {
  var units = [];
  var r = first;
  while (r <= last) {
    var group = mergeMap.rowGroups[r];
    var unitEnd = group && group.start === r ? Math.min(group.end, last) : r;
    var rows = [];
    var height = 0;
    var x;
    for (x = r; x <= unitEnd; x++) {
      rows.push(x);
      height += renderedRowHeight_(rowHeights, x, scale);
    }
    units.push({rows: rows, height: height});
    r = unitEnd + 1;
  }
  return units;
}

function chunkUnits_(units, capacity) {
  var chunks = [];
  var current = null;
  var i;
  var unit;

  for (i = 0; i < units.length; i++) {
    unit = units[i];
    if (!current) {
      current = {units: [unit], height: unit.height};
    } else if (current.height + unit.height <= capacity) {
      current.units.push(unit);
      current.height += unit.height;
    } else {
      chunks.push(current);
      current = {units: [unit], height: unit.height};
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

function flattenUnits_(units) {
  var rows = [];
  var i;
  var j;
  for (i = 0; i < units.length; i++) {
    for (j = 0; j < units[i].rows.length; j++) rows.push(units[i].rows[j]);
  }
  return rows;
}

function addUnitsToPage_(page, units) {
  var rows = flattenUnits_(units);
  page.rows = page.rows.concat(rows);
  for (var i = 0; i < units.length; i++) page.height += units[i].height;
}

function buildPagePlan_(rowHeights, mergeMap, visibleItemEnd, middleFirst,
  middleLast, termsFirst, lastRow, scale) {
  var headerUnits = rowsForRange_(1, HEADER_LAST_ROW, rowHeights, mergeMap, scale);
  var itemUnits = visibleItemEnd >= ITEM_FIRST_ROW
    ? rowsForRange_(ITEM_FIRST_ROW, visibleItemEnd, rowHeights, mergeMap, scale)
    : [];
  var middleUnits = middleFirst <= middleLast
    ? rowsForRange_(middleFirst, middleLast, rowHeights, mergeMap, scale)
    : [];
  var termsUnits = termsFirst <= lastRow
    ? rowsForRange_(termsFirst, lastRow, rowHeights, mergeMap, scale)
    : [];

  var headerHeight = sumUnitHeight_(headerUnits);
  var itemCapacity = Math.max(1, HTML_PAGE_HEIGHT - headerHeight);
  var itemChunks = chunkUnits_(itemUnits, itemCapacity);
  var pages = [];
  var i;

  for (i = 0; i < itemChunks.length; i++) {
    pages.push({
      rows: flattenUnits_(headerUnits).concat(flattenUnits_(itemChunks[i].units)),
      height: headerHeight + itemChunks[i].height,
      header: true,
      kind: 'items',
      continuation: false
    });
  }

  // The item section always has at least MIN_VISIBLE_ITEMS in normal use.
  // Keep this fallback so a malformed/empty template still exports safely.
  if (!pages.length) {
    pages.push({
      rows: flattenUnits_(headerUnits),
      height: headerHeight,
      header: true,
      kind: 'items',
      continuation: false
    });
  }

  var finalPage = pages[pages.length - 1];
  var allMiddleHeight = sumUnitHeight_(middleUnits);
  var allTermsHeight = sumUnitHeight_(termsUnits);
  var canKeepEverything = finalPage.height + allMiddleHeight + allTermsHeight
    <= HTML_PAGE_HEIGHT;

  if (canKeepEverything) {
    addUnitsToPage_(finalPage, middleUnits);
    addUnitsToPage_(finalPage, termsUnits);
  } else {
    // Reserve the small continuation line whenever another page is required.
    // This keeps the Total Price row on page one whenever it really fits.
    var middleCapacity = Math.max(
      1,
      HTML_PAGE_HEIGHT - finalPage.height - CONTINUATION_NOTE_HEIGHT
    );
    var middleChunks = chunkUnits_(middleUnits, middleCapacity);

    if (middleChunks.length) {
      addUnitsToPage_(finalPage, middleChunks[0].units);
      for (i = 1; i < middleChunks.length; i++) {
        pages.push({
          rows: flattenUnits_(middleChunks[i].units),
          height: middleChunks[i].height,
          header: false,
          kind: 'charges',
          continuation: false
        });
      }
    }

    if (termsUnits.length) {
      var termsCapacity = HTML_PAGE_HEIGHT - CONTINUATION_NOTE_HEIGHT;
      var remainingTerms = termsUnits;
      var termsChunks = chunkUnits_(remainingTerms, Math.max(1, termsCapacity));
      for (i = 0; i < termsChunks.length; i++) {
        pages.push({
          rows: flattenUnits_(termsChunks[i].units),
          height: termsChunks[i].height,
          header: false,
          kind: 'terms',
          continuation: false
        });
      }
    }
  }

  // A page repeats rows 1–25 only when it is an item page. Every earlier page
  // gets the small right-aligned continuation label when another page follows.
  for (i = 0; i < pages.length; i++) {
    pages[i].continuation = i < pages.length - 1;
  }

  return {
    pages: pages,
    itemPageCount: itemChunks.length || 1,
    itemOverflow: itemChunks.length > 1,
    summary: pages.map(function (page) {
      return {
        kind: page.kind,
        header: page.header,
        rows: page.rows.length,
        height: page.height,
        continuation: page.continuation
      };
    })
  };
}

function sumUnitHeight_(units) {
  var total = 0;
  for (var i = 0; i < units.length; i++) total += units[i].height;
  return total;
}

function buildHtml_(data, colWidths, rowHeights, bgs, fcs, fss, fws,
  ffs, fsts, hAs, vAs, mergeMap, pages, scale, imageData, lastCol) {
  var css = '';
  css += '@page{size:A4 portrait;margin:10mm 8mm;}';
  css += '*{box-sizing:border-box;}';
  css += 'html,body{margin:0;padding:0;}';
  css += 'body{font-family:Roboto,Arial,sans-serif;font-size:9pt;}';
  css += '.pdf-page{width:100%;position:relative;page-break-after:always;break-after:page;}';
  css += '.pdf-page:last-child{page-break-after:auto;break-after:auto;}';
  css += '.pdf-table{border-collapse:collapse;width:100%;table-layout:fixed;}';
  css += '.pdf-table tr{page-break-inside:avoid;break-inside:avoid;}';
  css += '.pdf-table td{padding:1px 3px;border:0.5px solid #999;';
  css += 'word-wrap:break-word;white-space:normal;overflow:visible;}';
  css += '.cont-note{text-align:right;font-size:7pt;color:#777;font-style:italic;';
  css += 'height:' + CONTINUATION_NOTE_HEIGHT + 'px;line-height:' +
    CONTINUATION_NOTE_HEIGHT + 'px;}';

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    css + '</style></head><body>';
  var colgroup = '<colgroup>';
  for (var c = 0; c < lastCol; c++) {
    colgroup += '<col style="width:' + Math.round(colWidths[c] * scale) + 'px;">';
  }
  colgroup += '</colgroup>';

  for (var p = 0; p < pages.length; p++) {
    var page = pages[p];
    html += '<section class="pdf-page"><table class="pdf-table">' + colgroup +
      '<tbody>';
    for (var i = 0; i < page.rows.length; i++) {
      html += makeRow_(
        page.rows[i], data, rowHeights, bgs, fcs, fss, fws, ffs, fsts,
        hAs, vAs, mergeMap, scale, imageData, lastCol
      );
    }
    html += '</tbody></table>';
    if (page.continuation) html += '<div class="cont-note">Cont. on next page</div>';
    html += '</section>';
  }
  return html + '</body></html>';
}

function makeRow_(r, data, rowHeights, bgs, fcs, fss, fws, ffs, fsts,
  hAs, vAs, mergeMap, scale, imageData, lastCol) {
  var ri = r - 1;
  var rh = renderedRowHeight_(rowHeights, r, scale);
  var html = '<tr style="height:' + rh + 'px;">';
  var c = 0;

  while (c < lastCol) {
    var key = r + ',' + (c + 1);
    if (mergeMap.skip[key]) {
      c++;
      continue;
    }

    var colspan = 1;
    var rowspan = 1;
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
    var imgTag = '';

    for (var x = 0; x < imageData.length; x++) {
      if (imageData[x].row === r && imageData[x].col === c + 1) {
        imgTag = '<img src="' + imageData[x].data + '" style="max-height:' +
          Math.round(rh * 1.2) + 'px;max-width:100%;">';
        break;
      }
    }

    var style = '';
    if (bg && bg !== '#ffffff') style += 'background:' + bg + ';';
    if (fc) style += 'color:' + fc + ';';
    style += 'font-size:' + Math.max(Math.round(fs * 0.85), 6) + 'pt;';
    if (fw === 'bold') style += 'font-weight:bold;';
    if (fst === 'italic') style += 'font-style:italic;';
    if (ff) style += 'font-family:' + ff + ',sans-serif;';
    style += 'text-align:' + (ha || 'left') + ';';
    style += 'vertical-align:' + (va || 'middle') + ';';

    var td = '<td style="' + style + '"';
    if (colspan > 1) td += ' colspan="' + colspan + '"';
    if (rowspan > 1) td += ' rowspan="' + rowspan + '"';
    td += '>' + (imgTag || esc_(val)) + '</td>';
    html += td;
    c += colspan;
  }
  return html + '</tr>';
}

function buildMergeMap_(merges) {
  var starts = {};
  var skip = {};
  var rowGroups = {};
  for (var i = 0; i < merges.length; i++) {
    var range = merges[i];
    var sr = range.getRow();
    var sc = range.getColumn();
    var nr = range.getNumRows();
    var nc = range.getNumColumns();
    starts[sr + ',' + sc] = {rows: nr, cols: nc};
    for (var mr = sr; mr < sr + nr; mr++) {
      for (var mc = sc; mc < sc + nc; mc++) {
        if (mr === sr && mc === sc) continue;
        skip[mr + ',' + mc] = true;
      }
      rowGroups[mr] = {start: sr, end: sr + nr - 1};
    }
  }
  return {starts: starts, skip: skip, rowGroups: rowGroups};
}

function getImagesBase64_(sheet) {
  var imgs = [];
  try {
    var sheetImages = sheet.getImages();
    for (var i = 0; i < sheetImages.length; i++) {
      var image = sheetImages[i];
      var anchor = image.getAnchorCell();
      var blob = image.getBlob();
      imgs.push({
        row: anchor.getRow(),
        col: anchor.getColumn(),
        data: 'data:' + blob.getContentType() + ';base64,' +
          Utilities.base64Encode(blob.getBytes())
      });
    }
  } catch (e) {
    Logger.log('Image error: ' + e.message);
  }
  return imgs;
}

function esc_(value) {
  var str = String(value);
  str = str.replace(/&/g, String.fromCharCode(38) + 'amp;');
  str = str.replace(/</g, String.fromCharCode(38) + 'lt;');
  str = str.replace(/>/g, String.fromCharCode(38) + 'gt;');
  str = str.replace(/"/g, String.fromCharCode(38) + 'quot;');
  return str.replace(/\n/g, '<br>');
}

function buildFileName_(sheet) {
  var ref = '';
  try {
    ref = String(sheet.getRange(QUOTATION_REF_CELL).getValue() || '').trim();
  } catch (e) {}
  var stamp = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy_HHmm'
  );
  var name = ref ? ref.replace(/[\\/:*?"<>|]/g, '-') : 'Quotation';
  return name + '_' + stamp + '.pdf';
}
