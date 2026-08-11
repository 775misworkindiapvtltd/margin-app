/*
 * NATIVE GOOGLE SHEETS QUOTATION PDF EXPORT
 *
 * This version does not use HTML-to-PDF. It makes a temporary copy of the
 * spreadsheet, creates one temporary sheet tab per PDF page, and exports the
 * visible page tabs as a PDF.
 *
 * Copy this complete file into one Apps Script file named QuotationPDF.gs.
 */

var QPDF_FOLDER_ID = '1RKsuuzsjY5qwPqUAM2KNxuQHylC3i7QP';
var QPDF_SHEET_NAME = 'PDF FINAL';
var QPDF_HEADER_LAST = 25;
var QPDF_ITEM_FIRST = 26;
var QPDF_ITEM_LAST = 67;
var QPDF_MIN_ITEMS = 5;
var QPDF_TERMS_TEXT = 'TECHNICAL TERMS & CONDITIONS';
var QPDF_TERMS_FALLBACK = 77;
var QPDF_TOTAL_PRICE_TEXT = 'TOTAL PRICE IN INR';
var QPDF_REF_CELL = 'N16';

// These units are used to plan the temporary Sheet tabs. They match the
// A4 export dimensions used below and are based on the source sheet width.
var QPDF_PAGE_WIDTH = 800;
var QPDF_PAGE_HEIGHT = 1050;
var QPDF_CONTINUATION_HEIGHT = 14;
var QPDF_MAX_COMPRESSION_OVERFLOW = 1.08;
var QPDF_MIN_ROW_HEIGHT = 8;
// Use the available blank space on page 1 for about four more product rows.
// Later item pages keep the normal capacity so the continuation remains stable.
var QPDF_EXTRA_FIRST_PAGE_ROWS = 4;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PDF Export')
    .addItem('Generate PDF', 'generateQuotationPDF')
    .addToUi();
}

function generateQuotationPDF() {
  var sourceBook = SpreadsheetApp.getActiveSpreadsheet();
  var source = sourceBook.getSheetByName(QPDF_SHEET_NAME);
  if (!source) throw new Error('Sheet not found: ' + QPDF_SHEET_NAME);

  var outFolder = DriveApp.getFolderById(QPDF_FOLDER_ID);
  var tempFile = DriveApp.getFileById(sourceBook.getId()).makeCopy(
    '__QPDF_TEMP_' + Date.now(),
    outFolder
  );
  var tempBook = SpreadsheetApp.openById(tempFile.getId());

  try {
    var template = tempBook.getSheetByName(QPDF_SHEET_NAME);
    if (!template) throw new Error('Copied sheet not found: ' + QPDF_SHEET_NAME);

    template.showRows(1, template.getMaxRows());
    SpreadsheetApp.flush();

    // Remove unused product rows from the temporary copy only. This removes
    // blank item space without changing the original quotation sheet.
    var lastItem = qpdfLastItem_(template);
    var visibleEnd = Math.max(
      lastItem,
      QPDF_ITEM_FIRST + QPDF_MIN_ITEMS - 1
    );
    visibleEnd = Math.min(visibleEnd, QPDF_ITEM_LAST);

    var deleteCount = QPDF_ITEM_LAST - visibleEnd;
    if (deleteCount > 0) {
      template.deleteRows(visibleEnd + 1, deleteCount);
    }
    SpreadsheetApp.flush();

    // Rows after the product block shift up after blank rows are deleted.
    var termsRow = qpdfTermsRow_(template);
    var lastRow = template.getLastRow();
    var totalPriceRow = qpdfFindRowContaining_(
      template,
      QPDF_TOTAL_PRICE_TEXT,
      visibleEnd + 1,
      Math.max(visibleEnd, termsRow - 1)
    );

    var plan = qpdfPlan_(
      template,
      visibleEnd,
      termsRow,
      lastRow
    );

    Logger.log('Last filled item row in source: ' + lastItem);
    Logger.log('Visible item rows: ' + (visibleEnd - QPDF_ITEM_FIRST + 1));
    Logger.log('Rows physically deleted: ' + deleteCount);
    Logger.log('Total price row after deletion: ' + totalPriceRow);
    Logger.log('Terms row after deletion: ' + termsRow);
    Logger.log('Item pages: ' + plan.itemPageCount);
    Logger.log('Item overflow: ' + plan.itemOverflow);
    Logger.log('Page plan: ' + JSON.stringify(plan.summary));

    // Performance optimization only: remove rows after the last content row
    // once, before copying page tabs. This does not alter any planned row,
    // height, image, column width, or page boundary; it only prevents every
    // template.copyTo() from carrying unused blank rows.
    qpdfDeleteAfter_(template, lastRow);
    SpreadsheetApp.flush();

    var pageNo = 0;
    for (var i = 0; i < plan.pages.length; i++) {
      pageNo++;
      qpdfMakePage_(
        template,
        tempBook,
        pageNo,
        plan.pages[i]
      );
    }

    if (!plan.pages.length) {
      throw new Error('No PDF pages were planned.');
    }

    qpdfOnlyPageTabs_(tempBook);
    SpreadsheetApp.flush();

    // A copied spreadsheet can need a short moment before its export endpoint
    // becomes available. qpdfExport_ retries a 404 instead of failing early.
    Utilities.sleep(1500);

    var pdf = qpdfExport_(tempBook);
    var fileName = qpdfFileName_(source);
    var file = outFolder.createFile(pdf.setName(fileName));
    var url = file.getUrl();

    Logger.log('PDF URL: ' + url);
    sourceBook.toast('PDF created: ' + fileName, 'Quotation PDF', 10);
    return url;
  } finally {
    tempFile.setTrashed(true);
  }
}

function qpdfLastItem_(sheet) {
  var rowCount = QPDF_ITEM_LAST - QPDF_ITEM_FIRST + 1;
  var colCount = Math.max(1, sheet.getLastColumn() - 1);
  var values = sheet.getRange(
    QPDF_ITEM_FIRST,
    2,
    rowCount,
    colCount
  ).getDisplayValues();
  var last = QPDF_ITEM_FIRST - 1;

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c] || '').trim() !== '') {
        last = QPDF_ITEM_FIRST + r;
        break;
      }
    }
  }

  return last;
}

function qpdfTermsRow_(sheet) {
  var found = sheet.createTextFinder(QPDF_TERMS_TEXT)
    .matchCase(false)
    .matchEntireCell(false)
    .findNext();

  if (found) return found.getRow();

  // Keep the known template fallback, but never return beyond the data.
  return Math.min(QPDF_TERMS_FALLBACK, sheet.getLastRow() + 1);
}

function qpdfFindRowContaining_(sheet, text, firstRow, lastRow) {
  if (firstRow > lastRow) return 0;

  var values = sheet.getRange(
    firstRow,
    1,
    lastRow - firstRow + 1,
    sheet.getLastColumn()
  ).getDisplayValues();
  var needle = String(text).toUpperCase();

  for (var r = 0; r < values.length; r++) {
    if (values[r].join(' ').toUpperCase().indexOf(needle) !== -1) {
      return firstRow + r;
    }
  }

  return 0;
}

function qpdfHeight_(sheet, first, last) {
  if (first > last) return 0;

  var total = 0;
  for (var r = first; r <= last; r++) {
    total += sheet.getRowHeight(r);
  }
  return total;
}

function qpdfPageHeight_(sheet) {
  var width = 0;
  for (var c = 1; c <= sheet.getLastColumn(); c++) {
    width += sheet.getColumnWidth(c);
  }

  return Math.max(
    1,
    Math.round(QPDF_PAGE_HEIGHT * width / QPDF_PAGE_WIDTH)
  );
}

function qpdfRows_(sheet, first, last) {
  var rows = [];
  if (first > last) return rows;

  for (var r = first; r <= last; r++) {
    rows.push({
      row: r,
      height: sheet.getRowHeight(r)
    });
  }
  return rows;
}

function qpdfChunks_(rows, capacity) {
  var result = [];
  var current = null;

  for (var i = 0; i < rows.length; i++) {
    var item = rows[i];

    if (!current) {
      current = {
        start: item.row,
        end: item.row,
        height: item.height
      };
    } else if (current.height + item.height <= capacity) {
      current.end = item.row;
      current.height += item.height;
    } else {
      result.push(current);
      current = {
        start: item.row,
        end: item.row,
        height: item.height
      };
    }
  }

  if (current) result.push(current);
  return result;
}

function qpdfItemChunks_(rows, normalCapacity) {
  if (!rows.length) return [];

  var averageHeight = qpdfRowsHeight_(rows) / rows.length;
  var firstCapacity = normalCapacity +
    averageHeight * QPDF_EXTRA_FIRST_PAGE_ROWS;
  var firstChunk = qpdfChunks_(rows, firstCapacity)[0];
  var consumed = firstChunk.end - rows[0].row + 1;
  var remaining = rows.slice(consumed);
  var result = [firstChunk];

  // All later item pages use the normal capacity. Only page 1 receives the
  // extra four-row allowance requested for the visible blank space.
  return result.concat(qpdfChunks_(remaining, normalCapacity));
}

function qpdfPlan_(sheet, visibleEnd, termsRow, lastRow) {
  var pageHeight = qpdfPageHeight_(sheet);
  var headerHeight = qpdfHeight_(sheet, 1, QPDF_HEADER_LAST);
  var itemCapacity = Math.max(1, pageHeight - headerHeight);
  var itemRows = qpdfRows_(sheet, QPDF_ITEM_FIRST, visibleEnd);
  var middleRows = qpdfRows_(sheet, visibleEnd + 1, termsRow - 1);
  var termsRows = qpdfRows_(sheet, termsRow, lastRow);
  var itemChunks = qpdfItemChunks_(itemRows, itemCapacity);
  var pages = [];

  for (var i = 0; i < itemChunks.length; i++) {
    pages.push({
      start: itemChunks[i].start,
      end: itemChunks[i].end,
      height: headerHeight + itemChunks[i].height,
      header: true,
      kind: 'items'
    });
  }

  // The normal quotation always has at least five visible product rows. Keep
  // a safe header-only fallback for an empty/malformed template.
  if (!pages.length) {
    pages.push({
      start: 1,
      end: QPDF_HEADER_LAST,
      height: headerHeight,
      header: true,
      kind: 'items'
    });
  }

  // Keep all charges, including TOTAL PRICE, on the final item page when the
  // required height is close enough to one page to be safely compacted. This
  // is the important case from the user's PDF: the Total Price line stays on
  // page one instead of being unnecessarily pushed down.
  if (middleRows.length) {
    var finalItemPage = pages[pages.length - 1];
    var middleHeight = qpdfRowsHeight_(middleRows);
    var withCharges = finalItemPage.height + middleHeight;
    var hasLaterPage = termsRows.length > 0;
    var reservedNote = hasLaterPage ? QPDF_CONTINUATION_HEIGHT : 0;
    var canCompact = withCharges + reservedNote <=
      pageHeight * QPDF_MAX_COMPRESSION_OVERFLOW;

    if (canCompact) {
      finalItemPage.end = middleRows[middleRows.length - 1].row;
      finalItemPage.height = withCharges;
    } else {
      var middleCapacity = Math.max(
        1,
        pageHeight - finalItemPage.height - reservedNote
      );
      var middleChunks = qpdfChunks_(middleRows, middleCapacity);

      if (middleChunks.length) {
        finalItemPage.end = middleChunks[0].end;
        finalItemPage.height += middleChunks[0].height;
      }

      for (var m = 1; m < middleChunks.length; m++) {
        pages.push({
          start: middleChunks[m].start,
          end: middleChunks[m].end,
          height: middleChunks[m].height,
          header: false,
          kind: 'charges'
        });
      }
    }
  }

  // Put the complete terms block on the last content page when it fits.
  // This is preferred over creating a new page: in the user's quotation the
  // terms block fits in the blank space below Total Price on page 2.
  var termsAppended = false;
  if (termsRows.length) {
    var finalContentPage = pages[pages.length - 1];
    var termsHeight = qpdfRowsHeight_(termsRows);
    var reservedNote = finalContentPage.continuation === false
      ? 0
      : QPDF_CONTINUATION_HEIGHT;
    var combinedHeight = finalContentPage.height + termsHeight + reservedNote;
    var canAppendTerms = combinedHeight <=
      pageHeight * QPDF_MAX_COMPRESSION_OVERFLOW;

    if (canAppendTerms) {
      finalContentPage.end = termsRows[termsRows.length - 1].row;
      finalContentPage.height = combinedHeight;
      finalContentPage.kind = 'content_with_terms';
      termsAppended = true;
    }
  }

  // Technical terms remain headerless. If the complete block does not fit on
  // the final content page, split it at row boundaries; the final Thanks/name
  // row is never discarded.
  if (termsRows.length && !termsAppended) {
    var termsCapacity = Math.max(
      1,
      pageHeight - QPDF_CONTINUATION_HEIGHT
    );
    var termsChunks = qpdfChunks_(termsRows, termsCapacity);

    for (var t = 0; t < termsChunks.length; t++) {
      pages.push({
        start: termsChunks[t].start,
        end: termsChunks[t].end,
        height: termsChunks[t].height,
        header: false,
        kind: 'terms'
      });
    }
  }

  for (var p = 0; p < pages.length; p++) {
    pages[p].continuation = p < pages.length - 1;
  }

  return {
    pages: pages,
    itemPageCount: itemChunks.length || 1,
    itemOverflow: itemChunks.length > 1,
    summary: pages.map(function(page) {
      return {
        kind: page.kind,
        header: page.header,
        start: page.start,
        end: page.end,
        height: page.height,
        continuation: page.continuation
      };
    })
  };
}

function qpdfRowsHeight_(rows) {
  var total = 0;
  for (var i = 0; i < rows.length; i++) total += rows[i].height;
  return total;
}

function qpdfMakePage_(template, book, number, plan) {
  // Copy the whole template so drawings/images are preserved. The previous
  // approach called OverGridImage.getBlob(), which is not available in Apps
  // Script and caused the execution to stall while logging image errors.
  var page = template.copyTo(book);
  page.setName('QPDF_PAGE_' + number);
  page.setHiddenGridlines(true);
  page.setFrozenRows(0);
  page.showRows(1, page.getMaxRows());

  qpdfRemoveWrongImages_(page, plan.start, plan.end, plan.header);
  qpdfDeleteAfter_(page, plan.end);

  if (plan.header) {
    var gap = plan.start - (QPDF_HEADER_LAST + 1);
    if (gap > 0) {
      page.deleteRows(QPDF_HEADER_LAST + 1, gap);
    }
  } else if (plan.start > 1) {
    page.deleteRows(1, plan.start - 1);
  }

  if (plan.continuation) {
    qpdfAddContinuationNote_(page);
  }

  // Give the final terms/Thanks row enough room before compacting only a
  // slightly oversized temporary page.
  if (plan.kind === 'terms') {
    var finalRow = page.getMaxRows();
    page.setRowHeight(
      finalRow,
      Math.max(page.getRowHeight(finalRow), 30)
    );
  }

  qpdfFitPageToOnePage_(page);
  SpreadsheetApp.flush();
}

function qpdfDeleteAfter_(page, lastKeptRow) {
  var extra = page.getMaxRows() - lastKeptRow;
  if (extra > 0) {
    page.deleteRows(lastKeptRow + 1, extra);
  }
}

function qpdfRemoveWrongImages_(page, bodyStart, bodyEnd, hasHeader) {
  var images = page.getImages();

  for (var i = images.length - 1; i >= 0; i--) {
    var row = images[i].getAnchorCell().getRow();
    var keepHeader = hasHeader && row >= 1 && row <= QPDF_HEADER_LAST;
    var keepBody = row >= bodyStart && row <= bodyEnd;

    if (!keepHeader && !keepBody) {
      try {
        images[i].remove();
      } catch (error) {
        Logger.log(
          'Image removal skipped at row ' + row + ': ' + error.message
        );
      }
    }
  }
}

function qpdfTrimNewPage_(page, lastKeptRow) {
  var extra = page.getMaxRows() - lastKeptRow;
  if (extra > 0) {
    page.deleteRows(lastKeptRow + 1, extra);
  }
}

function qpdfAddContinuationNote_(page) {
  var last = page.getMaxRows();
  var allowed = qpdfPageHeight_(page) - 8;
  var used = qpdfHeight_(page, 1, last);
  var spacerHeight = Math.floor(
    allowed - used - QPDF_CONTINUATION_HEIGHT
  );

  // Keep the continuation label at the physical bottom-right of the page.
  // Blank spacer rows use the already available white space; no content rows,
  // page boundaries, or item placement are changed.
  while (spacerHeight > 0) {
    page.insertRowAfter(last);
    var spacerRow = page.getMaxRows();
    var chunk = Math.min(400, spacerHeight);
    var spacer = page.getRange(
      spacerRow,
      1,
      1,
      page.getLastColumn()
    );
    spacer.clear();
    spacer.setBackground('#ffffff');
    spacer.setBorder(false, false, false, false, false, false);
    page.setRowHeight(spacerRow, Math.max(1, chunk));
    last = spacerRow;
    spacerHeight -= chunk;
  }

  page.insertRowAfter(last);
  var noteRow = page.getMaxRows();
  var note = page.getRange(
    noteRow,
    1,
    1,
    page.getLastColumn()
  );

  note.merge();
  note.setValue('Cont. on next page');
  note.setHorizontalAlignment('right');
  note.setVerticalAlignment('middle');
  note.setFontSize(7);
  note.setFontWeight('bold');
  note.setFontStyle('normal');
  note.setFontColor('#777777');
  page.setRowHeight(noteRow, QPDF_CONTINUATION_HEIGHT);
}

function qpdfFitPageToOnePage_(page) {
  var last = page.getMaxRows();
  var allowed = qpdfPageHeight_(page) - 8;
  var total = qpdfHeight_(page, 1, last);

  if (total <= allowed || total <= 0) return;

  var factor = allowed / total;
  for (var r = 1; r <= last; r++) {
    var oldHeight = page.getRowHeight(r);
    page.setRowHeight(
      r,
      Math.max(
        QPDF_MIN_ROW_HEIGHT,
        Math.floor(oldHeight * factor)
      )
    );
  }
}

function qpdfOnlyPageTabs_(book) {
  var sheets = book.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('QPDF_PAGE_') === 0) {
      sheets[i].showSheet();
    }
  }

  for (var j = 0; j < sheets.length; j++) {
    if (sheets[j].getName().indexOf('QPDF_PAGE_') !== 0) {
      sheets[j].hideSheet();
    }
  }
}

function qpdfExport_(book) {
  var base =
    'https://docs.google.com/spreadsheets/d/' +
    book.getId();

  var url = base +
    '/export?format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&scale=4' +
    '&top_margin=0.10' +
    '&bottom_margin=0.10' +
    '&left_margin=0.10' +
    '&right_margin=0.10' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false';

  var options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };

  var response;
  for (var attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      return response.getBlob().setContentType('application/pdf');
    }

    if (response.getResponseCode() !== 404 || attempt === 3) {
      throw new Error(
        'PDF export failed: ' +
        response.getResponseCode() + ' ' +
        response.getContentText().slice(0, 500)
      );
    }

    Utilities.sleep(1500 * attempt);
  }

  throw new Error('PDF export failed without a response.');
}

function qpdfFileName_(sheet) {
  var ref = String(
    sheet.getRange(QPDF_REF_CELL).getDisplayValue() || ''
  ).trim();

  var stamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd-MM-yyyy_HHmmss'
  );

  var safeRef = ref
    ? ref.replace(/[\\/:*?"<>|]/g, '-')
    : 'Quotation';

  return safeRef + '_' + stamp + '.pdf';
}
