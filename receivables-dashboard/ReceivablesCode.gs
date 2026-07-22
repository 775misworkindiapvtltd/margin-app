/**
 * RECEIVABLES DASHBOARD — Google Apps Script backend
 * Part of Finance 360° Command Center (standalone module for now — will be merged
 * into the main Finance 360° project later).
 *
 * Files in this Apps Script project:
 *   ReceivablesCode.gs        (this file — server-side)
 *   ReceivablesDashboard.html (the entire UI — CSS + HTML + client JS)
 *
 * HOW SHEETS ARE FOUND
 * ---------------------
 * Per the spec, we never hardcode column letters — headers are matched by name.
 * We also auto-detect WHICH TAB is the Receivables / Receipt / Balance sheet by
 * scanning every tab's header row for a required "signature" of column names,
 * instead of relying on an exact tab name (tab names differ between users/sheets).
 * If your tabs have different names, this still works as long as the header row
 * contains the expected column names (Party_Name, Pending Amount, Due_Date, etc).
 *
 * If you want to force a specific tab name, set the constants below.
 */

/* ---- Optional: force exact tab names (leave blank string '' to keep auto-detect) ---- */
var FORCE_SHEET_NAMES = {
  receivables: '',   // e.g. 'RECEIVABLES'
  receipt:     '',   // e.g. 'RECEIPT LIST'
  balance:     ''    // e.g. 'BALANCE SHEET' / 'LEDGER'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('ReceivablesDashboard')
    .evaluate()
    .setTitle('Receivables Dashboard — Finance 360°')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================= SHEET AUTO-DETECTION ============================= */

function normHeader_(h) {
  return String(h || '').toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

// Returns true if `headers` (already normalized) contains at least one alias from
// every group in `signature` (array of alias-arrays).
function headersMatchSignature_(normHeaders, signature) {
  return signature.every(function (aliasGroup) {
    return aliasGroup.some(function (alias) {
      return normHeaders.some(function (h) { return h === alias || h.indexOf(alias) !== -1; });
    });
  });
}

var SHEET_SIGNATURES = {
  receivables: [
    ['PARTY_NAME'],
    ['PENDING_AMOUNT'],
    ['DUE_DATE']
  ],
  receipt: [
    ['AMOUNT'],
    ['PARTY_NAME', 'CUSTOMER']
  ],
  balance: [
    ['VOUCHER_DATE'],
    ['VOUCHER_DEBIT', 'VOUCHER_CREDIT']
  ]
};

/**
 * Finds the sheet (tab) matching a required column signature.
 * kind = 'receivables' | 'receipt' | 'balance'
 * Returns the Sheet object, or null if nothing matches.
 */
function findSheetByKind_(kind) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var forcedName = FORCE_SHEET_NAMES[kind];
  if (forcedName) {
    var forced = ss.getSheetByName(forcedName);
    if (forced) return forced;
  }
  var signature = SHEET_SIGNATURES[kind];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastCol = sh.getLastColumn();
    if (lastCol < 1 || sh.getLastRow() < 1) continue;
    var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var normed = headerRow.map(normHeader_);
    if (headersMatchSignature_(normed, signature)) return sh;
  }
  return null;
}

/**
 * Reads a sheet fully and returns { headers: [...], rows: [[...], [...]] }.
 * Dates are returned as native Date objects (Apps Script's HtmlService bridge
 * serializes these to real JS Date objects on the client automatically).
 * Returns { found:false } if the sheet could not be located.
 */
function readSheetByKind_(kind) {
  var sh = findSheetByKind_(kind);
  if (!sh) return { found: false, headers: [], rows: [] };
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { found: true, sheetName: sh.getName(), headers: [], rows: [] };
  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  var rows = data.slice(1).filter(function (row) {
    return row.some(function (c) { return c !== '' && c !== null; });
  });
  return { found: true, sheetName: sh.getName(), headers: headers, rows: rows };
}

/* ============================= PUBLIC API (called from client) ============================= */

/**
 * Main bootstrap call — loads Receivables + Receipt + (lazy) Balance sheet metadata.
 * Balance/ledger rows are loaded on-demand per customer via getLedgerForCustomer()
 * to avoid pulling a potentially huge ledger sheet on every page load.
 */
function getReceivablesBootstrap() {
  var receivables = readSheetByKind_('receivables');
  var receipt = readSheetByKind_('receipt');
  var balanceSheet = findSheetByKind_('balance');
  return {
    receivables: receivables,
    receipt: receipt,
    balanceAvailable: !!balanceSheet,
    generatedAt: new Date()
  };
}

/**
 * Returns ledger rows (Balance sheet) filtered to one customer (by normalized
 * Party Name / Voucher Particular match), for the Customer Drill-down > Ledger tab.
 */
function getLedgerForCustomer(partyName) {
  var sh = findSheetByKind_('balance');
  if (!sh) return { found: false, headers: [], rows: [] };
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { found: true, headers: [], rows: [] };
  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  var normed = headers.map(normHeader_);
  var particularIdx = -1;
  for (var i = 0; i < normed.length; i++) {
    if (normed[i].indexOf('VOUCHER_PARTICULAR') !== -1 || normed[i].indexOf('PARTY_NAME') !== -1) { particularIdx = i; break; }
  }
  var target = String(partyName || '').trim().toUpperCase();
  var rows = data.slice(1).filter(function (row) {
    if (particularIdx === -1) return true;
    return String(row[particularIdx] || '').trim().toUpperCase().indexOf(target) !== -1;
  });
  return { found: true, headers: headers, rows: rows };
}

/**
 * Utility exposed to client to know the server's "today" (script timezone),
 * used to default the As-On-Date filter.
 */
function getServerToday() {
  return new Date();
}
