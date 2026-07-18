/**
 * MATERIAL RECEIVING — Standalone Web App (Google Apps Script)
 * Files in this Apps Script project:
 *   Code.gs      (this file)
 *   Index.html   (the UI file)
 *
 * Required sheet tabs:
 *   LOGIN PAGE              -> NAME | ID | PASSWORD | MATERIAL RECEIVED VIEW ENTRY | MATERIAL ENTRY ADD
 *   MATERIAL REC RESPONSES  -> TIMESTAMP | PO NO | INVOICE UPLOAD | VENDOR NAME | PENDING QTY | INV QTY |
 *                              DUE DATE | BANDEL | ADDRESS & GST NUMBER VERIFICATION | EWAY BILL VERIFICATION |
 *                              LR BILL VERIFIED | Invoice number | invoice date | EWAY BILL VERIFICATION IMAGE |
 *                              CHANGE BRAND | BRAND | SALES ORDER ID | ITEM NAME | PENDING QTY | REC QTY |
 *                              CANCEL QTY | PO RATE | SIZE | UNIT | INVOICE RATE | INWARD BATCH NO |
 *                              GROSS WEIGHT | REMARKS | NEW UNIQUE NO | OUTWARD BATCH NO |
 *                              MATRIAL REC IMAGE MULTIPLE IMAGE | REC QTY | PEND QTY | STATUS | LOGIN NAME
 *
 * Deploy: Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access "Anyone" > Deploy.
 */

/* ---- Sheet name constants ---- */
var SHEETS = {
  login:      'LOGIN PAGE',
  matRecResp: 'MATERIAL REC RESPONSES'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Material Receiving')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sheetToObjects_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  return data.slice(1)
    .filter(function (row) { return row.some(function (c) { return c !== ''; }); })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function fmtTimestamp_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss');
  return (v === undefined || v === null) ? '' : v;
}
function fmtDateOnly_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  return (v === undefined || v === null) ? '' : v;
}
function fmtValue_(v) { return (v === undefined || v === null) ? '' : v; }
function isYes_(v) { return String(v || '').trim().toUpperCase() === 'YES'; }

function pick_(r, names) {
  for (var i = 0; i < names.length; i++) {
    if (Object.prototype.hasOwnProperty.call(r, names[i])) return r[names[i]];
  }
  var keys = Object.keys(r);
  for (var i = 0; i < names.length; i++) {
    var target = names[i].replace(/\s+/g, ' ').trim().toUpperCase();
    for (var j = 0; j < keys.length; j++) {
      if (keys[j].replace(/\s+/g, ' ').trim().toUpperCase() === target) return r[keys[j]];
    }
  }
  return '';
}

function mapUser_(r) {
  return {
    name: r['NAME'] || '', id: String(r['ID'] || '').trim(), password: String(r['PASSWORD'] || '').trim(),
    matRecView: isYes_(pick_(r, ['MATERIAL RECEIVED VIEW ENTRY'])),
    matRecAdd:  isYes_(pick_(r, ['MATERIAL ENTRY ADD']))
  };
}

function mapMatRec_(rows) {
  return rows.map(function (r) {
    return {
      timestamp: fmtTimestamp_(r['TIMESTAMP']),
      poNo: fmtValue_(r['PO NO']),
      invoiceUpload: fmtValue_(r['INVOICE UPLOAD']),
      vendorName: fmtValue_(r['VENDOR NAME']),
      pendingQtyTop: fmtValue_(r['PENDING QTY']),
      invQty: fmtValue_(r['INV QTY']),
      dueDate: fmtDateOnly_(r['DUE DATE']),
      bandel: fmtValue_(r['BANDEL']),
      addressGst: fmtValue_(r['ADDRESS & GST NUMBER VERIFICATION']),
      ewayBill: fmtValue_(r['EWAY BILL VERIFICATION']),
      lrBillVerified: fmtValue_(r['LR BILL VERIFIED']),
      invoiceNumber: fmtValue_(r['Invoice number']),
      invoiceDate: fmtDateOnly_(r['invoice date']),
      ewayBillImage: fmtValue_(r['EWAY BILL VERIFICATION IMAGE']),
      changeBrand: fmtValue_(r['CHANGE BRAND']),
      brand: fmtValue_(r['BRAND']),
      salesOrderId: fmtValue_(r['SALES ORDER ID']),
      itemName: fmtValue_(r['ITEM NAME']),
      pendingQty: fmtValue_(r['PENDING QTY']),
      recQty: fmtValue_(r['REC QTY']),
      cancelQty: fmtValue_(r['CANCEL QTY']),
      poRate: fmtValue_(r['PO RATE']),
      size: fmtValue_(r['SIZE']),
      unit: fmtValue_(r['UNIT']),
      invoiceRate: fmtValue_(r['INVOICE RATE']),
      inwardBatchNo: fmtValue_(r['INWARD BATCH NO']),
      grossWeight: fmtValue_(r['GROSS WEIGHT']),
      remarks: fmtValue_(r['REMARKS']),
      newUniqueNo: fmtValue_(r['NEW UNIQUE NO']),
      outwardBatchNo: fmtValue_(r['OUTWARD BATCH NO']),
      matRecImage: fmtValue_(r['MATRIAL REC IMAGE MULTIPLE IMAGE']),
      status: fmtValue_(r['STATUS']),
      loginName: fmtValue_(r['LOGIN NAME'] || r['LOGIN ID'] || '')
    };
  });
}

function getUserPermissions(id) {
  var usersRaw = sheetToObjects_(SHEETS.login);
  var match = usersRaw.find(function (r) { return String(r['ID'] || '').trim().toLowerCase() === String(id || '').trim().toLowerCase(); });
  if (!match) return null;
  var u = mapUser_(match);
  delete u.password;
  return u;
}

function getLoginData() {
  var usersRaw = sheetToObjects_(SHEETS.login);
  return { users: usersRaw.map(mapUser_) };
}

function getBootstrapData(perms) {
  var loadAll = !perms;
  var needMatRec = loadAll || perms.matRecView || perms.matRecAdd;

  var result = {
    matRecResp: needMatRec ? mapMatRec_(sheetToObjects_(SHEETS.matRecResp)) : [],
    users: sheetToObjects_(SHEETS.login).map(mapUser_),
    missingSheets: []
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var have = {};
  ss.getSheets().forEach(function (s) { have[s.getName().trim().toUpperCase()] = true; });
  Object.keys(SHEETS).forEach(function (k) {
    if (!have[SHEETS[k].toUpperCase()]) result.missingSheets.push(SHEETS[k]);
  });

  return result;
}

/**
 * Save Material Received entries.
 * payload = { topFields:{...}, rows:[{...},...], loginId, loginName }
 */
function saveMatRecEntries(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload) return { status: 'error', message: 'No data received.' };

    var perms = getUserPermissions(payload.loginId || '');
    if (!perms || !perms.matRecAdd) return { status: 'error', message: 'You do not have permission to add Material Received entries.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEETS.matRecResp);
    if (!sh) return { status: 'error', message: 'Sheet not found: ' + SHEETS.matRecResp };

    var top = payload.topFields || {};
    var rows = payload.rows || [];
    if (!rows.length) return { status: 'error', message: 'No item rows to save.' };

    var now = new Date();
    var ts = fmtTimestamp_(now);
    var loginName = payload.loginName || '';

    var dataRows = rows.map(function (r) {
      return [
        ts,
        top.poNo || '',
        top.invoiceUpload || '',
        top.vendorName || '',
        top.pendingQtyTop || '',
        top.invQty || '',
        top.dueDate || '',
        top.bandel || '',
        top.addressGst || '',
        top.ewayBill || '',
        top.lrBillVerified || '',
        top.invoiceNumber || '',
        top.invoiceDate || '',
        top.ewayBillImage || '',
        r.changeBrand || '',
        r.brand || '',
        r.salesOrderId || '',
        r.itemName || '',
        r.pendingQty || '',
        r.recQty || '',
        r.cancelQty || '',
        r.poRate || '',
        r.size || '',
        r.unit || '',
        r.invoiceRate || '',
        r.inwardBatchNo || '',
        r.grossWeight || '',
        r.remarks || '',
        r.newUniqueNo || '',
        r.outwardBatchNo || '',
        r.matRecImage || '',
        loginName
      ];
    });

    sh.getRange(sh.getLastRow() + 1, 1, dataRows.length, 32).setValues(dataRows);

    return { status: 'ok', saved: dataRows.length };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}
