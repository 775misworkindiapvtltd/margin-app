/**
 * MARGIN ENTRY — Equity + Commodity  (Google Apps Script web app)
 * Files in this Apps Script project:
 *   Code.gs      (this file)
 *   Index.html   (the UI file)
 *
 * Required sheet tabs (names must match EXACTLY):
 *   LOGIN PAGE            -> NAME | ID | PASSWORD | EQUITY ENTRY | COMODDITY ENTRY | MASTER EQUITY | MASTER COMODDITY | RESPONSES 2 EQUITY | RESPONSES EQUITY | RESPONSES COMODDITY | RESPONSES 2 COMODDITY
 *   MASTER EQUITY         -> CODE | USER ID | SOFTWARE | NAME | GROUP NAME
 *   MASTER COMODDITY      -> CODE | USER ID | SOFTWARE | NAME | GROUP NAME
 *   RESPONSES EQUITY      -> TIMESTAMP | SELECT DATE | CODE | USER ID | SOFTWARE | NAME | GROUP NAME | MARGIN AS PER RMS | MARGIN ALLOCATED ON ID | BRANCH NAME | LOGIN NAME
 *   RESPONSES COMODDITY   -> (same columns as RESPONSES EQUITY)
 *   RESPONSES 2 EQUITY    -> TIMESTAMP | SELECT DATE | GROUP NAME | GROUP MARGIN | LOGIN NAME
 *   RESPONSES 2 COMODDITY -> (same columns as RESPONSES 2 EQUITY)
 *   DROPDOWN              -> Branch names in Column B
 *
 * Deploy: Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access "Anyone" > Deploy.
 */

/* ---- Sheet name constants ---- */
var SHEETS = {
  login:        'LOGIN PAGE',
  masterEquity: 'MASTER EQUITY',
  masterComm:   'MASTER COMODDITY',
  respEquity:   'RESPONSES EQUITY',
  respComm:     'RESPONSES COMODDITY',
  resp2Equity:  'RESPONSES 2 EQUITY',
  resp2Comm:    'RESPONSES 2 COMODDITY'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Margin Entry Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sheetToObjects_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function (h) { return String(h).trim(); });
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

/* ---- Row mappers ---- */
function mapMaster_(rows) {
  return rows.map(function (r) {
    return { code: r['CODE'] || '', userId: r['USER ID'] || '', software: r['SOFTWARE'] || '', name: r['NAME'] || '', group: r['GROUP NAME'] || '' };
  });
}
function mapResp1_(rows) {
  return rows.map(function (r) {
    return {
      timestamp: fmtTimestamp_(r['TIMESTAMP']), date: fmtDateOnly_(r['SELECT DATE']),
      code: r['CODE'] || '', userId: r['USER ID'] || '', software: r['SOFTWARE'] || '', name: r['NAME'] || '',
      group: r['GROUP NAME'] || '', marginRMS: fmtValue_(r['MARGIN AS PER RMS']), marginAllocated: fmtValue_(r['MARGIN ALLOCATED ON ID']),
      loginName: r['LOGIN NAME'] || ''
    };
  });
}
function mapResp2_(rows) {
  return rows.map(function (r) {
    return { timestamp: fmtTimestamp_(r['TIMESTAMP']), date: fmtDateOnly_(r['SELECT DATE']), group: r['GROUP NAME'] || '', margin: fmtValue_(r['GROUP MARGIN']), loginName: r['LOGIN NAME'] || '' };
  });
}
function mapUser_(r) {
  return {
    name: r['NAME'] || '', id: String(r['ID'] || '').trim(), password: String(r['PASSWORD'] || '').trim(),
    equityEntry:     isYes_(r['EQUITY ENTRY']),
    commodityEntry:  isYes_(r['COMODDITY ENTRY']),
    masterEquity:    isYes_(r['MASTER EQUITY']),
    masterCommodity: isYes_(r['MASTER COMODDITY']),
    resp1Equity:     isYes_(r['RESPONSES EQUITY']),
    resp1Commodity:  isYes_(r['RESPONSES COMODDITY']),
    resp2Equity:     isYes_(r['RESPONSES 2 EQUITY']),
    resp2Commodity:  isYes_(r['RESPONSES 2 COMODDITY'])
  };
}

// Re-check one user's permissions (used when switching pages, so a revoked YES takes effect fast)
function getUserPermissions(id) {
  var usersRaw = sheetToObjects_(SHEETS.login);
  var match = usersRaw.find(function (r) { return String(r['ID'] || '').trim().toLowerCase() === String(id || '').trim().toLowerCase(); });
  if (!match) return null;
  var u = mapUser_(match);
  delete u.password;
  return u;
}

function getBootstrapData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('bootstrap_v3');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var result = {
    masterEquity:    mapMaster_(sheetToObjects_(SHEETS.masterEquity)),
    masterCommodity: mapMaster_(sheetToObjects_(SHEETS.masterComm)),
    respEquity:      mapResp1_(sheetToObjects_(SHEETS.respEquity)),
    respCommodity:   mapResp1_(sheetToObjects_(SHEETS.respComm)),
    resp2Equity:     mapResp2_(sheetToObjects_(SHEETS.resp2Equity)),
    resp2Commodity:  mapResp2_(sheetToObjects_(SHEETS.resp2Comm)),
    users:           sheetToObjects_(SHEETS.login).map(mapUser_),
    branchName:      [],
    branchError:     '',
    missingSheets:   []
  };

  // Warn about any missing tabs (helps debugging)
  var have = {};
  ss.getSheets().forEach(function (s) { have[s.getName().trim().toUpperCase()] = true; });
  Object.keys(SHEETS).forEach(function (k) {
    if (!have[SHEETS[k].toUpperCase()]) result.missingSheets.push(SHEETS[k]);
  });

  // Branch names from DROPDOWN tab, column B
  try {
    var ddSheet = null, allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var nm = allSheets[s].getName().trim().toUpperCase();
      if (nm === 'DROPDOWN' || nm === 'DROPDOWNS' || nm === 'DROP DOWN') { ddSheet = allSheets[s]; break; }
    }
    if (ddSheet) {
      var lastRow = ddSheet.getLastRow();
      if (lastRow >= 1) {
        var ddData = ddSheet.getRange('B1:B' + lastRow).getValues(), seen = {};
        ddData.forEach(function (row) {
          var v = String(row[0] || '').trim();
          if (v && v.toUpperCase() !== 'BRANCH NAME' && v.toUpperCase() !== 'BRANCH' && v.toUpperCase() !== 'NAME' && !seen[v]) {
            result.branchName.push(v); seen[v] = true;
          }
        });
        result.branchName.sort();
      }
    } else {
      result.branchError = 'DROPDOWN sheet not found.';
    }
  } catch (e) { result.branchError = e.message; }

  try { cache.put('bootstrap_v3', JSON.stringify(result), 60); } catch (e) {}
  return result;
}

/**
 * Save entries for one segment.
 * payload = { segment:'equity'|'commodity', responses1:[...], responses2:[...] }
 */
function saveEntries(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var seg = (payload && payload.segment) === 'commodity' ? 'commodity' : 'equity';
    var s1Name = seg === 'commodity' ? SHEETS.respComm  : SHEETS.respEquity;
    var s2Name = seg === 'commodity' ? SHEETS.resp2Comm : SHEETS.resp2Equity;
    var savedA = 0, savedB = 0;

    if (payload.responses1 && payload.responses1.length) {
      var sheet1 = ss.getSheetByName(s1Name);
      if (!sheet1) return { status: 'error', message: 'Sheet not found: ' + s1Name };
      var rowsA = payload.responses1.map(function (r) {
        // TIMESTAMP | SELECT DATE | CODE | USER ID | SOFTWARE | NAME | GROUP NAME | MARGIN AS PER RMS | MARGIN ALLOCATED ON ID | BRANCH NAME | LOGIN NAME
        return [r.timestamp, r.date, r.code, r.userId, r.software, r.name, r.group, r.marginRMS, r.marginAllocated, r.branchName || '', r.loginName || ''];
      });
      sheet1.getRange(sheet1.getLastRow() + 1, 1, rowsA.length, 11).setValues(rowsA);
      savedA = rowsA.length;
    }

    if (payload.responses2 && payload.responses2.length) {
      var sheet2 = ss.getSheetByName(s2Name);
      if (!sheet2) return { status: 'error', message: 'Sheet not found: ' + s2Name };
      var rowsB = payload.responses2.map(function (r) {
        // TIMESTAMP | SELECT DATE | GROUP NAME | GROUP MARGIN | LOGIN NAME
        return [r.timestamp, r.date, r.group, r.margin, r.loginName || ''];
      });
      sheet2.getRange(sheet2.getLastRow() + 1, 1, rowsB.length, 5).setValues(rowsB);
      savedB = rowsB.length;
    }

    CacheService.getScriptCache().remove('bootstrap_v3');
    return { status: 'ok', savedA: savedA, savedB: savedB, segment: seg };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}
