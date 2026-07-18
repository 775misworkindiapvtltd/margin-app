/**
 * MARGIN ENTRY — Equity + Commodity  (Google Apps Script web app)
 * Files in this Apps Script project:
 *   Code.gs      (this file)
 *   Index.html   (the UI file)
 *
 * Required sheet tabs (names must match EXACTLY):
 *   LOGIN PAGE            -> NAME | ID | PASSWORD | EQUITY ENTRY | COMODDITY ENTRY | MASTER EQUITY | MASTER COMODDITY | RESPONSES 2 EQUITY | RESPONSES EQUITY | RESPONSES COMODDITY | RESPONSES 2 COMODDITY | MASTER EQUITY ADD ENTRY | MASTER COMODDITY ADD ENTRY
 *   MASTER EQUITY         -> CODE | USER ID | SOFTWARE | NAME | GROUP NAME | BRANCH NAME | TIMESTAMP | LOGIN ID
 *   MASTER COMODDITY      -> CODE | USER ID | SOFTWARE | NAME | GROUP NAME | BRANCH NAME | TIMESTAMP | LOGIN ID
 *
 * NOTE: "MASTER EQUITY ADD ENTRY" / "MASTER COMODDITY ADD ENTRY" are new permission
 * columns on LOGIN PAGE. Mark YES for a user to show them the "+ Add" button on the
 * corresponding Master Data page (lets them add a new master row via a popup form).
 * The code also tolerates the "MATER ..." (missing S) spelling in case that's what
 * already exists in your sheet — either header name works.
 *   RESPONSES EQUITY      -> TIMESTAMP | SELECT DATE | CODE | USER ID | SOFTWARE | NAME | GROUP NAME | MARGIN AS PER RMS | MARGIN ALLOCATED ON ID | BRANCH NAME | LOGIN NAME
 *   RESPONSES COMODDITY   -> (same columns as RESPONSES EQUITY)
 *   RESPONSES 2 EQUITY    -> TIMESTAMP | SELECT DATE | GROUP NAME | GROUP MARGIN | LOGIN NAME
 *   RESPONSES 2 COMODDITY -> (same columns as RESPONSES 2 EQUITY)
 *
 * NOTE: Branch Name is now read from the MASTER sheets (per CODE), same as
 * User ID / Software / Name / Group Name — NOT from a separate DROPDOWN sheet.
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
  // Normalize header text: collapse any internal whitespace (including newlines from
  // wrapped cells) into single spaces so that "MATER\nEQUITY\nADD ENTRY" matches
  // the expected key "MATER EQUITY ADD ENTRY".
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

/* ---- Row mappers ---- */
function mapMaster_(rows) {
  return rows.map(function (r) {
    return {
      code: r['CODE'] || '', userId: r['USER ID'] || '', software: r['SOFTWARE'] || '', name: r['NAME'] || '',
      group: r['GROUP NAME'] || '', branchName: r['BRANCH NAME'] || '',
      timestamp: fmtTimestamp_(r['TIMESTAMP']), loginId: r['LOGIN ID'] || ''
    };
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
// Reads a header value tolerating alternate spellings. The comparison is case-insensitive
// and whitespace-normalized so minor differences (extra space, different case) are handled.
function pick_(r, names) {
  // First: try exact match (fastest path)
  for (var i = 0; i < names.length; i++) {
    if (Object.prototype.hasOwnProperty.call(r, names[i])) return r[names[i]];
  }
  // Fallback: case-insensitive + whitespace-normalized fuzzy match against all keys in r
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
    equityEntry:     isYes_(pick_(r, ['EQUITY ENTRY'])),
    commodityEntry:  isYes_(pick_(r, ['COMODDITY ENTRY', 'COMMODITY ENTRY'])),
    masterEquity:    isYes_(pick_(r, ['MASTER EQUITY'])),
    masterCommodity: isYes_(pick_(r, ['MASTER COMODDITY', 'MASTER COMMODITY'])),
    resp1Equity:     isYes_(pick_(r, ['RESPONSES EQUITY'])),
    resp1Commodity:  isYes_(pick_(r, ['RESPONSES COMODDITY', 'RESPONSES COMMODITY'])),
    resp2Equity:     isYes_(pick_(r, ['RESPONSES 2 EQUITY'])),
    resp2Commodity:  isYes_(pick_(r, ['RESPONSES 2 COMODDITY', 'RESPONSES 2 COMMODITY'])),
    masterEquityAdd:    isYes_(pick_(r, ['MASTER EQUITY ADD ENTRY', 'MATER EQUITY ADD ENTRY'])),
    masterCommodityAdd: isYes_(pick_(r, ['MASTER COMODDITY ADD ENTRY', 'MATER COMODDITY ADD ENTRY', 'MASTER COMMODITY ADD ENTRY', 'MATER COMMODITY ADD ENTRY']))
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

/**
 * Lightweight login-only data: reads only the LOGIN PAGE sheet (fast, <1s).
 * Called on initial page load so the login screen appears instantly while heavy
 * data (Master/Responses) loads later after successful login.
 */
function getLoginData() {
  var usersRaw = sheetToObjects_(SHEETS.login);
  return { users: usersRaw.map(mapUser_) };
}

function getBootstrapData(forceFresh) {
  var cache = CacheService.getScriptCache();
  if (!forceFresh) {
    var cached = cache.get('bootstrap_v4');
    if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var result = {
    masterEquity:    mapMaster_(sheetToObjects_(SHEETS.masterEquity)),
    masterCommodity: mapMaster_(sheetToObjects_(SHEETS.masterComm)),
    respEquity:      mapResp1_(sheetToObjects_(SHEETS.respEquity)),
    respCommodity:   mapResp1_(sheetToObjects_(SHEETS.respComm)),
    resp2Equity:     mapResp2_(sheetToObjects_(SHEETS.resp2Equity)),
    resp2Commodity:  mapResp2_(sheetToObjects_(SHEETS.resp2Comm)),
    users:           sheetToObjects_(SHEETS.login).map(mapUser_),
    missingSheets:   []
  };

  // Warn about any missing tabs (helps debugging)
  var have = {};
  ss.getSheets().forEach(function (s) { have[s.getName().trim().toUpperCase()] = true; });
  Object.keys(SHEETS).forEach(function (k) {
    if (!have[SHEETS[k].toUpperCase()]) result.missingSheets.push(SHEETS[k]);
  });

  try { cache.put('bootstrap_v4', JSON.stringify(result), 60); } catch (e) {}
  return result;
}

/**
 * Add one new row to a MASTER sheet (Equity or Commodity).
 * payload = { segment:'equity'|'commodity', code, userId, software, name, group, branchName, loginId }
 * Appends CODE|USER ID|SOFTWARE|NAME|GROUP NAME|BRANCH NAME|TIMESTAMP|LOGIN ID.
 * Server-side re-checks the "...ADD ENTRY" permission before writing, so this can't be
 * called by a user who doesn't have the button shown to them.
 */
function addMasterEntry(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!payload) return { status: 'error', message: 'No data received.' };
    var seg = payload.segment === 'commodity' ? 'commodity' : 'equity';

    // Permission check (server-side, cannot be bypassed from the client)
    var perms = getUserPermissions(payload.loginId);
    var allowed = perms && (seg === 'commodity' ? perms.masterCommodityAdd : perms.masterEquityAdd);
    if (!allowed) return { status: 'error', message: 'You do not have permission to add master entries.' };

    var code = String(payload.code || '').trim();
    var userId = String(payload.userId || '').trim();
    var software = String(payload.software || '').trim();
    var name = String(payload.name || '').trim();
    var group = String(payload.group || '').trim();
    var branchName = String(payload.branchName || '').trim();
    if (!code || !userId || !software || !name || !group || !branchName) {
      return { status: 'error', message: 'All 6 fields are required.' };
    }

    var sheetName = seg === 'commodity' ? SHEETS.masterComm : SHEETS.masterEquity;
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh) return { status: 'error', message: 'Sheet not found: ' + sheetName };

    var now = new Date();
    sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[code, userId, software, name, group, branchName, now, payload.loginId || '']]);

    CacheService.getScriptCache().remove('bootstrap_v4');
    return {
      status: 'ok', segment: seg,
      row: { code: code, userId: userId, software: software, name: name, group: group, branchName: branchName, timestamp: fmtTimestamp_(now), loginId: payload.loginId || '' }
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
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

    CacheService.getScriptCache().remove('bootstrap_v4');
    return { status: 'ok', savedA: savedA, savedB: savedB, segment: seg };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}
