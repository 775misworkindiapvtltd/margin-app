/**
 * DATABASE.GS
 * All Google Sheet read/write logic + punch in/out business rules live here.
 */

var SHEET_EMP  = 'Employees';
var SHEET_ATT  = 'Attendance';

/* Employees columns (1-based) */
var EMP_COL = { id: 1, name: 2, department: 3, role: 4, password: 5, active: 6 };
/* Attendance columns (1-based) */
var ATT_COL = {
  date: 1, empId: 2, empName: 3, punchIn: 4, punchOut: 5, workingHours: 6, status: 7,
  latIn: 8, lngIn: 9, addrIn: 10, selfieIn: 11,
  latOut: 12, lngOut: 13, addrOut: 14, selfieOut: 15,
  device: 16, browser: 17, lateIn: 18, earlyOut: 19, remarks: 20
};

/** Creates the Employees / Attendance tabs (with headers + one sample admin) if they don't already exist. */
function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var emp = ss.getSheetByName(SHEET_EMP);
  if (!emp) {
    emp = ss.insertSheet(SHEET_EMP);
    emp.getRange(1, 1, 1, 6).setValues([['Employee ID', 'Name', 'Department', 'Role', 'Password', 'Active']]);
    emp.getRange(2, 1, 2, 6).setValues([
      ['ADM001', 'Admin User', 'Management', 'Admin', 'admin123', 'YES'],
      ['EMP001', 'Test Employee', 'Operations', 'Employee', 'emp123', 'YES']
    ]);
    emp.setFrozenRows(1);
  }

  var att = ss.getSheetByName(SHEET_ATT);
  if (!att) {
    att = ss.insertSheet(SHEET_ATT);
    att.getRange(1, 1, 1, 20).setValues([[
      'Date', 'Employee ID', 'Employee Name', 'Punch In Time', 'Punch Out Time', 'Working Hours', 'Status',
      'Latitude In', 'Longitude In', 'Address In', 'Selfie In',
      'Latitude Out', 'Longitude Out', 'Address Out', 'Selfie Out',
      'Device', 'Browser', 'Late In', 'Early Out', 'Remarks'
    ]]);
    att.setFrozenRows(1);
  }
}

/* ---- generic sheet -> array-of-objects helper (kept local & simple, no header assumptions needed since we use fixed columns above) ---- */
function getSheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function fmtDate_(d) {
  return Utilities.formatDate(d, CONFIG.timeZone, 'dd-MMM-yyyy');
}
function fmtDateKey_(d) {
  return Utilities.formatDate(d, CONFIG.timeZone, 'yyyy-MM-dd');
}
function fmtTime_(d) {
  return Utilities.formatDate(d, CONFIG.timeZone, 'hh:mm a');
}
function todayKey_() {
  return fmtDateKey_(new Date());
}

/* ============================================================
 * LOGIN
 * ============================================================ */
function login(employeeId, password) {
  ensureSheets_();
  var sh = getSheet_(SHEET_EMP);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = String(row[EMP_COL.id - 1] || '').trim();
    if (!id) continue;
    if (id.toLowerCase() === String(employeeId || '').trim().toLowerCase()) {
      var pw = String(row[EMP_COL.password - 1] || '').trim();
      var active = String(row[EMP_COL.active - 1] || '').trim().toUpperCase() === 'YES';
      if (!active) return { status: 'error', message: 'Your account is inactive. Contact admin.' };
      if (pw !== String(password || '').trim()) return { status: 'error', message: 'Invalid Employee ID or Password.' };
      return {
        status: 'ok',
        user: {
          id: id,
          name: row[EMP_COL.name - 1] || '',
          department: row[EMP_COL.department - 1] || '',
          role: row[EMP_COL.role - 1] || 'Employee'
        }
      };
    }
  }
  return { status: 'error', message: 'Invalid Employee ID or Password.' };
}

/* ============================================================
 * TODAY STATUS  — used by dashboard to decide which big button to show
 * ============================================================ */
function getTodayStatus(employeeId) {
  ensureSheets_();
  var row = findTodayRow_(employeeId);
  if (!row) return { punchedIn: false, punchedOut: false };
  return {
    punchedIn: !!row.values[ATT_COL.punchIn - 1],
    punchedOut: !!row.values[ATT_COL.punchOut - 1],
    punchInTime: row.values[ATT_COL.punchIn - 1] ? fmtTime_(row.values[ATT_COL.punchIn - 1]) : '',
    punchOutTime: row.values[ATT_COL.punchOut - 1] ? fmtTime_(row.values[ATT_COL.punchOut - 1]) : ''
  };
}

/** Finds today's attendance row (by date key + employee id). Returns {rowIndex(1-based), values} or null. */
function findTodayRow_(employeeId) {
  var sh = getSheet_(SHEET_ATT);
  var data = sh.getDataRange().getValues();
  var key = todayKey_();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][ATT_COL.date - 1];
    if (!d) continue;
    var dk = (d instanceof Date) ? fmtDateKey_(d) : String(d);
    var eid = String(data[i][ATT_COL.empId - 1] || '').trim();
    if (dk === key && eid.toLowerCase() === String(employeeId || '').trim().toLowerCase()) {
      return { rowIndex: i + 1, values: data[i] };
    }
  }
  return null;
}

/* ============================================================
 * PUNCH IN
 * payload = { employeeId, employeeName, lat, lng, selfieBase64, device, browser }
 * ============================================================ */
function punchIn(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    if (!payload) return { status: 'error', message: 'No data received.' };
    var employeeId = String(payload.employeeId || '').trim();
    if (!employeeId) return { status: 'error', message: 'Employee not identified. Please login again.' };
    if (payload.lat == null || payload.lng == null) return { status: 'error', message: 'Location is required. Please allow GPS access and try again.' };
    if (!payload.selfieBase64) return { status: 'error', message: 'Selfie is required. Please capture a photo and try again.' };

    var existing = findTodayRow_(employeeId);
    if (existing && existing.values[ATT_COL.punchIn - 1]) {
      return { status: 'error', message: 'You have already punched in today. Duplicate punch-in is not allowed.' };
    }

    var now = new Date();
    var address = reverseGeocode_(payload.lat, payload.lng);
    var selfieUrl = saveSelfieToDrive_(payload.selfieBase64, employeeId, 'IN');
    var lateInfo = computeLateIn_(now);

    var sh = getSheet_(SHEET_ATT);
    if (existing) {
      // Row for today already exists (shouldn't normally happen without punch-in, but be safe)
      var r = existing.rowIndex;
      sh.getRange(r, ATT_COL.punchIn).setValue(now);
      sh.getRange(r, ATT_COL.status).setValue('Present');
      sh.getRange(r, ATT_COL.latIn).setValue(payload.lat);
      sh.getRange(r, ATT_COL.lngIn).setValue(payload.lng);
      sh.getRange(r, ATT_COL.addrIn).setValue(address);
      sh.getRange(r, ATT_COL.selfieIn).setValue(selfieUrl);
      sh.getRange(r, ATT_COL.device).setValue(payload.device || '');
      sh.getRange(r, ATT_COL.browser).setValue(payload.browser || '');
      sh.getRange(r, ATT_COL.lateIn).setValue(lateInfo.label);
    } else {
      var rowVals = new Array(20).fill('');
      rowVals[ATT_COL.date - 1] = now;
      rowVals[ATT_COL.empId - 1] = employeeId;
      rowVals[ATT_COL.empName - 1] = payload.employeeName || '';
      rowVals[ATT_COL.punchIn - 1] = now;
      rowVals[ATT_COL.status - 1] = 'Present';
      rowVals[ATT_COL.latIn - 1] = payload.lat;
      rowVals[ATT_COL.lngIn - 1] = payload.lng;
      rowVals[ATT_COL.addrIn - 1] = address;
      rowVals[ATT_COL.selfieIn - 1] = selfieUrl;
      rowVals[ATT_COL.device - 1] = payload.device || '';
      rowVals[ATT_COL.browser - 1] = payload.browser || '';
      rowVals[ATT_COL.lateIn - 1] = lateInfo.label;
      sh.getRange(sh.getLastRow() + 1, 1, 1, 20).setValues([rowVals]);
    }

    return { status: 'ok', punchInTime: fmtTime_(now), late: lateInfo.isLate, lateBy: lateInfo.label };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * PUNCH OUT
 * payload = { employeeId, lat, lng, selfieBase64, device, browser }
 * ============================================================ */
function punchOut(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    if (!payload) return { status: 'error', message: 'No data received.' };
    var employeeId = String(payload.employeeId || '').trim();
    if (!employeeId) return { status: 'error', message: 'Employee not identified. Please login again.' };
    if (payload.lat == null || payload.lng == null) return { status: 'error', message: 'Location is required. Please allow GPS access and try again.' };
    if (!payload.selfieBase64) return { status: 'error', message: 'Selfie is required. Please capture a photo and try again.' };

    var existing = findTodayRow_(employeeId);
    if (!existing || !existing.values[ATT_COL.punchIn - 1]) {
      return { status: 'error', message: 'You must Punch In before you can Punch Out.' };
    }
    if (existing.values[ATT_COL.punchOut - 1]) {
      return { status: 'error', message: 'You have already punched out today. Duplicate punch-out is not allowed.' };
    }

    var now = new Date();
    var address = reverseGeocode_(payload.lat, payload.lng);
    var selfieUrl = saveSelfieToDrive_(payload.selfieBase64, employeeId, 'OUT');
    var punchInTime = existing.values[ATT_COL.punchIn - 1];
    var earlyInfo = computeEarlyOut_(now);
    var hours = computeWorkingHours_(punchInTime, now);

    var sh = getSheet_(SHEET_ATT);
    var r = existing.rowIndex;
    sh.getRange(r, ATT_COL.punchOut).setValue(now);
    sh.getRange(r, ATT_COL.workingHours).setValue(hours.label);
    sh.getRange(r, ATT_COL.latOut).setValue(payload.lat);
    sh.getRange(r, ATT_COL.lngOut).setValue(payload.lng);
    sh.getRange(r, ATT_COL.addrOut).setValue(address);
    sh.getRange(r, ATT_COL.selfieOut).setValue(selfieUrl);
    sh.getRange(r, ATT_COL.earlyOut).setValue(earlyInfo.label);
    if (payload.device) sh.getRange(r, ATT_COL.device).setValue(payload.device);
    if (payload.browser) sh.getRange(r, ATT_COL.browser).setValue(payload.browser);

    return { status: 'ok', punchOutTime: fmtTime_(now), workingHours: hours.label, early: earlyInfo.isEarly, earlyBy: earlyInfo.label };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * EMPLOYEE: MY ATTENDANCE HISTORY (month/year)
 * ============================================================ */
function getMyAttendance(employeeId, month, year) {
  ensureSheets_();
  var sh = getSheet_(SHEET_ATT);
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var d = data[i][ATT_COL.date - 1];
    if (!(d instanceof Date)) continue;
    if (d.getMonth() !== month || d.getFullYear() !== year) continue;
    var eid = String(data[i][ATT_COL.empId - 1] || '').trim();
    if (eid.toLowerCase() !== String(employeeId || '').trim().toLowerCase()) continue;
    rows.push(mapAttRow_(data[i]));
  }
  rows.sort(function (a, b) { return a._sortDate - b._sortDate; });
  return { rows: rows, summary: buildSummary_(rows) };
}

function mapAttRow_(r) {
  var d = r[ATT_COL.date - 1];
  return {
    date: (d instanceof Date) ? fmtDate_(d) : String(d || ''),
    _sortDate: (d instanceof Date) ? d.getTime() : 0,
    empId: r[ATT_COL.empId - 1] || '',
    empName: r[ATT_COL.empName - 1] || '',
    punchIn: r[ATT_COL.punchIn - 1] ? fmtTime_(r[ATT_COL.punchIn - 1]) : '',
    punchOut: r[ATT_COL.punchOut - 1] ? fmtTime_(r[ATT_COL.punchOut - 1]) : '',
    workingHours: r[ATT_COL.workingHours - 1] || '',
    status: r[ATT_COL.status - 1] || '',
    latIn: r[ATT_COL.latIn - 1] || '', lngIn: r[ATT_COL.lngIn - 1] || '', addrIn: r[ATT_COL.addrIn - 1] || '', selfieIn: r[ATT_COL.selfieIn - 1] || '',
    latOut: r[ATT_COL.latOut - 1] || '', lngOut: r[ATT_COL.lngOut - 1] || '', addrOut: r[ATT_COL.addrOut - 1] || '', selfieOut: r[ATT_COL.selfieOut - 1] || '',
    lateIn: r[ATT_COL.lateIn - 1] || '', earlyOut: r[ATT_COL.earlyOut - 1] || '', remarks: r[ATT_COL.remarks - 1] || ''
  };
}

function buildSummary_(rows) {
  var present = 0, late = 0, early = 0, totalMinutes = 0, inMinutesSum = 0, inCount = 0, outMinutesSum = 0, outCount = 0;
  rows.forEach(function (r) {
    if (r.punchIn) present++;
    if (r.lateIn && /min/i.test(r.lateIn)) late++;
    if (r.earlyOut && /min/i.test(r.earlyOut)) early++;
    var m = /([\d.]+)\s*h/i.exec(r.workingHours || '');
    if (m) totalMinutes += parseFloat(m[1]) * 60;
    var pm = parseTimeToMinutes_(r.punchIn); if (pm != null) { inMinutesSum += pm; inCount++; }
    var qm = parseTimeToMinutes_(r.punchOut); if (qm != null) { outMinutesSum += qm; outCount++; }
  });
  return {
    present: present,
    absent: 0,
    late: late,
    early: early,
    totalDays: rows.length,
    workingHours: (totalMinutes / 60).toFixed(1),
    avgInTime: inCount ? minutesToTime_(inMinutesSum / inCount) : '--',
    avgOutTime: outCount ? minutesToTime_(outMinutesSum / outCount) : '--'
  };
}

function parseTimeToMinutes_(t) {
  if (!t) return null;
  var m = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i.exec(t.trim());
  if (!m) return null;
  var h = parseInt(m[1], 10) % 12;
  if (m[3].toUpperCase() === 'PM') h += 12;
  return h * 60 + parseInt(m[2], 10);
}
function minutesToTime_(mins) {
  mins = Math.round(mins);
  var h = Math.floor(mins / 60) % 24, m = mins % 60;
  var ampm = h >= 12 ? 'PM' : 'AM';
  var h12 = h % 12; if (h12 === 0) h12 = 12;
  return (h12 < 10 ? '0' + h12 : h12) + ':' + (m < 10 ? '0' + m : m) + ' ' + ampm;
}

/* ============================================================
 * ADMIN: LIST EMPLOYEES (for filter dropdown + Employees page)
 * ============================================================ */
function getEmployees() {
  ensureSheets_();
  var sh = getSheet_(SHEET_EMP);
  var data = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][EMP_COL.id - 1] || '').trim();
    if (!id) continue;
    list.push({
      id: id,
      name: data[i][EMP_COL.name - 1] || '',
      department: data[i][EMP_COL.department - 1] || '',
      role: data[i][EMP_COL.role - 1] || 'Employee',
      active: String(data[i][EMP_COL.active - 1] || '').trim().toUpperCase() === 'YES'
    });
  }
  return list;
}

/** Admin adds a new employee. payload = {id, name, department, role, password} */
function addEmployee(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSheets_();
    if (!payload) return { status: 'error', message: 'No data received.' };
    var id = String(payload.id || '').trim();
    var name = String(payload.name || '').trim();
    var password = String(payload.password || '').trim();
    if (!id || !name || !password) return { status: 'error', message: 'Employee ID, Name and Password are required.' };

    var sh = getSheet_(SHEET_EMP);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][EMP_COL.id - 1] || '').trim().toLowerCase() === id.toLowerCase()) {
        return { status: 'error', message: 'Employee ID already exists.' };
      }
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([[
      id, name, payload.department || '', payload.role === 'Admin' ? 'Admin' : 'Employee', password, 'YES'
    ]]);
    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * ADMIN: PUNCH ENTRIES (filterable table)
 * filters = { month, year, employeeId, date }  (month is 0-based, all optional/-1 = any)
 * ============================================================ */
function getAdminPunchEntries(filters) {
  ensureSheets_();
  filters = filters || {};
  var sh = getSheet_(SHEET_ATT);
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var d = data[i][ATT_COL.date - 1];
    if (!(d instanceof Date)) continue;
    if (filters.month != null && filters.month !== -1 && d.getMonth() !== filters.month) continue;
    if (filters.year != null && filters.year !== -1 && d.getFullYear() !== filters.year) continue;
    if (filters.employeeId) {
      var eid = String(data[i][ATT_COL.empId - 1] || '').trim();
      if (eid.toLowerCase() !== String(filters.employeeId).trim().toLowerCase()) continue;
    }
    if (filters.date) {
      if (fmtDateKey_(d) !== filters.date) continue;
    }
    rows.push(mapAttRow_(data[i]));
  }
  rows.sort(function (a, b) { return b._sortDate - a._sortDate; });
  return rows;
}
