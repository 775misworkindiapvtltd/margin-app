/**
 * TALLY DASHBOARD — Google Apps Script web app
 * =============================================================
 * Files in this Apps Script project:
 *   Code.gs        (this file — server side)
 *   Index.html     (root HTML shell)
 *   CSS.html       (all styles — desktop + mobile responsive)
 *   Common.html    (state, login, shell/nav, generic grid engine, helpers)
 *   Dashboard.html (home page summary cards)
 *   Grids.html     (Expense / Payables / Receivables / Receipt / Payment grids)
 *   Overdue.html   (ageing analysis — Receivables & Payables buckets)
 *   Ledger.html    (party ledger statement + printable/downloadable PDF view)
 *
 * REQUIRED GOOGLE SHEET TABS (names must match EXACTLY):
 *
 *   LOGIN PAGE  -> NAME | ID | PASSWORD | DASHBOARD | EXPENSE | PAYABLES |
 *                  RECEIVABLES | RECEIPT | PAYMENT | LEDGER | OVERDUE
 *                  (mark YES in a column to grant that user the page)
 *
 *   EXPENSE     -> TIMESTAMP | DATE | VOUCHER NUMBER | PARTY NAME | Group |
 *                  Sub_Group | DESIGN NUMBER | ITEM NAME | QTY | RATE | AMOUNT | TYPE
 *
 *   PAYABLES    -> TIMESTAMP | Bill_Date | Bill_Ref_No | Party_Name | Party_Group |
 *                  Sub_Group | Main_Group | Sales Person | Phone No. | Payment Terms |
 *                  Closing_Balance | Due_Date | Overdue_Days
 *
 *   RECEIVABLES -> TIMESTAMP | Bill_Date | Bill_Ref_No | Party_Name | Party_Group |
 *                  Sub_Group | Main_Group | Sales Person | Phone No. | Payment Terms |
 *                  Pending Amount | Due_Date | Overdue_Days
 *
 *   Receipt     -> Timestamp | Voucher_Number | Date | Group | Sub_Group |
 *                  Ledger_Name(Cr) | Ledger_Amount(Cr) | Bill_Type | Bill_Name | Bill_Amount |
 *                  Type | Ledger_Name(Dr) | Ledger_Amount(Dr) | Bank_Party_Name |
 *                  Transaction_Type | Inst_No | Inst_Date | Bank_Name
 *                  (read POSITIONALLY — header repeats "Ledger_Name"/"Ledger_Amount" twice)
 *
 *   PAYMENT     -> same 18 columns/layout as Receipt (positional read)
 *
 *   Balance     -> Timestamp | Name | Alias | Subgroup | Group | Main Group | GUID |
 *                  Master ID | Alter ID | Bill By Bill | Credit Period | Credit Limit |
 *                  State | Country | Pincode | Email | Email CC | Contact Person | Mobile |
 *                  Opening Date | Opening Particular | Opening Debit | Opening Credit |
 *                  Opening Balance | Voucher Date | Voucher Particular | Voucher Type |
 *                  Voucher No | Voucher Debit | Voucher Credit | Closing Date |
 *                  Closing Particular | Closing Debit | Closing Credit | Closing Balance
 *                  (this is the party ledger source — one row per voucher line per party,
 *                  Opening filled on the first row for a party, Closing filled on the last)
 *
 * NOTE ON ASSUMPTIONS: The original screenshots/PDF-formula attachments referenced in
 * chat did not come through to this tool, so the visual layout below (dashboard cards,
 * grid columns, ledger PDF layout) follows the standard/common Tally ledger format and
 * the exact sheet headers you pasted as text. Everything is easy to re-skin once you
 * share the actual screenshots — the data layer (Code.gs) will not need to change.
 *
 * Deploy: Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access
 * "Anyone" > Deploy.
 */

/* ---- Sheet name constants ---- */
var SHEETS = {
  login:       'LOGIN PAGE',
  expense:     'EXPENSE',
  payables:    'PAYABLES',
  receivables: 'RECEIVABLES',
  receipt:     'Receipt',
  payment:     'PAYMENT',
  balance:     'Balance'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Tally Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================= GENERIC SHEET READ ============================= */
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

// Positional read (used for Receipt / PAYMENT where header names repeat, e.g. "Ledger_Name"
// appears twice — once for the Credit ledger block, once for the Debit ledger block).
function sheetToRows_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(function (row) { return row.some(function (c) { return c !== ''; }); });
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
function numOrZero_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = parseFloat(String(v).replace(/[₹,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
function isYes_(v) { return String(v || '').trim().toUpperCase() === 'YES'; }

// Tolerant header lookup: exact match first, then case/whitespace-insensitive fallback.
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

/* ============================= ROW MAPPERS ============================= */
function mapUser_(r) {
  return {
    name: r['NAME'] || '', id: String(r['ID'] || '').trim(), password: String(r['PASSWORD'] || '').trim(),
    dashboard:   isYes_(pick_(r, ['DASHBOARD'])),
    expense:     isYes_(pick_(r, ['EXPENSE'])),
    payables:    isYes_(pick_(r, ['PAYABLES'])),
    receivables: isYes_(pick_(r, ['RECEIVABLES'])),
    receipt:     isYes_(pick_(r, ['RECEIPT'])),
    payment:     isYes_(pick_(r, ['PAYMENT'])),
    ledger:      isYes_(pick_(r, ['LEDGER'])),
    overdue:     isYes_(pick_(r, ['OVERDUE']))
  };
}

function mapExpense_(rows) {
  return rows.map(function (r) {
    return {
      timestamp: fmtTimestamp_(r['TIMESTAMP']), date: fmtDateOnly_(r['DATE']),
      voucherNumber: fmtValue_(r['VOUCHER NUMBER']), partyName: r['PARTY NAME'] || '',
      group: r['Group'] || '', subGroup: r['Sub_Group'] || '',
      designNumber: fmtValue_(r['DESIGN NUMBER']), itemName: fmtValue_(r['ITEM NAME']),
      qty: fmtValue_(r['QTY']), rate: fmtValue_(r['RATE']), amount: numOrZero_(r['AMOUNT']),
      type: r['TYPE'] || ''
    };
  });
}

function mapPayables_(rows) {
  return rows.map(function (r) {
    return {
      timestamp: fmtTimestamp_(r['TIMESTAMP']), billDate: fmtDateOnly_(r['Bill_Date']),
      billRefNo: fmtValue_(r['Bill_Ref_No']), partyName: r['Party_Name'] || '',
      partyGroup: r['Party_Group'] || '', subGroup: r['Sub_Group'] || '', mainGroup: r['Main_Group'] || '',
      salesPerson: fmtValue_(pick_(r, ['Sales Person'])), phoneNo: fmtValue_(pick_(r, ['Phone No.', 'Phone No'])),
      paymentTerms: fmtValue_(r['Payment Terms']), closingBalance: numOrZero_(r['Closing_Balance']),
      dueDate: fmtDateOnly_(r['Due_Date']), overdueDays: numOrZero_(r['Overdue_Days'])
    };
  });
}

function mapReceivables_(rows) {
  return rows.map(function (r) {
    return {
      timestamp: fmtTimestamp_(r['TIMESTAMP']), billDate: fmtDateOnly_(r['Bill_Date']),
      billRefNo: fmtValue_(r['Bill_Ref_No']), partyName: r['Party_Name'] || '',
      partyGroup: r['Party_Group'] || '', subGroup: r['Sub_Group'] || '', mainGroup: r['Main_Group'] || '',
      salesPerson: fmtValue_(pick_(r, ['Sales Person'])), phoneNo: fmtValue_(pick_(r, ['Phone No.', 'Phone No'])),
      paymentTerms: fmtValue_(r['Payment Terms']), pendingAmount: numOrZero_(pick_(r, ['Pending Amount'])),
      dueDate: fmtDateOnly_(r['Due_Date']), overdueDays: numOrZero_(r['Overdue_Days'])
    };
  });
}

// Positional mapper for Receipt / PAYMENT (18 columns, header names repeat so we can't
// key off header text reliably — read by column index instead).
function mapReceiptRows_(rows) {
  return rows.map(function (row) {
    return {
      timestamp: fmtTimestamp_(row[0]), voucherNumber: fmtValue_(row[1]), date: fmtDateOnly_(row[2]),
      group: fmtValue_(row[3]), subGroup: fmtValue_(row[4]),
      creditLedgerName: fmtValue_(row[5]), creditLedgerAmount: numOrZero_(row[6]),
      billType: fmtValue_(row[7]), billName: fmtValue_(row[8]), billAmount: numOrZero_(row[9]),
      type: fmtValue_(row[10]),
      debitLedgerName: fmtValue_(row[11]), debitLedgerAmount: numOrZero_(row[12]),
      bankPartyName: fmtValue_(row[13]), transactionType: fmtValue_(row[14]),
      instNo: fmtValue_(row[15]), instDate: fmtDateOnly_(row[16]), bankName: fmtValue_(row[17])
    };
  });
}

var BALANCE_COLS = [
  'timestamp','name','alias','subgroup','group','mainGroup','guid','masterId','alterId',
  'billByBill','creditPeriod','creditLimit','state','country','pincode','email','emailCC',
  'contactPerson','mobile','openingDate','openingParticular','openingDebit','openingCredit',
  'openingBalance','voucherDate','voucherParticular','voucherType','voucherNo','voucherDebit',
  'voucherCredit','closingDate','closingParticular','closingDebit','closingCredit','closingBalance'
];

// Positional mapper for the Balance/Ledger source tab (35 columns, exact order given by user).
function mapBalanceRows_(rows) {
  return rows.map(function (row) {
    var o = {};
    BALANCE_COLS.forEach(function (key, i) {
      var v = row[i];
      if (key === 'timestamp') v = fmtTimestamp_(v);
      else if (key === 'openingDate' || key === 'voucherDate' || key === 'closingDate') v = fmtDateOnly_(v);
      else if (key === 'openingDebit' || key === 'openingCredit' || key === 'openingBalance' ||
               key === 'voucherDebit' || key === 'voucherCredit' ||
               key === 'closingDebit' || key === 'closingCredit' || key === 'closingBalance') v = numOrZero_(v);
      else v = fmtValue_(v);
      o[key] = v;
    });
    return o;
  });
}

/* ============================= LOGIN / BOOTSTRAP ============================= */
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loadAll = !perms;
  var needExpense     = loadAll || perms.dashboard || perms.expense;
  var needPayables     = loadAll || perms.dashboard || perms.payables || perms.overdue;
  var needReceivables   = loadAll || perms.dashboard || perms.receivables || perms.overdue;
  var needReceipt       = loadAll || perms.dashboard || perms.receipt;
  var needPayment      = loadAll || perms.dashboard || perms.payment;
  var needBalance      = loadAll || perms.dashboard || perms.ledger;

  var result = {
    expense:     needExpense     ? mapExpense_(sheetToObjects_(SHEETS.expense))         : [],
    payables:    needPayables    ? mapPayables_(sheetToObjects_(SHEETS.payables))        : [],
    receivables: needReceivables ? mapReceivables_(sheetToObjects_(SHEETS.receivables))  : [],
    receipt:     needReceipt     ? mapReceiptRows_(sheetToRows_(SHEETS.receipt))         : [],
    payment:     needPayment     ? mapReceiptRows_(sheetToRows_(SHEETS.payment))         : [],
    balanceParties: needBalance  ? getLedgerPartyList_()                                  : [],
    users:       sheetToObjects_(SHEETS.login).map(mapUser_),
    missingSheets: []
  };

  var have = {};
  ss.getSheets().forEach(function (s) { have[s.getName().trim().toUpperCase()] = true; });
  Object.keys(SHEETS).forEach(function (k) {
    if (!have[SHEETS[k].toUpperCase()]) result.missingSheets.push(SHEETS[k]);
  });

  return result;
}

/* ============================= LEDGER (Balance tab) ============================= */

// Lightweight party directory for the Ledger picker (name + group + mobile + running closing).
function getLedgerPartyList_() {
  var rows = mapBalanceRows_(sheetToRows_(SHEETS.balance));
  var byName = {};
  rows.forEach(function (r) {
    if (!r.name) return;
    if (!byName[r.name]) {
      byName[r.name] = {
        name: r.name, subgroup: r.subgroup, group: r.group, mobile: r.mobile,
        state: r.state, closingBalance: 0, hasClosing: false
      };
    }
    if (r.closingParticular) { byName[r.name].closingBalance = r.closingDebit - r.closingCredit; byName[r.name].hasClosing = true; }
  });
  return Object.keys(byName).sort().map(function (n) { return byName[n]; });
}

/**
 * Builds the full statement for one party from the Balance tab.
 * Returns { header:{...party info...}, opening:{date,particular,debit,credit,balance},
 *           entries:[{date,particular,vchType,vchNo,debit,credit,balance}],
 *           closing:{date,particular,debit,credit,balance} }
 */
function getLedgerForParty(name) {
  var rows = mapBalanceRows_(sheetToRows_(SHEETS.balance)).filter(function (r) { return r.name === name; });
  if (!rows.length) return null;

  var first = rows[0];
  var header = {
    name: first.name, alias: first.alias, subgroup: first.subgroup, group: first.group,
    mainGroup: first.mainGroup, state: first.state, country: first.country, pincode: first.pincode,
    email: first.email, emailCC: first.emailCC, contactPerson: first.contactPerson, mobile: first.mobile,
    creditPeriod: first.creditPeriod, creditLimit: first.creditLimit, billByBill: first.billByBill,
    address: first.guid /* GUID column often holds the registered address text in this export */
  };

  // Opening: first row that actually carries an opening particular.
  var openingRow = rows.find(function (r) { return r.openingParticular; }) || first;
  var opening = {
    date: openingRow.openingDate, particular: openingRow.openingParticular || 'Opening Balance',
    debit: openingRow.openingDebit, credit: openingRow.openingCredit, balance: openingRow.openingBalance
  };

  var running = opening.debit - opening.credit;
  var entries = [];
  rows.forEach(function (r) {
    if (!r.voucherDate && !r.voucherParticular && !r.voucherNo) return;
    running += (r.voucherDebit - r.voucherCredit);
    entries.push({
      date: r.voucherDate, particular: r.voucherParticular, vchType: r.voucherType, vchNo: r.voucherNo,
      debit: r.voucherDebit, credit: r.voucherCredit, balance: running
    });
  });

  var closingRow = rows.slice().reverse().find(function (r) { return r.closingParticular; }) || rows[rows.length - 1];
  var closing = {
    date: closingRow.closingDate, particular: closingRow.closingParticular || 'Closing Balance',
    debit: closingRow.closingDebit, credit: closingRow.closingCredit,
    balance: (closingRow.closingDebit || closingRow.closingCredit) ? (closingRow.closingDebit - closingRow.closingCredit) : running
  };

  return { header: header, opening: opening, entries: entries, closing: closing };
}
