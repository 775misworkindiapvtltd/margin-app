/**
 * SMART ATTENDANCE MANAGEMENT SYSTEM (Google Apps Script web app)
 * ------------------------------------------------------------------
 * Files in this Apps Script project:
 *   Code.gs        (this file — entry point + config)
 *   Database.gs    (sheet read/write, punch in/out logic)
 *   Utilities.gs   (selfie -> Drive, reverse geocode, late/early calc)
 *   Index.html     (root template, wires up all client HTML "modules")
 *   Style.html     (all CSS)
 *   Common.html    (shared JS: state, login screen, google.script.run bridge, render loop, clock)
 *   Capture.html   (shared selfie + GPS capture modal used for Punch In / Punch Out)
 *   Dashboard.html (employee shell: Punch In/Out screen + My Attendance history + Profile)
 *   Admin.html     (admin shell: left sidebar, Punch Entries, Employees, view-details modal)
 *
 * REQUIRED GOOGLE SHEET TABS (created automatically the first time doGet runs,
 * or you can run setupSheets() once from the Apps Script editor, or click the
 * "🕒 Attendance Setup" menu in the Sheet itself):
 *
 *   Employees  -> Employee ID | Name | Department | Role | Password | Active
 *                 Role must be exactly "Admin" or "Employee". Active must be YES/NO.
 *
 *   Attendance -> Date | Employee ID | Employee Name | Punch In Time | Punch Out Time |
 *                 Working Hours | Status | Latitude In | Longitude In | Address In |
 *                 Selfie In | Latitude Out | Longitude Out | Address Out | Selfie Out |
 *                 Device | Browser | Late In | Early Out | Remarks
 *
 *   Settings   -> Company Name | In Time | Out Time | Allow Late Time (mins)
 *                 Single active config row (row 2). Editable from the Admin ->
 *                 Settings page in the app, or directly in the sheet.
 *                 "Allow Late Time" is a grace period: punching in within this
 *                 many minutes after In Time still counts as On Time.
 *
 * DEPLOY: Deploy > New deployment > type "Web app" > Execute as "Me" >
 *         Who has access "Anyone" (or "Anyone with Google account") > Deploy.
 *
 * NOTE ON SELFIES: Selfies are uploaded to a Drive folder named
 * "Attendance Selfies" (auto-created) and the file is set to
 * "Anyone with link can view" so the images can be shown inside the web app
 * and to admins. If your Workspace policy blocks link-sharing, selfies will
 * still be saved but may not preview for other users — ask your Workspace
 * admin to allow link sharing for this app's Drive folder.
 */

/* ---- Global configuration (tweak these for your office) ---- */
var CONFIG = {
  officeStartTime: '10:00',   // Punch In after this time => "Late In"
  officeEndTime:   '19:00',   // Punch Out before this time => "Early Out"
  driveFolderName: 'Attendance Selfies',
  timeZone: Session.getScriptTimeZone()
};

function doGet(e) {
  try {
    ensureSheets_(); // auto-create Employees / Attendance / Settings tabs + sample rows if missing
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Smart Attendance Management System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // If ANYTHING fails before the HTML is returned (missing/misnamed file,
    // sheet permission issue, etc.) show a clear diagnostic page instead of
    // a blank white screen or Apps Script's generic error page.
    return HtmlService.createHtmlOutput(buildBootErrorPage_(err));
  }
}

/**
 * Loads an HTML file's content for <?!= include('X') ?>. If the file is
 * missing or misnamed (the #1 cause of a totally blank web app page), this
 * throws a clear error that doGet()'s catch block turns into a diagnostic
 * page — instead of Apps Script failing silently/blank.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    throw new Error(
      'Could not load the "' + filename + '" HTML file. ' +
      'In the Apps Script editor, check that a file named exactly "' + filename + '" ' +
      '(Apps Script auto-adds .html — do NOT type ".html" yourself when naming it, ' +
      'otherwise it becomes "' + filename + '.html.html" and cannot be found) exists ' +
      'and is not empty. Original error: ' + err.message
    );
  }
}

/** Builds a plain, dependency-free HTML error page (no include() calls) so it can never itself fail to render. */
function buildBootErrorPage_(err) {
  var msg = String(err && err.message || err || 'Unknown error');
  var safe = msg.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Attendance App — Setup Issue</title>' +
    '<style>' +
    'body{font-family:system-ui,-apple-system,Arial,sans-serif;background:#0F172A;color:#111827;margin:0;' +
    'min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;}' +
    '.box{background:#fff;max-width:560px;border-radius:16px;padding:28px 26px;box-shadow:0 24px 60px -12px rgba(0,0,0,.4);}' +
    'h1{font-size:18px;margin:0 0 12px;color:#B91C1C;}' +
    'p{font-size:13.5px;line-height:1.6;margin:0 0 12px;color:#374151;}' +
    '.err{background:#FEF2F2;color:#B91C1C;border-radius:10px;padding:12px 14px;font-size:12.5px;' +
    'font-family:monospace;white-space:pre-wrap;word-break:break-word;margin-bottom:14px;}' +
    'ol{font-size:12.5px;color:#374151;line-height:1.8;padding-left:20px;}' +
    'code{background:#F1F5F9;padding:2px 5px;border-radius:4px;}' +
    '</style></head><body>' +
    '<div class="box">' +
    '<h1>⚠️ The app could not start</h1>' +
    '<p>The server hit an error before it could render the page. This is why you saw a blank page instead of the login screen.</p>' +
    '<div class="err">' + safe + '</div>' +
    '<p><b>Most common causes (check in this order):</b></p>' +
    '<ol>' +
    '<li>An HTML file is missing or misnamed. In the Apps Script editor, file names should be exactly: <code>Index</code>, <code>Style</code>, <code>AppScript</code> (Apps Script auto-adds ".html" — do not type it yourself).</li>' +
    '<li>This script is not bound to a Google Sheet (Employees/Attendance/Settings tabs need a spreadsheet). Open it via <b>your Sheet → Extensions → Apps Script</b>, not via script.google.com directly.</li>' +
    '<li>The deployment is running an old/cached version. Go to <b>Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy</b>, then reload this URL.</li>' +
    '</ol>' +
    '<p style="margin-top:14px;"><button onclick="location.reload()" style="background:#2563EB;color:#fff;border:none;border-radius:8px;padding:10px 18px;font-size:13px;font-weight:600;cursor:pointer;">Reload Page</button></p>' +
    '</div></body></html>';
}

/**
 * Adds a custom menu ("🕒 Attendance Setup") to the Google Sheet's menu bar
 * every time the spreadsheet is opened, so the sheets can be created with a
 * simple click — no need to open the Apps Script editor at all.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🕒 Attendance Setup')
    .addItem('▶ Create / Repair Sheets Now', 'setupSheets')
    .addToUi();
}

/**
 * One-click setup: creates the Employees / Attendance tabs (with sample rows)
 * if they don't already exist, and shows a confirmation popup.
 * Can be triggered either from the "🕒 Attendance Setup" menu in the Sheet,
 * or by running it directly from the Apps Script editor (click ▶ Run).
 */
function setupSheets() {
  ensureSheets_();
  try {
    SpreadsheetApp.getUi().alert('✅ Done! "Employees", "Attendance" and "Settings" sheets are ready to use.');
  } catch (err) {
    // getUi() fails when run from the Apps Script editor without a UI context — that's fine, sheets are still created.
  }
}
