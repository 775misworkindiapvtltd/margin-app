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
  ensureSheets_(); // auto-create Employees / Attendance tabs + sample rows if missing
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Smart Attendance Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
