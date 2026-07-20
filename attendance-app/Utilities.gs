/**
 * UTILITIES.GS
 * Selfie storage (Drive), reverse geocoding, and late/early/working-hours helpers.
 */

/** Returns (creating if needed) the Drive folder used to store attendance selfies. */
function getSelfieFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.driveFolderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.driveFolderName);
}

/**
 * Saves a base64 selfie (data URL or raw base64) to Drive and returns a viewable URL.
 * type is 'IN' or 'OUT', used only for a friendly filename.
 */
function saveSelfieToDrive_(base64Data, employeeId, type) {
  try {
    if (!base64Data) return '';
    var commaIdx = base64Data.indexOf(',');
    var meta = commaIdx !== -1 ? base64Data.substring(0, commaIdx) : '';
    var rawBase64 = commaIdx !== -1 ? base64Data.substring(commaIdx + 1) : base64Data;
    var mime = 'image/jpeg';
    var m = /data:(image\/[a-zA-Z]+);base64/.exec(meta);
    if (m) mime = m[1];

    var bytes = Utilities.base64Decode(rawBase64);
    var ts = Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyyMMdd_HHmmss');
    var ext = mime.indexOf('png') !== -1 ? 'png' : 'jpg';
    var fileName = employeeId + '_' + type + '_' + ts + '.' + ext;
    var blob = Utilities.newBlob(bytes, mime, fileName);

    var folder = getSelfieFolder_();
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // Some Workspace domains block link-sharing; the file is still saved.
    }
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch (err) {
    // Never block a punch just because the selfie upload had an issue saving —
    // but selfie capture itself is still mandatory on the client before this is called.
    return '';
  }
}

/**
 * Converts latitude/longitude into a human-readable address using Apps Script's
 * built-in Maps geocoder (no separate API key needed inside Apps Script).
 */
function reverseGeocode_(lat, lng) {
  try {
    var response = Maps.newGeocoder().reverseGeocode(lat, lng);
    if (response && response.results && response.results.length > 0) {
      return response.results[0].formatted_address;
    }
  } catch (err) {
    // fall through to coordinate string
  }
  return 'Lat: ' + lat + ', Lng: ' + lng;
}

/** Returns {isLate, label} comparing punch-in time against CONFIG.officeStartTime. */
function computeLateIn_(punchInDate) {
  var startParts = CONFIG.officeStartTime.split(':');
  var startMinutes = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);
  var actualMinutes = punchInDate.getHours() * 60 + punchInDate.getMinutes();
  if (actualMinutes > startMinutes) {
    var diff = actualMinutes - startMinutes;
    return { isLate: true, label: diff + ' min late' };
  }
  return { isLate: false, label: 'On Time' };
}

/** Returns {isEarly, label} comparing punch-out time against CONFIG.officeEndTime. */
function computeEarlyOut_(punchOutDate) {
  var endParts = CONFIG.officeEndTime.split(':');
  var endMinutes = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);
  var actualMinutes = punchOutDate.getHours() * 60 + punchOutDate.getMinutes();
  if (actualMinutes < endMinutes) {
    var diff = endMinutes - actualMinutes;
    return { isEarly: true, label: diff + ' min early' };
  }
  return { isEarly: false, label: 'On Time' };
}

/** Returns {hours, label} working duration between punch in and punch out Date objects. */
function computeWorkingHours_(punchInDate, punchOutDate) {
  var ms = punchOutDate.getTime() - punchInDate.getTime();
  if (ms < 0) ms = 0;
  var totalMinutes = Math.round(ms / 60000);
  var h = Math.floor(totalMinutes / 60);
  var m = totalMinutes % 60;
  return { hours: totalMinutes / 60, label: h + 'h ' + m + 'm' };
}
