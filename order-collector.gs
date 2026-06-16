/**
 * Nosh7 order collector  (Google Apps Script)
 *
 * Turns a Google Sheet into the "one page" where every order lands.
 * Each submitted order becomes a new row. Self-healing: if the app ever sends
 * a new field, a new column is added automatically, so it never mismatches.
 *
 * SETUP (one time):
 *   1. Create a Google Sheet. Extensions > Apps Script.
 *   2. Paste this whole file. Deploy > New deployment > Web app.
 *        Execute as: Me        Who has access: Anyone
 *      Copy the /exec URL into index.html CONFIG.ORDER_WEBHOOK.
 *
 * UPDATING (when this code changes):
 *   Paste the new code, then Deploy > Manage deployments > edit (pencil)
 *   > Version: New version > Deploy. The /exec URL stays the same.
 */

// Preferred column order for a fresh sheet. Any extra keys are appended after these.
var PREFERRED = [
  "placedAt", "orderNo", "name", "phone", "address",
  "plan", "deliveries", "slot", "preference", "days", "startDate",
  "addons", "instructions", "total", "payment", "paymentId"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");
    var keys = Object.keys(data);

    var headers;
    if (sheet.getLastRow() === 0) {
      // fresh sheet: preferred keys first, then any others
      headers = PREFERRED.filter(function (k) { return keys.indexOf(k) >= 0; })
        .concat(keys.filter(function (k) { return PREFERRED.indexOf(k) < 0; }));
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      // existing sheet: read headers, append any new keys as new columns
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .filter(function (h) { return h !== "" && h !== null; });
      var added = keys.filter(function (k) { return headers.indexOf(k) < 0; });
      if (added.length) {
        added.forEach(function (k) { headers.push(k); });
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    var row = headers.map(function (h) { return data[h] !== undefined ? data[h] : ""; });
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Opening the /exec URL in a browser confirms it is deployed.
function doGet() {
  return ContentService.createTextOutput("Nosh7 order collector is live.");
}
