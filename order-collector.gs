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
  "placedAt", "orderNo", "status", "verified", "name", "phone", "address",
  "plan", "deliveries", "slot", "preference", "days", "startDate",
  "addons", "instructions", "total", "payment", "paymentId", "reason"
];

// HMAC-SHA256 as lowercase hex (matches the Cloudflare worker's hmacHex).
function hmacHex_(secret, message) {
  var raw = Utilities.computeHmacSha256Signature(message, secret);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

// "yes" if a paid order carries a valid ticket, "UNVERIFIED" if paid but the ticket
// is missing/forged, "" for non-paid rows. The ticket secret lives only in Script
// Properties (Project Settings > Script Properties: ORDER_TICKET_SECRET) and the
// worker, never in the browser, so a fake "paid" row cannot carry a valid ticket.
function verifyTicket_(data) {
  if (String(data.status) !== "paid") return "";
  var secret = PropertiesService.getScriptProperties().getProperty("ORDER_TICKET_SECRET");
  if (!secret || !data.ticket) return "UNVERIFIED";
  var expect = hmacHex_(secret, (data.orderNo || "") + "|" + (data.paymentId || ""));
  return expect === data.ticket ? "yes" : "UNVERIFIED";
}

// Run this ONCE from the editor to clean up: pick "resetSheet" in the function
// dropdown at the top, then click Run. It wipes all rows and writes a fresh,
// correctly ordered header row. (status = paid / failed / submitted.)
function resetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");
  sheet.clear();
  sheet.getRange(1, 1, 1, PREFERRED.length).setValues([PREFERRED]);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    // authenticate paid orders against the ticket, then drop the raw ticket
    data.verified = verifyTicket_(data);
    delete data.ticket;
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
