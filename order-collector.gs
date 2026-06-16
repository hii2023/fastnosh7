/**
 * Nosh7 order collector  (Google Apps Script)
 *
 * This turns a Google Sheet into the "one page" where every order lands.
 * Each submitted order becomes a new row, which you can then automate
 * (Zapier, Make, Apps Script triggers, or your own system) into your CRM.
 *
 * SETUP (one time, about 5 minutes):
 *   1. Create a new Google Sheet. Name the first tab "Orders".
 *   2. Extensions > Apps Script. Delete the sample code, paste this whole file.
 *   3. Click Deploy > New deployment > type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *      Deploy, authorise, and copy the Web app URL (ends in /exec).
 *   4. In index.html CONFIG, set:  ORDER_WEBHOOK: "PASTE_THE_/exec_URL"
 *   5. Place a test order. A new row should appear within a second or two.
 *
 * The header row is written automatically on the first order.
 */

var HEADERS = [
  "placedAt", "orderNo", "name", "phone", "house", "building", "address",
  "plan", "deliveries", "slot", "preference", "days", "startDate",
  "addons", "instructions", "total", "payment", "paymentId"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // avoid two orders writing the same row
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders")
             || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Orders");

    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    var row = HEADERS.map(function (k) { return data[k] !== undefined ? data[k] : ""; });
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Lets you open the /exec URL in a browser to confirm it is deployed.
function doGet() {
  return ContentService.createTextOutput("Nosh7 order collector is live.");
}
