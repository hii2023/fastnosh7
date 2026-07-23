/**
 * Nosh7 order collector  (Google Apps Script)
 *
 * Every submitted order becomes one row in the "Orders" sheet, in a clean,
 * fixed column order with readable headers. Self-healing: if the app ever sends
 * a brand-new field, a new column is appended automatically.
 *
 * SETUP (one time):
 *   1. Open the sheet. Extensions > Apps Script. Paste this whole file. Save.
 *   2. Project Settings (gear) > Script Properties >
 *        ORDER_TICKET_SECRET = (same value set in the Cloudflare worker)
 *   3. Deploy > New deployment > Web app.  Execute as: Me   Who has access: Anyone
 *      Copy the /exec URL into index.html CONFIG.ORDER_WEBHOOK.
 *   4. From the editor function dropdown pick "setupSheet" and click Run.
 *      (Reorders + reformats existing rows; keeps all your data.)
 *
 * WHATSAPP ORDER ALERTS via CallMeBot (optional, free):
 *   1. On the owner phone (9712989498) save +34 644 44 21 48 as a contact and
 *      send it this WhatsApp message: "I allow callmebot to send me messages".
 *      It replies with your personal apikey within a couple of minutes.
 *   2. Script Properties > add:
 *        CALLMEBOT_APIKEY = (the key from step 1)
 *        OWNER_WHATSAPP   = 919712989498
 *   3. Deploy a new version. Every paid order now also pings the owner phone.
 *      Run "testWhatsAppAlert" from the editor dropdown to send a test message.
 *
 * UPDATING (when this code changes):
 *   Paste new code > Save > Deploy > Manage deployments > edit (pencil) >
 *   Version: New version > Deploy. The /exec URL stays the same.
 */

// Column order + friendly headers. key = field the app sends, label = sheet header.
var COLUMNS = [
  { key: "placedAt",           label: "Placed At" },
  { key: "orderNo",            label: "Order No" },
  { key: "status",             label: "Status" },
  { key: "verified",           label: "Verified" },
  { key: "name",               label: "Name" },
  { key: "phone",              label: "Phone" },
  { key: "house",              label: "House / Flat" },
  { key: "building",           label: "Building / Apartment" },
  { key: "area",               label: "Area" },
  { key: "pincode",            label: "Pincode" },
  { key: "plan",               label: "Plan" },
  { key: "deliveries",         label: "Meals" },
  { key: "slot",               label: "Slot" },
  { key: "preference",         label: "Preference" },
  { key: "days",               label: "Days" },
  { key: "startDate",          label: "Start Date" },
  { key: "addons",             label: "Add-ons" },
  { key: "instructions",       label: "Instructions" },
  { key: "deliveryFeePerMeal", label: "Distance Fee / Delivery" },
  { key: "deliveryFeeTotal",   label: "Distance Fee Total" },
  { key: "total",              label: "Total" },
  { key: "payment",            label: "Payment" },
  { key: "paymentId",          label: "Payment ID" },
  { key: "address",            label: "Full Address" },
  { key: "map",                label: "Map (GPS)" },
  { key: "distanceKm",         label: "Distance (km)" },
  { key: "lat",                label: "Lat" },
  { key: "lng",                label: "Lng" },
  { key: "reason",             label: "Reason / Note" }
];

var MONEY_KEYS  = ["deliveryFeePerMeal", "deliveryFeeTotal", "total"];
var HIDDEN_KEYS = [];   // GPS Lat/Lng kept visible for location visibility

// Build a clickable Google Maps link from coordinates (shown as "Open in Maps").
function mapsLink_(lat, lng) {
  if (lat === undefined || lat === null || lat === "" ||
      lng === undefined || lng === null || lng === "") return "";
  return '=HYPERLINK("https://www.google.com/maps?q=' + lat + ',' + lng + '","Open in Maps")';
}

function labelFor_(key) {
  for (var i = 0; i < COLUMNS.length; i++) if (COLUMNS[i].key === key) return COLUMNS[i].label;
  return key;
}
// Resolve a header cell (which may be a friendly label OR a raw key) back to its field key.
function keyForHeader_(header) {
  var h = String(header);
  for (var i = 0; i < COLUMNS.length; i++) if (COLUMNS[i].label === h || COLUMNS[i].key === h) return COLUMNS[i].key;
  return h;
}

// ---- ticket verification (authenticity of "paid" rows) ----
function hmacHex_(secret, message) {
  var raw = Utilities.computeHmacSha256Signature(message, secret);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}
function verifyTicket_(data) {
  if (String(data.status) !== "paid") return "";
  var secret = PropertiesService.getScriptProperties().getProperty("ORDER_TICKET_SECRET");
  if (!secret || !data.ticket) return "UNVERIFIED";
  var expect = hmacHex_(secret, (data.orderNo || "") + "|" + (data.paymentId || ""));
  return expect === data.ticket ? "yes" : "UNVERIFIED";
}

// ---- incoming orders ----
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    data.verified = verifyTicket_(data);   // authenticate before storing
    delete data.ticket;                     // never store the raw ticket
    data.map = mapsLink_(data.lat, data.lng); // clickable map link from GPS coords
    var keys = Object.keys(data);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");

    var headerKeys;
    if (sheet.getLastRow() === 0) {
      // fresh sheet: full defined order, plus any unexpected keys at the end
      headerKeys = COLUMNS.map(function (c) { return c.key; });
      keys.forEach(function (k) { if (headerKeys.indexOf(k) < 0) headerKeys.push(k); });
      sheet.getRange(1, 1, 1, headerKeys.length).setValues([headerKeys.map(labelFor_)]);
      formatSheet_(sheet, headerKeys);
    } else {
      // existing sheet: map current headers to keys, append any genuinely new keys
      var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .filter(function (h) { return h !== "" && h !== null; });
      headerKeys = headerRow.map(keyForHeader_);
      var added = keys.filter(function (k) { return headerKeys.indexOf(k) < 0; });
      if (added.length) {
        added.forEach(function (k) { headerKeys.push(k); });
        sheet.getRange(1, 1, 1, headerKeys.length).setValues([headerKeys.map(labelFor_)]);
      }
    }

    var row = headerKeys.map(function (k) { return data[k] !== undefined ? data[k] : ""; });

    // De-dupe by Order No: exactly one row per order. The same order legitimately arrives
    // more than once — a "pending" pre-write before payment, then "paid" after, plus
    // possible beacon/fetch retries and the server-side Razorpay webhook. Each message
    // MERGES into the row: a field the message does not supply keeps its existing value.
    // This is what lets a partial update (the webhook flipping status to "paid") land
    // without blanking the address/name the pre-write already saved.
    var isNew = true;
    var prevStatus = "";
    var record = data;                          // what we alert from (the full stored row)
    var noCol = headerKeys.indexOf("orderNo");  // 0-based
    var statusCol = headerKeys.indexOf("status");
    if (noCol >= 0 && data.orderNo) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var existing = sheet.getRange(2, noCol + 1, lastRow - 1, 1).getValues();
        for (var r = 0; r < existing.length; r++) {
          if (String(existing[r][0]) === String(data.orderNo)) {
            var rowNum = r + 2;
            var current = sheet.getRange(rowNum, 1, 1, headerKeys.length).getValues()[0];
            if (statusCol >= 0) prevStatus = String(current[statusCol] || "");
            var merged = headerKeys.map(function (k, i) {
              var incoming = data[k];
              return (incoming !== undefined && incoming !== "" && incoming !== null) ? incoming : current[i];
            });
            sheet.getRange(rowNum, 1, 1, merged.length).setValues([merged]);
            record = {};
            headerKeys.forEach(function (k, i) { record[k] = merged[i]; });
            isNew = false;
            break;
          }
        }
      }
    }
    if (isNew) sheet.appendRow(row);

    // WhatsApp alert to the owner, exactly once — when the order FIRST reaches a real
    // paid/submitted state, whether that is the first insert or an update flipping a
    // pending pre-write to paid. Built from the full stored row so even a bare webhook
    // update alerts with the customer's name/address. Isolated: an alert failure can
    // never lose an order.
    var nowStatus     = String(data.status || "");
    var isRealNow     = (nowStatus === "paid" || nowStatus === "submitted");
    var wasRealBefore = (prevStatus === "paid" || prevStatus === "submitted");
    if (isRealNow && !wasRealBefore) {
      try { sendWhatsAppAlert_(record); } catch (waErr) { /* order is already saved; ignore */ }
    }

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

// ---- WhatsApp owner alert (CallMeBot) ----
// Sends one message per real order (paid or UPI-submitted, never failed) to the
// owner phone. Needs Script Properties CALLMEBOT_APIKEY + OWNER_WHATSAPP.
function sendWhatsAppAlert_(data) {
  var status = String(data.status || "");
  if (status !== "paid" && status !== "submitted") return;

  var props  = PropertiesService.getScriptProperties();
  var apikey = props.getProperty("CALLMEBOT_APIKEY");
  var phone  = props.getProperty("OWNER_WHATSAPP");
  if (!apikey || !phone) return;   // alerts not configured; sheet still works

  var payLine;
  if (status === "paid") {
    payLine = (data.verified === "yes")
      ? "Rs " + data.total + "  PAID (verified)"
      : "Rs " + data.total + "  PAID but UNVERIFIED, check the sheet";
  } else {
    payLine = "Rs " + data.total + "  UPI submitted, confirm payment manually";
  }

  var lines = [
    "NEW ORDER " + (data.orderNo || ""),
    (data.name || "") + " | " + (data.phone || ""),
    (data.plan || "") + (data.addons ? " + " + data.addons : ""),
    payLine,
    "Slot: " + (data.slot || "") + " | Start: " + (data.startDate || ""),
    "Days: " + (data.days || "") + " | " + (data.preference || ""),
    "Addr: " + (data.address || "") + (data.pincode ? " (" + data.pincode + ")" : "")
  ];
  if (data.distanceKm) {
    lines.push("Distance: " + data.distanceKm + " km" +
      (Number(data.deliveryFeeTotal) > 0 ? " | Fee Rs " + data.deliveryFeeTotal : " | Free zone"));
  }
  if (data.lat && data.lng) lines.push("https://www.google.com/maps?q=" + data.lat + "," + data.lng);
  if (data.instructions) lines.push("Note: " + data.instructions);

  var url = "https://api.callmebot.com/whatsapp.php" +
    "?phone="  + encodeURIComponent(phone) +
    "&apikey=" + encodeURIComponent(apikey) +
    "&text="   + encodeURIComponent(lines.join("\n"));
  UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
}

// Run this from the editor dropdown after setting the Script Properties:
// it sends a sample order alert to the owner phone so you can verify the setup.
function testWhatsAppAlert() {
  sendWhatsAppAlert_({
    status: "paid", verified: "yes", orderNo: "NOSH-TEST01",
    name: "Test Customer", phone: "9999999999",
    plan: "Healthy Fresh Monthly", addons: "Healthy drink",
    total: 6224, slot: "Lunch 11:30 AM to 1:30 PM", startDate: "Mon, 7 Jul",
    days: "Mon to Sat", preference: "Regular",
    address: "B-204, Shivalik Shilp, Iscon Cross Road", pincode: "380015",
    distanceKm: "6.40", deliveryFeeTotal: 250, lat: 23.03, lng: 72.51,
    instructions: "Call on arrival"
  });
}

// ===========================================================================
// Run "setupSheet" ONCE from the editor to reorder existing rows into the new
// clean column order and apply all formatting. It KEEPS your existing data.
// ===========================================================================
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");

  var values = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  var oldHeader = values.length ? values[0] : [];
  var oldKeys = oldHeader.map(keyForHeader_);

  // target order: all defined columns, then any extra keys found in the old sheet
  var targetKeys = COLUMNS.map(function (c) { return c.key; });
  oldKeys.forEach(function (k) { if (k && targetKeys.indexOf(k) < 0) targetKeys.push(k); });

  // rebuild every row in the new order, mapping by key
  var out = [targetKeys.map(labelFor_)];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < oldHeader.length; c++) obj[oldKeys[c]] = values[r][c];
    obj.map = mapsLink_(obj.lat, obj.lng); // backfill map link for existing rows
    out.push(targetKeys.map(function (k) { return obj[k] !== undefined ? obj[k] : ""; }));
  }

  sheet.clear();
  sheet.getRange(1, 1, out.length, targetKeys.length).setValues(out);
  formatSheet_(sheet, targetKeys);
  SpreadsheetApp.getActiveSpreadsheet().toast("Orders sheet reformatted (" + (out.length - 1) + " rows kept).", "Done", 5);
}

// Apply header styling, freezing, currency formats, colour rules, hidden cols.
function formatSheet_(sheet, keys) {
  var nCols = keys.length;
  var maxRows = sheet.getMaxRows();
  var idx = function (k) { return keys.indexOf(k) + 1; }; // 1-based, 0 if absent

  // header band
  sheet.getRange(1, 1, 1, nCols)
    .setFontWeight("bold").setFontColor("#ffffff").setBackground("#1f7a3d")
    .setVerticalAlignment("middle").setHorizontalAlignment("left").setWrap(true);
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(Math.min(3, nCols)); // freeze Placed At / Order No / Status

  // money columns -> "Rs 1,250"
  MONEY_KEYS.forEach(function (k) {
    var i = idx(k);
    if (i > 0) sheet.getRange(2, i, Math.max(maxRows - 1, 1), 1).setNumberFormat('"Rs "#,##0');
  });

  // colour rules for Verified + Status
  var rules = [];
  function colRange(k) { var i = idx(k); return i > 0 ? sheet.getRange(2, i, Math.max(maxRows - 1, 1), 1) : null; }
  var vR = colRange("verified");
  if (vR) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("yes")
      .setBackground("#d9ead3").setFontColor("#137333").setBold(true).setRanges([vR]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("UNVERIFIED")
      .setBackground("#f4cccc").setFontColor("#a50e0e").setBold(true).setRanges([vR]).build());
  }
  var sR = colRange("status");
  if (sR) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("paid")
      .setFontColor("#137333").setBold(true).setRanges([sR]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("failed")
      .setFontColor("#a50e0e").setRanges([sR]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("submitted")
      .setFontColor("#7f6000").setRanges([sR]).build());
  }
  sheet.setConditionalFormatRules(rules);

  // tidy widths, then hide technical columns
  sheet.autoResizeColumns(1, nCols);
  HIDDEN_KEYS.forEach(function (k) { var i = idx(k); if (i > 0) sheet.hideColumns(i); });

  // cap a few that auto-size too wide
  var caps = { address: 240, area: 200, instructions: 200, days: 150, slot: 150 };
  Object.keys(caps).forEach(function (k) { var i = idx(k); if (i > 0) sheet.setColumnWidth(i, caps[k]); });
}
