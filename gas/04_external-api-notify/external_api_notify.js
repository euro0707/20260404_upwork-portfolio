// ============================================================
// 04: External API Notify
//   Google Sheets trigger → Slack / LINE / Generic Webhook
// ============================================================
// Setup:
//   1. Set your webhook URLs in the CONFIG section below
//   2. Run testSlack() / testLine() / testWebhook() to verify
//   3. Run setupTrigger() to fire automatically on sheet edits
// ============================================================

// ----------------------------------------------------------------
// CONFIG — set your endpoints here
// ----------------------------------------------------------------
var CONFIG = {
  // Slack: create at https://api.slack.com/apps → Incoming Webhooks
  slack_webhook_url: "YOUR_SLACK_WEBHOOK_URL",

  // LINE Notify: get token at https://notify-bot.line.me/my/
  line_notify_token: "YOUR_LINE_NOTIFY_TOKEN",

  // Generic webhook (e.g. webhook.site, Zapier, Make, n8n, etc.)
  generic_webhook_url: "https://webhook.site/e3112e00-699e-49c4-b235-b5ac16df8426",

  // Spreadsheet to monitor
  spreadsheet_id: "YOUR_SPREADSHEET_ID",
  sheet_name:     "Notifications"
};

// ----------------------------------------------------------------
// notifyAll: send to all configured channels at once
// ----------------------------------------------------------------
function notifyAll(message, data) {
  var results = [];

  if (CONFIG.slack_webhook_url !== "YOUR_SLACK_WEBHOOK_URL") {
    results.push("Slack: " + sendSlack(message, data));
  }
  if (CONFIG.line_notify_token !== "YOUR_LINE_NOTIFY_TOKEN") {
    results.push("LINE: " + sendLine(message));
  }
  if (CONFIG.generic_webhook_url !== "YOUR_GENERIC_WEBHOOK_URL") {
    results.push("Webhook: " + sendGenericWebhook(message, data));
  }

  Logger.log(results.join(" | "));
  return results;
}

// ----------------------------------------------------------------
// sendSlack: post a message to Slack via Incoming Webhook
// ----------------------------------------------------------------
function sendSlack(message, data) {
  var payload = {
    text: message,
    attachments: data ? [{
      color: "#36a64f",
      fields: Object.keys(data).map(function(key) {
        return { title: key, value: String(data[key]), short: true };
      })
    }] : []
  };

  var response = UrlFetchApp.fetch(CONFIG.slack_webhook_url, {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return response.getResponseCode() === 200 ? "OK" : "Error " + response.getResponseCode();
}

// ----------------------------------------------------------------
// sendLine: send a LINE Notify message
// ----------------------------------------------------------------
function sendLine(message) {
  var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
    method:  "post",
    headers: { "Authorization": "Bearer " + CONFIG.line_notify_token },
    payload: { message: message },
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  return result.status === 200 ? "OK" : "Error: " + result.message;
}

// ----------------------------------------------------------------
// sendGenericWebhook: POST JSON to any webhook endpoint
// ----------------------------------------------------------------
function sendGenericWebhook(message, data) {
  var payload = {
    message:   message,
    timestamp: new Date().toISOString(),
    source:    "Google Apps Script",
    data:      data || {}
  };

  var response = UrlFetchApp.fetch(CONFIG.generic_webhook_url, {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return response.getResponseCode() >= 200 && response.getResponseCode() < 300
    ? "OK (" + response.getResponseCode() + ")"
    : "Error " + response.getResponseCode();
}

// ----------------------------------------------------------------
// onSheetEdit: triggered when a row is added to the sheet
// ----------------------------------------------------------------
function onSheetEdit(e) {
  var sheet = e.source.getSheetByName(CONFIG.sheet_name);
  if (!sheet || e.range.getSheet().getName() !== CONFIG.sheet_name) return;

  var row = e.range.getRow();
  if (row < 2) return; // skip header

  var values = sheet.getRange(row, 1, 1, 4).getValues()[0];
  var type    = values[0];
  var title   = values[1];
  var detail  = values[2];
  var amount  = values[3];

  if (!title) return;

  var message = "[" + type + "] " + title;
  var data    = { Detail: detail, Amount: amount };

  notifyAll(message, data);
  sheet.getRange(row, 5).setValue("Sent").setBackground("#c6efce");
}

// ----------------------------------------------------------------
// setupTrigger: register the onEdit trigger (run once)
// ----------------------------------------------------------------
function setupTrigger() {
  var ss = SpreadsheetApp.openById(CONFIG.spreadsheet_id);

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "onSheetEdit") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("onSheetEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("Trigger registered for: " + ss.getName());
}

// ----------------------------------------------------------------
// setupSampleSheet: create the notification spreadsheet
// ----------------------------------------------------------------
function setupSampleSheet() {
  var ss    = SpreadsheetApp.create("Notification Triggers");
  var sheet = ss.getActiveSheet().setName(CONFIG.sheet_name);

  var headers = ["Type", "Title", "Detail", "Amount", "Status"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");

  sheet.getRange(2, 1, 1, 4).setValues([["Alert", "New Order", "Order #1234 received", "$500"]]);
  sheet.autoResizeColumns(1, 5);

  // Update CONFIG with new ID
  Logger.log("SPREADSHEET_ID = \"" + ss.getId() + "\"");
  Logger.log("Sheet URL: " + ss.getUrl());
}

// ----------------------------------------------------------------
// Test functions
// ----------------------------------------------------------------
function testSlack() {
  var result = sendSlack(
    ":bell: Test notification from GAS",
    { Source: "Google Apps Script", Time: new Date().toLocaleString() }
  );
  Logger.log("Slack result: " + result);
}

function testLine() {
  var result = sendLine("\n[GAS Test]\nThis is a test notification from Google Apps Script.");
  Logger.log("LINE result: " + result);
}

function testWebhook() {
  var result = sendGenericWebhook(
    "Test notification from GAS",
    { source: "Google Apps Script", time: new Date().toISOString() }
  );
  Logger.log("Webhook result: " + result);
}

function testAll() {
  notifyAll(
    ":rocket: Test from Google Apps Script",
    { Event: "Manual Test", Time: new Date().toLocaleString() }
  );
}
