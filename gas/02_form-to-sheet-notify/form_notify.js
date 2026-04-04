// ============================================================
// 02: Form Submission → Sheet Logging → Email Notification
// ============================================================
// Form fields (in order):
//   [0] Timestamp
//   [1] Full Name
//   [2] Email Address
//   [3] Subject of Inquiry  (dropdown)
//   [4] Your Inquiry Details
//   [5] How urgent is this inquiry?  (radio)
//
// Setup:
//   1. Link the Google Form to a spreadsheet (Form > Responses > Spreadsheet icon)
//   2. Set SPREADSHEET_ID below (from the spreadsheet URL)
//   3. Run setupTrigger() once from the GAS editor
// ============================================================

var SPREADSHEET_ID = "1y-1hgq-6-_XpdkZL-cH9PJMoO1Z9PNQ4ePt6Xb5uz6Q";
var ADMIN_EMAIL    = Session.getEffectiveUser().getEmail();
var SHEET_NAME     = "Form Responses 1";

// ----------------------------------------------------------------
// onFormSubmit: triggered automatically on each form submission
// ----------------------------------------------------------------
function onFormSubmit(e) {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet  = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var values = e.values;

  var timestamp = values[0] || "";
  var name      = values[1] || "";
  var email     = values[2] || "";
  var subject   = values[3] || "";
  var details   = values[4] || "";
  var urgency   = values[5] || "";

  // --- 1. Mark status in the sheet (column G) ---
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 7).setValue("Notified");

  // --- 2. Notify admin ---
  sendAdminNotification(ss, timestamp, name, email, subject, details, urgency);

  // --- 3. Auto-reply to submitter ---
  if (email) {
    sendAutoReply(name, email, subject, details);
  }
}

// ----------------------------------------------------------------
// sendAdminNotification: email to the admin/operator
// ----------------------------------------------------------------
function sendAdminNotification(ss, timestamp, name, email, subject, details, urgency) {
  var emailSubject = "[New Inquiry] " + subject + " — " + name;
  var body = [
    "A new inquiry has been submitted.",
    "",
    "Timestamp : " + timestamp,
    "Name      : " + name,
    "Email     : " + email,
    "Subject   : " + subject,
    "Urgency   : " + urgency,
    "",
    "Inquiry Details:",
    details,
    "",
    "---",
    "Open the sheet: " + ss.getUrl()
  ].join("\n");

  MailApp.sendEmail({
    to:      ADMIN_EMAIL,
    subject: emailSubject,
    body:    body
  });
}

// ----------------------------------------------------------------
// sendAutoReply: confirmation email to the submitter
// ----------------------------------------------------------------
function sendAutoReply(name, email, subject, details) {
  var replySubject = "Thank you for your inquiry, " + name;
  var responseTime = "1-2 business days";

  var body = [
    "Hi " + name + ",",
    "",
    "Thank you for contacting us. We have received your inquiry",
    "and will get back to you within " + responseTime + ".",
    "",
    "--- Your submission ---",
    "Subject : " + subject,
    "Details : " + details,
    "----------------------",
    "",
    "Best regards,",
    "Support Team"
  ].join("\n");

  MailApp.sendEmail({
    to:      email,
    subject: replySubject,
    body:    body
  });
}

// ----------------------------------------------------------------
// setupTrigger: run once manually to register the form-submit trigger
// ----------------------------------------------------------------
function setupTrigger() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Remove existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("Trigger registered: " + ss.getName());
}

// ----------------------------------------------------------------
// testNotification: manual test without a real form submission
// ----------------------------------------------------------------
function testNotification() {
  var fakeEvent = {
    values: [
      new Date().toLocaleString("en-US"),
      "Test User",
      Session.getEffectiveUser().getEmail(),
      "Technical Support",
      "This is a test inquiry from testNotification().",
      "Medium (Requires attention within 24 hours)"
    ]
  };
  onFormSubmit(fakeEvent);
  Logger.log("Test notification sent.");
}
