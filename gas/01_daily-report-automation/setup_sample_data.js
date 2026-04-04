/**
 * Setup: Creates a sample spreadsheet for testing Daily Report Automation.
 * Run this function ONCE from the GAS editor.
 * It will log the Spreadsheet ID — copy it into CONFIG.SPREADSHEET_ID.
 */
function createSampleSpreadsheet() {
  const ss = SpreadsheetApp.create("Daily Report - Sample Data");
  const sheet = ss.getActiveSheet();
  sheet.setName("Data");

  // Header row
  sheet.appendRow(["date", "category", "item", "amount", "status"]);

  // Sample data: yesterday + a few days back
  const today = new Date();
  const rows = [
    // Yesterday
    [_offsetDate(today, 1), "Sales",  "Product A",  15000, "Completed"],
    [_offsetDate(today, 1), "Sales",  "Product B",   8000, "Completed"],
    [_offsetDate(today, 1), "Refund", "Product C",   3000, "Pending"],
    // 2 days ago
    [_offsetDate(today, 2), "Sales",  "Product D",  22000, "Completed"],
    [_offsetDate(today, 2), "Sales",  "Product E",   5500, "Completed"],
    // 3 days ago
    [_offsetDate(today, 3), "Sales",  "Product F",  18000, "Completed"],
    [_offsetDate(today, 3), "Support","Ticket #101",     0, "Resolved"],
  ];

  rows.forEach(row => sheet.appendRow(row));

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#2563eb");
  headerRange.setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  // Auto-resize columns
  sheet.autoResizeColumns(1, 5);

  Logger.log("===========================================");
  Logger.log("Spreadsheet created!");
  Logger.log("URL: " + ss.getUrl());
  Logger.log("Spreadsheet ID: " + ss.getId());
  Logger.log("===========================================");
  Logger.log("Copy the ID above into CONFIG.SPREADSHEET_ID");
}

function _offsetDate(base, n) {
  const d = new Date(base);
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
}
