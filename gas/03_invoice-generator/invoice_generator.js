// ============================================================
// 03: Invoice Generator
//   Spreadsheet → Google Docs template → PDF → Drive + Email
// ============================================================
// Setup:
//   1. Create a Google Docs invoice template (see README)
//   2. Create a spreadsheet with invoice data (see README)
//   3. Set TEMPLATE_DOC_ID, SPREADSHEET_ID, OUTPUT_FOLDER_ID below
//   4. Run generateAllPendingInvoices() or generateInvoice(row) manually
// ============================================================

var TEMPLATE_DOC_ID  = "1E_I5ibBkD0qNiwq7yKNjOHazUnCdQIz3A74bjGwQT5M";
var SPREADSHEET_ID   = "1MNdad8D2kv5sxjzblqjoA7Lfn7USbgBG5FUqLpuO77Y";
var OUTPUT_FOLDER_ID = "1EzJXBV-8v1aqL7AXK7ydsK2N-n-uruUW";
var SHEET_NAME       = "Invoices";

// Column indices (1-based)
var COL = {
  invoice_number : 1,
  issue_date     : 2,
  due_date       : 3,
  client_name    : 4,
  client_email   : 5,
  item1_desc     : 6,
  item1_qty      : 7,
  item1_price    : 8,
  item2_desc     : 9,
  item2_qty      : 10,
  item2_price    : 11,
  item3_desc     : 12,
  item3_qty      : 13,
  item3_price    : 14,
  notes          : 15,
  status         : 16   // "Pending" → set to "Sent" after processing
};

// ----------------------------------------------------------------
// generateAllPendingInvoices: process all rows with status "Pending"
// ----------------------------------------------------------------
function generateAllPendingInvoices() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
                .getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();

  var count = 0;
  for (var i = 1; i < data.length; i++) { // skip header row
    if (data[i][COL.status - 1] === "Pending") {
      generateInvoice(i + 1); // 1-based row number
      count++;
    }
  }
  Logger.log("Processed " + count + " invoice(s).");
}

// ----------------------------------------------------------------
// generateInvoice: generate PDF for a specific row
// ----------------------------------------------------------------
function generateInvoice(row) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
                .getSheetByName(SHEET_NAME);

  var invoiceNumber = sheet.getRange(row, COL.invoice_number).getValue();
  var issueDate     = formatDate(sheet.getRange(row, COL.issue_date).getValue());
  var dueDate       = formatDate(sheet.getRange(row, COL.due_date).getValue());
  var clientName    = sheet.getRange(row, COL.client_name).getValue();
  var clientEmail   = sheet.getRange(row, COL.client_email).getValue();
  var notes         = sheet.getRange(row, COL.notes).getValue();

  // Build line items
  var items = [];
  var subtotal = 0;
  for (var col = COL.item1_desc; col <= COL.item3_desc; col += 3) {
    var desc  = sheet.getRange(row, col).getValue();
    var qty   = sheet.getRange(row, col + 1).getValue();
    var price = sheet.getRange(row, col + 2).getValue();
    if (desc) {
      items.push({ desc: desc, qty: qty, price: price, total: qty * price });
      subtotal += qty * price;
    }
  }

  var tax   = Math.round(subtotal * 0.1 * 100) / 100;
  var total = Math.round((subtotal + tax) * 100) / 100;

  // --- 1. Copy template ---
  var templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  var folder       = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  var docName      = "Invoice_" + invoiceNumber + "_" + clientName;
  var newFile      = templateFile.makeCopy(docName, folder);
  var doc          = DocumentApp.openById(newFile.getId());
  var body         = doc.getBody();

  // --- 2. Replace placeholders ---
  body.replaceText("{{invoice_number}}", invoiceNumber);
  body.replaceText("{{issue_date}}",     issueDate);
  body.replaceText("{{due_date}}",       dueDate);
  body.replaceText("{{client_name}}",    clientName);
  body.replaceText("{{client_email}}",   clientEmail);
  body.replaceText("{{notes}}",          notes || "");

  // Line items
  for (var i = 0; i < 3; i++) {
    var item = items[i] || { desc: "-", qty: "", price: "", total: "" };
    body.replaceText("{{item" + (i+1) + "_desc}}",  item.desc);
    body.replaceText("{{item" + (i+1) + "_qty}}",   item.qty  !== "" ? String(item.qty)   : "");
    body.replaceText("{{item" + (i+1) + "_price}}", item.price !== "" ? "$" + item.price.toFixed(2) : "");
    body.replaceText("{{item" + (i+1) + "_total}}", item.total !== "" ? "$" + item.total.toFixed(2) : "");
  }

  body.replaceText("{{subtotal}}", "$" + subtotal.toFixed(2));
  body.replaceText("{{tax}}",      "$" + tax.toFixed(2));
  body.replaceText("{{total}}",    "$" + total.toFixed(2));

  doc.saveAndClose();

  // --- 3. Export as PDF ---
  var pdfBlob = DriveApp.getFileById(newFile.getId())
                  .getAs(MimeType.PDF);
  pdfBlob.setName(docName + ".pdf");
  var pdfFile = folder.createFile(pdfBlob);

  // --- 4. Send email to client ---
  if (clientEmail) {
    sendInvoiceEmail(clientName, clientEmail, invoiceNumber, total, dueDate, pdfBlob);
  }

  // --- 5. Mark row as Sent ---
  sheet.getRange(row, COL.status).setValue("Sent");
  sheet.getRange(row, COL.status).setBackground("#c6efce"); // green

  Logger.log("Invoice generated: " + docName + " (PDF: " + pdfFile.getUrl() + ")");
  return pdfFile.getUrl();
}

// ----------------------------------------------------------------
// sendInvoiceEmail: email PDF to client
// ----------------------------------------------------------------
function sendInvoiceEmail(clientName, clientEmail, invoiceNumber, total, dueDate, pdfBlob) {
  var subject = "Invoice #" + invoiceNumber + " from Support Team";
  var body = [
    "Dear " + clientName + ",",
    "",
    "Please find your invoice attached.",
    "",
    "Invoice # : " + invoiceNumber,
    "Amount    : $" + total.toFixed(2),
    "Due Date  : " + dueDate,
    "",
    "Please make payment by the due date.",
    "If you have any questions, reply to this email.",
    "",
    "Thank you for your business!",
    "",
    "Best regards,",
    "Support Team"
  ].join("\n");

  MailApp.sendEmail({
    to:          clientEmail,
    subject:     subject,
    body:        body,
    attachments: [pdfBlob]
  });
}

// ----------------------------------------------------------------
// setupSampleData: populate the spreadsheet with sample invoices
// ----------------------------------------------------------------
function setupSampleData() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Headers
  var headers = [
    "Invoice #", "Issue Date", "Due Date",
    "Client Name", "Client Email",
    "Item 1 Desc", "Item 1 Qty", "Item 1 Price",
    "Item 2 Desc", "Item 2 Qty", "Item 2 Price",
    "Item 3 Desc", "Item 3 Qty", "Item 3 Price",
    "Notes", "Status"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");

  // Sample rows
  var today   = new Date();
  var due     = new Date(today); due.setDate(due.getDate() + 30);

  var rows = [
    [
      "INV-001", today, due,
      "Acme Corp", Session.getEffectiveUser().getEmail(),
      "Web Development", 10, 150,
      "UI Design", 5, 100,
      "", "", "",
      "Net 30", "Pending"
    ],
    [
      "INV-002", today, due,
      "Globex Inc", Session.getEffectiveUser().getEmail(),
      "GAS Automation Setup", 1, 500,
      "Training Session", 2, 200,
      "Bug Fixes", 3, 80,
      "Thank you!", "Pending"
    ]
  ];

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("Sample data created in sheet: " + SHEET_NAME);
}

// ----------------------------------------------------------------
// testSingleInvoice: generate invoice for row 2 only
// ----------------------------------------------------------------
function testSingleInvoice() {
  generateInvoice(2);
  Logger.log("Test invoice generated.");
}

// ----------------------------------------------------------------
// createTemplateAndSheet: one-time setup — creates Docs template,
//   Sheets data file, and Drive output folder, then logs all IDs.
//   Run this FIRST, then copy the IDs into the constants above.
// ----------------------------------------------------------------
function createTemplateAndSheet() {
  // 1. Drive output folder
  var folder = DriveApp.createFolder("Invoice PDFs");

  // 2. Google Docs invoice template
  var doc  = DocumentApp.create("Invoice Template");
  var body = doc.getBody();

  body.appendParagraph("INVOICE").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph("Invoice #: {{invoice_number}}");
  body.appendParagraph("Date:      {{issue_date}}");
  body.appendParagraph("Due Date:  {{due_date}}");
  body.appendParagraph("");
  body.appendParagraph("Bill To:").setBold(true);
  body.appendParagraph("{{client_name}}");
  body.appendParagraph("{{client_email}}");
  body.appendParagraph("");

  // Line items table
  var table = body.appendTable([
    ["Description",       "Qty",            "Unit Price",       "Total"          ],
    ["{{item1_desc}}",    "{{item1_qty}}",  "{{item1_price}}",  "{{item1_total}}"],
    ["{{item2_desc}}",    "{{item2_qty}}",  "{{item2_price}}",  "{{item2_total}}"],
    ["{{item3_desc}}",    "{{item3_qty}}",  "{{item3_price}}",  "{{item3_total}}"]
  ]);
  // Style header row
  var headerRow = table.getRow(0);
  for (var c = 0; c < 4; c++) {
    headerRow.getCell(c).setBackgroundColor("#4a86e8");
    headerRow.getCell(c).editAsText().setBold(true).setForegroundColor("#ffffff");
  }

  body.appendParagraph("");
  body.appendParagraph("Subtotal: {{subtotal}}");
  body.appendParagraph("Tax (10%): {{tax}}");
  body.appendParagraph("Total Due: {{total}}").setBold(true);
  body.appendParagraph("");
  body.appendParagraph("Notes: {{notes}}");

  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  // 3. Spreadsheet for invoice data
  var ss    = SpreadsheetApp.create("Invoice Data");
  var sheet = ss.getActiveSheet().setName("Invoices");

  var headers = [
    "Invoice #", "Issue Date", "Due Date",
    "Client Name", "Client Email",
    "Item 1 Desc", "Item 1 Qty", "Item 1 Price",
    "Item 2 Desc", "Item 2 Qty", "Item 2 Price",
    "Item 3 Desc", "Item 3 Qty", "Item 3 Price",
    "Notes", "Status"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");

  var today = new Date();
  var due   = new Date(today); due.setDate(due.getDate() + 30);
  sheet.getRange(2, 1, 1, 16).setValues([[
    "INV-001", today, due,
    "Acme Corp", Session.getEffectiveUser().getEmail(),
    "Web Development", 10, 150,
    "UI Design", 5, 100,
    "", "", "",
    "Net 30", "Pending"
  ]]);
  sheet.getRange(3, 1, 1, 16).setValues([[
    "INV-002", today, due,
    "Globex Inc", Session.getEffectiveUser().getEmail(),
    "GAS Automation Setup", 1, 500,
    "Training Session", 2, 200,
    "Bug Fixes", 3, 80,
    "Thank you!", "Pending"
  ]]);
  sheet.autoResizeColumns(1, headers.length);
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  // Log all IDs
  Logger.log("=== COPY THESE IDs INTO THE SCRIPT ===");
  Logger.log("TEMPLATE_DOC_ID  = \"" + doc.getId() + "\"");
  Logger.log("SPREADSHEET_ID   = \"" + ss.getId() + "\"");
  Logger.log("OUTPUT_FOLDER_ID = \"" + folder.getId() + "\"");
  Logger.log("======================================");
}

// ----------------------------------------------------------------
// Utility
// ----------------------------------------------------------------
function formatDate(date) {
  if (!date) return "";
  if (typeof date === "string") return date;
  var d = new Date(date);
  return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
}
