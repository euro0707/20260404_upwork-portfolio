/**
 * Daily / Weekly Report Automation
 * Google Apps Script — Portfolio Sample
 *
 * Features:
 * - Auto-detect column headers (Japanese & English)
 * - Daily and weekly report modes
 * - Email-client-compatible HTML (table-based layout)
 * - Error logging to a dedicated sheet
 * - Works as standalone or container-bound script
 */

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  // Target spreadsheet (leave empty if running as container-bound)
  SPREADSHEET_ID: "",

  EMAIL_TO: "your-email@example.com",
  EMAIL_SUBJECT_DAILY:  "[Daily Report]",
  EMAIL_SUBJECT_WEEKLY: "[Weekly Report]",

  TRIGGER_HOUR_DAILY:  7,  // 07:00 every day
  TRIGGER_HOUR_WEEKLY: 8,  // 08:00 every Monday

  SHEET_NAME:     "Data",
  LOG_SHEET_NAME: "_Log",

  // Keywords used to auto-detect columns from header row
  COLUMN_KEYWORDS: {
    DATE:     ["date", "日付", "日時"],
    CATEGORY: ["category", "カテゴリ", "分類"],
    ITEM:     ["item", "項目", "内容", "品名"],
    AMOUNT:   ["amount", "金額", "売上", "price"],
    STATUS:   ["status", "ステータス", "状態"],
  },

  // Fallback column indices (1-based) when header not found
  COLUMN_FALLBACK: {
    DATE: 1, CATEGORY: 2, ITEM: 3, AMOUNT: 4, STATUS: 5,
  },
};

// ============================================================
// ENTRY POINTS
// ============================================================
function sendDailyReport()  { _sendReport("daily");  }
function sendWeeklyReport() { _sendReport("weekly"); }

function _sendReport(mode) {
  try {
    const ss = CONFIG.SPREADSHEET_ID
      ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error("No data rows found");

    const cols      = _detectColumns(data[0]);
    const rows      = data.slice(1);
    const yesterday = _daysAgo(1);
    const summary   = _calculateSummary(rows, cols, yesterday, mode);
    const htmlBody  = _buildEmailHTML(summary, rows, cols, yesterday, mode);

    const subject = mode === "weekly"
      ? `${CONFIG.EMAIL_SUBJECT_WEEKLY} ${_fmt(yesterday)}`
      : `${CONFIG.EMAIL_SUBJECT_DAILY} ${_fmt(yesterday)}`;

    MailApp.sendEmail({ to: CONFIG.EMAIL_TO, subject, htmlBody });
    _log(`Report sent (${mode}) → ${CONFIG.EMAIL_TO}`);

  } catch (err) {
    _log(`ERROR: ${err.message}`, "ERROR");
    MailApp.sendEmail({
      to: CONFIG.EMAIL_TO,
      subject: "[ERROR] Report failed",
      body: `Error: ${err.message}\nTime: ${new Date().toLocaleString("ja-JP")}`,
    });
  }
}

// ============================================================
// COLUMN AUTO-DETECTION
// ============================================================
function _detectColumns(headers) {
  const norm = s => s.toString().trim().toLowerCase();
  const cols = {};
  for (const [key, keywords] of Object.entries(CONFIG.COLUMN_KEYWORDS)) {
    const idx = headers.findIndex(h => keywords.some(kw => norm(h).includes(norm(kw))));
    cols[key] = idx >= 0 ? idx : CONFIG.COLUMN_FALLBACK[key] - 1;
  }
  return cols;
}

// ============================================================
// AGGREGATION
// ============================================================
function _calculateSummary(rows, cols, targetDate, mode) {
  const summary = {
    target:     { amount: 0, count: 0 },
    thisWeek:   { amount: 0, count: 0 },
    thisMonth:  { amount: 0, count: 0 },
    byCategory: {},
    byStatus:   {},
  };

  rows.forEach(row => {
    const rawDate = row[cols.DATE];
    if (!rawDate) return;
    const rowDate = new Date(rawDate);
    if (isNaN(rowDate.getTime())) return; // skip invalid dates

    const amount   = Number(row[cols.AMOUNT])   || 0;
    const category = String(row[cols.CATEGORY]  || "Uncategorized").trim();
    const status   = String(row[cols.STATUS]    || "Unknown").trim();

    const isTarget = mode === "weekly"
      ? _isSameWeek(rowDate, targetDate)
      : _isSameDay(rowDate, targetDate);

    if (isTarget) {
      summary.target.amount += amount;
      summary.target.count++;

      if (!summary.byCategory[category]) {
        summary.byCategory[category] = { amount: 0, count: 0 };
      }
      summary.byCategory[category].amount += amount;
      summary.byCategory[category].count++;

      summary.byStatus[status] = (summary.byStatus[status] || 0) + 1;
    }

    if (_isSameWeek(rowDate, targetDate))  { summary.thisWeek.amount  += amount; summary.thisWeek.count++;  }
    if (_isSameMonth(rowDate, targetDate)) { summary.thisMonth.amount += amount; summary.thisMonth.count++; }
  });

  return summary;
}

// ============================================================
// HTML EMAIL (table-based layout for Outlook compatibility)
// ============================================================
function _buildEmailHTML(summary, rows, cols, targetDate, mode) {
  const title   = mode === "weekly" ? "Weekly Summary" : "Daily Summary";
  const dateStr = _fmt(targetDate);
  const label   = mode === "weekly" ? "Week" : "Yesterday";

  // KPI row
  const kpiHTML = `
    <table width="100%" cellpadding="0" cellspacing="8" style="padding:16px;background:#f8faff;">
      <tr>
        ${_kpiCell(label + " Count",  summary.target.count + " items",                      "#2563eb")}
        ${_kpiCell(label + " Total",  "¥" + summary.target.amount.toLocaleString(),          "#059669")}
        ${_kpiCell("Month Total",     "¥" + summary.thisMonth.amount.toLocaleString(),        "#7c3aed")}
      </tr>
    </table>`;

  // Category breakdown
  const categoryRows = Object.entries(summary.byCategory)
    .sort((a, b) => b[1].amount - a[1].amount)
    .map(([cat, d]) => `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;">${cat}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${d.count} items</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">¥${d.amount.toLocaleString()}</td>
      </tr>`).join("");

  // Status section
  const statusRows = Object.entries(summary.byStatus)
    .map(([s, c]) => `<td style="padding:4px 12px;background:#f0f4ff;border-radius:20px;font-size:13px;white-space:nowrap;">${s}: <strong>${c}</strong></td><td width="8"></td>`)
    .join("");

  // Detail rows filtered by period
  const detailRows = rows
    .filter(row => {
      const d = new Date(row[cols.DATE]);
      if (isNaN(d.getTime())) return false;
      return mode === "weekly" ? _isSameWeek(d, targetDate) : _isSameDay(d, targetDate);
    })
    .map(row => `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px;">${row[cols.CATEGORY] || "-"}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px;">${row[cols.ITEM]     || "-"}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px;text-align:right;">¥${(Number(row[cols.AMOUNT]) || 0).toLocaleString()}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee;font-size:13px;">${row[cols.STATUS]   || "-"}</td>
      </tr>`).join("");

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f5f7fa;font-family:'Helvetica Neue',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f7fa;">
<tr><td align="center" style="padding:24px;">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

  <!-- Header -->
  <tr><td style="background:linear-gradient(135deg,#2563eb,#1d4ed8);padding:28px 32px;">
    <p style="margin:0;color:rgba(255,255,255,0.7);font-size:13px;">${title}</p>
    <h1 style="margin:4px 0 0;color:#fff;font-size:22px;font-weight:700;">${dateStr}</h1>
  </td></tr>

  <!-- KPI -->
  <tr><td>${kpiHTML}</td></tr>

  <tr><td style="padding:0 32px 24px;">

    <!-- Category -->
    ${categoryRows ? `
    <h2 style="font-size:15px;font-weight:600;color:#374151;margin:24px 0 12px;">Category Breakdown</h2>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #eee;border-radius:8px;">
      <thead><tr style="background:#f9fafb;">
        <th style="padding:10px 12px;text-align:left;font-size:13px;color:#6b7280;font-weight:500;">Category</th>
        <th style="padding:10px 12px;text-align:right;font-size:13px;color:#6b7280;font-weight:500;">Count</th>
        <th style="padding:10px 12px;text-align:right;font-size:13px;color:#6b7280;font-weight:500;">Amount</th>
      </tr></thead>
      <tbody>${categoryRows}</tbody>
    </table>` : ""}

    <!-- Status -->
    ${statusRows ? `
    <h2 style="font-size:15px;font-weight:600;color:#374151;margin:24px 0 12px;">Status</h2>
    <table cellpadding="0" cellspacing="0"><tr>${statusRows}</tr></table>` : ""}

    <!-- Detail -->
    <h2 style="font-size:15px;font-weight:600;color:#374151;margin:24px 0 12px;">Detail</h2>
    ${detailRows ? `
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #eee;">
      <thead><tr style="background:#f9fafb;">
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#6b7280;font-weight:500;">Category</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#6b7280;font-weight:500;">Item</th>
        <th style="padding:10px 12px;text-align:right;font-size:12px;color:#6b7280;font-weight:500;">Amount</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#6b7280;font-weight:500;">Status</th>
      </tr></thead>
      <tbody>${detailRows}</tbody>
    </table>` : `<p style="color:#9ca3af;text-align:center;padding:24px;">No data for this period</p>`}

  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:16px 32px;background:#f9fafb;border-top:1px solid #eee;">
    <p style="margin:0;font-size:12px;color:#9ca3af;text-align:center;">
      Auto-generated by Google Apps Script &nbsp;|&nbsp; ${new Date().toLocaleString("ja-JP")}
    </p>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}

function _kpiCell(label, value, color) {
  return `<td width="33%" style="background:#fff;border-radius:10px;padding:16px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.06);">
    <p style="margin:0;font-size:12px;color:#6b7280;">${label}</p>
    <p style="margin:4px 0 0;font-size:22px;font-weight:700;color:${color};">${value}</p>
  </td>`;
}

// ============================================================
// TRIGGER SETUP
// ============================================================
function setupDailyTrigger() {
  _removeTriggers("sendDailyReport");
  ScriptApp.newTrigger("sendDailyReport")
    .timeBased().everyDays(1).atHour(CONFIG.TRIGGER_HOUR_DAILY).create();
  Logger.log(`Daily trigger set: every day at ${CONFIG.TRIGGER_HOUR_DAILY}:00`);
}

function setupWeeklyTrigger() {
  _removeTriggers("sendWeeklyReport");
  ScriptApp.newTrigger("sendWeeklyReport")
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(CONFIG.TRIGGER_HOUR_WEEKLY).create();
  Logger.log(`Weekly trigger set: every Monday at ${CONFIG.TRIGGER_HOUR_WEEKLY}:00`);
}

function _removeTriggers(fnName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fnName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ============================================================
// LOGGING
// ============================================================
function _log(message, level) {
  level = level || "INFO";
  Logger.log(`[${level}] ${message}`);
  try {
    const ss = CONFIG.SPREADSHEET_ID
      ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Level", "Message"]);
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([new Date(), level, message]);
  } catch (_) {
    // logging must not break main flow
  }
}

// ============================================================
// UTILITIES
// ============================================================
function _daysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return d;
}

function _isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear()
      && a.getMonth()    === b.getMonth()
      && a.getDate()     === b.getDate();
}

function _isSameWeek(a, b) {
  // ISO week: starts Monday
  const monday = new Date(b);
  const day = b.getDay();
  monday.setDate(b.getDate() - (day === 0 ? 6 : day - 1));
  monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);
  return a >= monday && a <= sunday;
}

function _isSameMonth(a, b) {
  return a.getFullYear() === b.getFullYear()
      && a.getMonth()    === b.getMonth();
}

function _fmt(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}/${m}/${d}`;
}
