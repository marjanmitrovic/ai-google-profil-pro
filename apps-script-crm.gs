/**
 * AI Google Profil Pro - CRM Apps Script v3
 *
 * Features:
 * - saves leads to Google Sheets
 * - sends email notification
 * - default status NEW
 * - CRM statuses supported:
 *   NEW, CONTACTED, OFFER SENT, WON, LOST
 */

const SHEET_NAME = "Leads";
const NOTIFICATION_EMAIL = "Marjan.posao@gmail.com";
const STATUSES = ["NEW", "CONTACTED", "OFFER SENT", "WON", "LOST"];

function doGet(e) {
  return json_({
    ok: true,
    message: "AI Google Profil Pro CRM is active",
    statuses: STATUSES
  });
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Spreadsheet not found. Open Apps Script from Google Sheet: Extensions -> Apps Script.");
    }

    const sheet = getOrCreateSheet_(ss);

    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No POST data received.");
    }

    const status = normalizeStatus_(data.status || "NEW");

    sheet.appendRow([
      new Date(),
      data.language || "",
      data.source || "",
      data.name || "",
      data.business || "",
      data.city || "",
      data.email || "",
      data.phone || "",
      data.profile || "",
      data.package || "",
      data.message || "",
      status,
      ""
    ]);

    sendNotification_(data, status);

    return json_({ ok: true, message: "Lead saved", status: status });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  }
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Created At",
      "Language",
      "Source",
      "Name",
      "Business",
      "City",
      "Email",
      "Phone",
      "Google Profile",
      "Package",
      "Message",
      "Status",
      "Notes"
    ]);
    sheet.setFrozenRows(1);
  }

  ensureStatusValidation_(sheet);
  sheet.autoResizeColumns(1, 13);
  return sheet;
}

function ensureStatusValidation_(sheet) {
  const maxRows = Math.max(sheet.getMaxRows(), 1000);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 12, maxRows - 1, 1).setDataValidation(rule);
}

function normalizeStatus_(status) {
  const clean = String(status || "").trim().toUpperCase();
  return STATUSES.includes(clean) ? clean : "NEW";
}

function sendNotification_(data, status) {
  const subject = "New lead - AI Google Profil Pro";
  const body =
    "New lead received:\n\n" +
    "Status: " + status + "\n" +
    "Name: " + (data.name || "") + "\n" +
    "Business: " + (data.business || "") + "\n" +
    "City: " + (data.city || "") + "\n" +
    "Email: " + (data.email || "") + "\n" +
    "Phone: " + (data.phone || "") + "\n" +
    "Google Profile: " + (data.profile || "") + "\n" +
    "Package: " + (data.package || "") + "\n" +
    "Language: " + (data.language || "") + "\n\n" +
    "Message:\n" + (data.message || "");

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
