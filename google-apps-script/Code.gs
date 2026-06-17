const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";
const DEFAULT_SHEET_NAME = "Expense Log";

function doPost(e) {
  const payload = parsePayload_(e);
  const expense = payload.expense || {};
  const sheetName = payload.sheetName || DEFAULT_SHEET_NAME;
  const sheet = getOrCreateSheet_(sheetName);

  ensureHeaders_(sheet);
  sheet.appendRow([
    expense.createdAt || new Date().toISOString(),
    expense.date || "",
    expense.amount || 0,
    expense.kind || "",
    expense.card || "",
    expense.category || "",
    expense.merchant || "",
    expense.note || "",
    expense.id || "",
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }

  return JSON.parse((e && e.postData && e.postData.contents) || "{}");
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, app: "Budget Tap", spreadsheetIdSet: !SPREADSHEET_ID.includes("PASTE_") }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(sheetName) {
  if (SPREADSHEET_ID.includes("PASTE_")) {
    throw new Error("Set SPREADSHEET_ID at the top of Code.gs.");
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function ensureHeaders_(sheet) {
  const headers = ["Created At", "Date", "Amount", "Kind", "Card", "Category", "Merchant", "Note", "Expense ID"];
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = current.every((cell) => cell === "");

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}
