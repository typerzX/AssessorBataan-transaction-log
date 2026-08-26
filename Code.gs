/**
 * Apps Script backend for the Transaction Receiving Log.
 * Routes each entry to the sheet tab matching its month
 * (tabs must be named exactly: January, February, March, ... December).
 *
 * Each monthly tab must have this header row in Row 1:
 * Tracking Code | Date | Receiving Office | Subject | Date Transmitted | Remarks
 */

const HEADERS = ["Tracking Code", "Date", "Receiving Office", "Subject", "Date Transmitted", "Remarks"];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetName = data.sheetName; // e.g. "January" — sent by the page
    const sheet = getOrCreateSheet(sheetName);

    sheet.appendRow([
      data.trackingCode || "",
      data.date || "",
      data.receivingOffice || "",
      data.subject || "",
      data.dateTransmitted || "",
      data.remarks || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestedSheet = e.parameter.sheet; // e.g. ?sheet=February

  const sheetName = requestedSheet || currentMonthName();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // drop header row

  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- helpers ----

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function currentMonthName() {
  const months = ["January","February","March","April","May","June",
                   "July","August","September","October","November","December"];
  return months[new Date().getMonth()];
}
