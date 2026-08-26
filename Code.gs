/**
 * Apps Script backend for the Transaction Receiving Log.
 * Deploy this bound to your Google Sheet (Extensions > Apps Script).
 *
 * Sheet must have this header row in Row 1 (exact order):
 * Tracking Code | Date | Receiving Office | Subject | Date Transmitted | Remarks
 */

const SHEET_NAME = "Sheet1"; // change if your tab is named differently

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    sheet.appendRow([
      data.trackingCode || "",
      data.date || "",
      data.receivingOffice || "",
      data.subject || "",
      data.dateTransmitted || "",
      data.remarks || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // drop header row

  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}
