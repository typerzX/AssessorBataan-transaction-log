/**
 * Apps Script backend for the Transaction Receiving Log.
 * HEADERS order:
 * Control No. | Date | Received | Office | Event/Subject | Date of Event | Officer in Charge | Classification | Remarks
 * Each monthly tab must have these headers in Row 1.
 */

const HEADERS = [
  "Control No.",
  "Date",
  "Received",
  "Office",
  "Event/Subject",
  "Date of Event",
  "Officer in Charge",
  "Classification",
  "Remarks"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetName = data.sheetName; // e.g. "January"
    const sheet = getOrCreateSheet(sheetName);

    // Append in exactly the HEADERS order
    sheet.appendRow([
      data.controlNo || "",
      data.date || "",
      data.received || "",
      data.office || "",
      data.eventSubject || "",
      data.dateOfEvent || "",
      data.officerInCharge || "",
      data.classification || "",
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
  const requestedSheet = e.parameter.sheet;
  const sheetName = requestedSheet || currentMonthName();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  // Drop the header row (first row)
  const rows = values.slice(1);

  // Format each row to ensure dates are clean strings
  const formattedRows = rows.map(row => {
    return row.map(cell => {
      if (cell instanceof Date) {
        // Format as YYYY-MM-DD
        const year = cell.getFullYear();
        const month = String(cell.getMonth() + 1).padStart(2, '0');
        const day = String(cell.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
      return cell;
    });
  });

  return ContentService
    .createTextOutput(JSON.stringify(formattedRows))
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
    // Optionally format the whole sheet as plain text to avoid automatic date conversion
    // but not necessary if we send strings.
  }
  return sheet;
}

function currentMonthName() {
  const months = ["January","February","March","April","May","June",
                  "July","August","September","October","November","December"];
  return months[new Date().getMonth()];
}
