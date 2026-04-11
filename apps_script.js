function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetId = data.sheet_id;
    const value = data.value;
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    // Find last row with data in column A (date range)
    const lastRow = sheet.getLastRow();
    // Write to column B of the last data row
    sheet.getRange(lastRow, 2).setValue(value);
    return ContentService
      .createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: "ThreeUp CAO Apps Script running"}))
    .setMimeType(ContentService.MimeType.JSON);
}
