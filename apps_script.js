function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetId = data.sheet_id;
    const action  = data.action || 'submit';
    const value   = data.value;
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheets()[0];

    // Find last row with data in column A (Date Range)
    // Column A always has the week date range written by the agent
    const lastRow = sheet.getLastRow();
    const colA    = sheet.getRange(1, 1, lastRow, 1).getValues();
    let targetRow = -1;
    for (let i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] !== '') {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: "No data row found — run agent first"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'confirm') {
      // Write "YES" to column C (CONFIRM column)
      sheet.getRange(targetRow, 3).setValue(value);
    } else {
      // action == 'submit' or unset: write review count to column B
      sheet.getRange(targetRow, 2).setValue(value);
    }

    return ContentService
      .createTextOutput(JSON.stringify({success: true, row: targetRow, action: action}))
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
