function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const sheetId = data.sheet_id;
    const action  = data.action || 'submit';
    const value   = data.value;
    const week    = data.week ? String(data.week).trim() : null;

    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: "No data rows found — run agent first"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let targetRow = -1;

    // 1. If a week label is provided, find the matching row by column A (Date Range)
    if (week) {
      const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        if (String(colA[i][0]).trim() === week) {
          targetRow = i + 2;  // 1-indexed, offset by header row
          break;
        }
      }
    }

    // 2. Fallback: use the last row that has data in column A
    if (targetRow === -1) {
      const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
      for (let i = colA.length - 1; i >= 0; i--) {
        if (colA[i][0] !== '') {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (targetRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: "No data row found — run agent first"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'confirm') {
      // Write "YES" to column C (CONFIRM)
      sheet.getRange(targetRow, 3).setValue(value);
    } else {
      // action == 'submit': write Google Reviews count to column B
      sheet.getRange(targetRow, 2).setValue(value);
    }

    return ContentService
      .createTextOutput(JSON.stringify({success: true, row: targetRow, action: action, week: week || 'last'}))
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
