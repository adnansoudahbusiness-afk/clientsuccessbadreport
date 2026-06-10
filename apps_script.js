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

    Logger.log('doPost received — action: ' + action + ', week: "' + week + '", value: ' + value);
    console.log('doPost week received: "' + week + '"');

    if (lastRow < 2) {
      Logger.log('Sheet has no data rows');
      return ContentService
        .createTextOutput(JSON.stringify({success: false, error: "No data rows found — run agent first", week: week, action: action}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let targetRow = -1;

    // 1. If a week label is provided, find the matching row by column A (case-insensitive, trimmed)
    if (week) {
      const weekNorm = week.toLowerCase().replace(/\s+/g, ' ');
      const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      Logger.log('Scanning ' + colA.length + ' rows for week match');
      for (let i = 0; i < colA.length; i++) {
        const cellVal  = String(colA[i][0]).trim();
        const cellNorm = cellVal.toLowerCase().replace(/\s+/g, ' ');
        Logger.log('Row ' + (i + 2) + ' col A: "' + cellVal + '"');
        if (cellNorm === weekNorm) {
          targetRow = i + 2; // 1-indexed, offset by header row
          Logger.log('Match found at row ' + targetRow);
          break;
        }
      }
      if (targetRow === -1) {
        Logger.log('No match found for week: "' + week + '"');
      }
    }

    // 2. No week match — always return an error (no fallback for any action)
    if (targetRow === -1) {
      Logger.log('No matching row for week: "' + week + '" (action: ' + action + ')');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'No row found matching week: "' + week + '"',
          week: week,
          action: action
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (targetRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'No data row found — run agent first',
          week: week,
          action: action
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'confirm') {
      // Write "YES" to column C (CONFIRM)
      sheet.getRange(targetRow, 3).setValue(value);
      Logger.log('Wrote "' + value + '" to row ' + targetRow + ' col C (CONFIRM)');
    } else if (action === 'submit_wp') {
      // Write Website Patients to column AF (32)
      sheet.getRange(targetRow, 32).setValue(value);
      Logger.log('Wrote ' + value + ' to row ' + targetRow + ' col AF (Website Patients)');
    } else {
      // action == 'submit': write Google Reviews count to column B
      sheet.getRange(targetRow, 2).setValue(value);
      Logger.log('Wrote ' + value + ' to row ' + targetRow + ' col B (Google Reviews)');
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        row: targetRow,
        action: action,
        week: week || 'last'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('doPost exception: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action  = e.parameter && e.parameter.action;

  // ── getSchedule: read CAO Schedule tab from AM dates sheet ──
  if (action === 'getSchedule') {
    try {
      const AM_SHEET_ID = '1Lx54-QMM6IONvnVoopNNEjT7oXLmSc_HZTZ4RTG_Qg4';
      const ss          = SpreadsheetApp.openById(AM_SHEET_ID);
      const sheet       = ss.getSheetByName('CAO Schedule');
      if (!sheet) {
        return ContentService
          .createTextOutput(JSON.stringify({rows: [], error: 'CAO Schedule tab not found'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({rows: []}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Header: Client, Strikes, Last Sent, Next Send, Type, Days, Status, Updated
      const raw  = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      const rows = raw
        .filter(function(r) { return r[0] !== ''; })
        .map(function(r) {
          return {
            client_name: r[0],
            strikes:     r[1] === '' ? 0 : Number(r[1]),
            last_sent:   r[2],
            next_send:   r[3],
            type:        r[4],
            days:        r[5] === '' ? null : Number(r[5]),
            status:      r[6],
            updated:     r[7],
          };
        });
      return ContentService
        .createTextOutput(JSON.stringify({rows: rows}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({rows: [], error: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // If sheet_id param provided, return all data rows for the client portal
  const sheetId = e.parameter && e.parameter.sheet_id;
  if (sheetId) {
    try {
      const ss      = SpreadsheetApp.openById(sheetId);
      const sheet   = ss.getSheets()[0];
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({rows: []}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Read up to 30 columns, skip header row (row 1)
      const raw  = sheet.getRange(2, 1, lastRow - 1, 32).getValues();
      const rows = raw.filter(function(r) { return r[0] !== ''; });
      return ContentService
        .createTextOutput(JSON.stringify({rows: rows}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({error: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default status response
  return ContentService
    .createTextOutput(JSON.stringify({status: "ThreeUp CAO Apps Script running"}))
    .setMimeType(ContentService.MimeType.JSON);
}
