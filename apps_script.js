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
  const action = e.parameter && e.parameter.action;

  // ── getSchedule: read CAO Schedule tab from AM dates sheet ──
  if (action === 'getSchedule') {
    try {
      const AM_SHEET_ID = '12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA';
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
      // Header: Client, Doctor, Strikes, Last Sent, Next Send, Type, Days, Due Today, Updated
      const raw  = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      const rows = raw
        .filter(function(r) { return r[0] !== ''; })
        .map(function(r) {
          return {
            client_name: r[0],
            doctor_name: r[1],
            strikes:     r[2] === '' ? 0 : Number(r[2]),
            last_sent:   r[3],
            next_send:   r[4],
            type:        r[5],
            days:        r[6] === '' ? null : Number(r[6]),
            due_today:   r[7] === 'TRUE',
            updated:     r[8],
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

  if (action === 'getHistory') {
    try {
      const AM_SHEET_ID = '12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA';
      const ss    = SpreadsheetApp.openById(AM_SHEET_ID);
      const sheet = ss.getSheetByName('CAO History');
      if (!sheet) return ContentService.createTextOutput(JSON.stringify({rows:[], error:'CAO History tab not found'})).setMimeType(ContentService.MimeType.JSON);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);
      const raw = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      const rows = raw.filter(function(r){return r[2]!=='';}).map(function(r){
        return {scheduled_date:String(r[0]),actual_date:String(r[1]),client:String(r[2]),am_name:String(r[3]),type:String(r[4]),status:String(r[5]),triggered_by:String(r[6]),reason:String(r[7]),updated:String(r[8])};
      });
      return ContentService.createTextOutput(JSON.stringify({rows:rows})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({rows:[],error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getFutureEvents') {
    try {
      const AM_SHEET_ID = '12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA';
      const ss    = SpreadsheetApp.openById(AM_SHEET_ID);
      const sheet = ss.getSheetByName('CAO Future Events');
      if (!sheet) return ContentService.createTextOutput(JSON.stringify({rows:[],error:'CAO Future Events tab not found'})).setMimeType(ContentService.MimeType.JSON);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);
      const raw = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      const rows = raw.filter(function(r){return r[0]!=='';}).map(function(r){
        return {client:String(r[0]),doctor_name:String(r[1]),am_name:String(r[2]),event_date:String(r[3]),type:String(r[4]),updated:String(r[5])};
      });
      return ContentService.createTextOutput(JSON.stringify({rows:rows})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({rows:[],error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getBreachMonitor') {
    try {
      const AM_SHEET_ID = '12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA';
      const ss    = SpreadsheetApp.openById(AM_SHEET_ID);
      const sheet = ss.getSheetByName('CAO Breach Monitor');
      if (!sheet) return ContentService.createTextOutput(JSON.stringify({rows:[],error:'CAO Breach Monitor tab not found'})).setMimeType(ContentService.MimeType.JSON);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);
      const raw = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
      const rows = raw.filter(function(r){return r[0]!=='';}).map(function(r){
        return {
          client:        String(r[0]),
          flag:          String(r[1]),
          streak:        Number(r[2]),
          max_streak:    Number(r[3]),
          alert_level:   String(r[4]),
          w1: String(r[5]), w2: String(r[6]), w3: String(r[7]),
          w4: String(r[8]), w5: String(r[9]), w6: String(r[10]),
          coverage_gap:  String(r[11]),
          updated:       String(r[12]),
          last_entry:    String(r[13]),
          missing_weeks: String(r[14]),
        };
      });
      return ContentService.createTextOutput(JSON.stringify({rows:rows})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({rows:[],error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
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
