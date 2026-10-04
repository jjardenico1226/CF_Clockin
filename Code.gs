/**
 * Crossroads Fellowship Inc. – Attendance backend (Google Apps Script)
 *
 * SETUP
 * 1. Create a new Google Sheet (e.g. "Crossroads Attendance").
 * 2. Extensions > Apps Script. Delete the sample code and paste this whole file. Save.
 * 3. Deploy > New deployment > type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Authorize when asked, then copy the Web App URL.
 * 5. Paste that URL into CONFIG.SCRIPT_URL at the top of index.html.
 */

const SHEET_NAME = 'Attendance';
const KEYS = ['ts','date','type','churchId','name','age','category','ageGroup','broughtBy','birthdate','address','howKnow','contactPerson','prayer'];
const HEADERS = ['Timestamp','Date','Type','Church ID','Name','Age','Category','Age Group','Brought By (Parent ID)','Birthdate','Address','How They Know the Church','Contact in Church','Prayer Request'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#cfe6fa');
    // keep text columns as plain text so "0001" and dates are not auto-converted
    [1, 2, 4, 10].forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Returns every attendance record
function doGet(e) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ records: [] });
  const values = sh.getRange(2, 1, last - 1, KEYS.length).getDisplayValues();
  const records = values.map(row => {
    const o = {};
    KEYS.forEach((k, i) => o[k] = row[i]);
    return o;
  });
  return json_({ records: records });
}

// Adds new records (members already clocked in for the same date are skipped)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'add') return json_({ error: 'unknown action' });

    const sh = getSheet_();
    const last = sh.getLastRow();
    const seen = {};
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, KEYS.length).getDisplayValues().forEach(r => {
        if (r[2] === 'Member') seen[r[1] + '|' + r[3]] = true;   // date | churchId
      });
    }

    const rows = [], skipped = [];
    (body.records || []).forEach(rec => {
      const key = rec.date + '|' + rec.churchId;
      if (rec.type === 'Member' && seen[key]) { skipped.push(rec); return; }
      if (rec.type === 'Member') seen[key] = true;
      rows.push(KEYS.map(k => rec[k] === undefined || rec[k] === null ? '' : String(rec[k])));
    });

    if (rows.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, KEYS.length).setNumberFormat('@').setValues(rows);
    }
    return json_({ added: rows.length, skipped: skipped });
  } finally {
    lock.releaseLock();
  }
}
