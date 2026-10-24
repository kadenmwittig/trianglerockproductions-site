/**
 * Classical Violence — Mailing List receiver
 * Paste into the sheet's Extensions > Apps Script editor, then Deploy > New deployment > Web app.
 * Every website sign-up becomes a new row: Timestamp | Name | Email | Interest | Source
 */
const SHEET_NAME = ''; // leave blank to use the first tab

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Spam trap: real people never fill the hidden "website" field.
  if (p.website) return json({ ok: true });

  const email = String(p.email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return json({ ok: false, error: 'invalid email' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Interest', 'Source']);
    }
    sheet.appendRow([
      new Date(),
      clean(p.name),
      clean(email),
      clean(p.interest || 'Supporter'),
      clean(p.source || 'Classical Violence page')
    ]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

// Stops values starting with = + - @ from being treated as spreadsheet formulas.
function clean(v) {
  const s = String(v || '').slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
