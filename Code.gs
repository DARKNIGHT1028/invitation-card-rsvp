// Google Apps Script backend for the SeJasmin invitation.
const SHEET_ID  = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const ADMIN_KEY = 'CHANGE_THIS_PASSWORD';   // password for index.html#admin
const TAB = 'RSVP';
const HEAD = ['ID', 'Timestamp', 'Guest Name', 'Attendance'];

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function safe_(t) { return /^[=+\-@]/.test(t) ? "'" + t : t; }

function doGet(e) {
  const p = e.parameter || {};
  if (p.action !== 'list') return json_({ ok: true });
  if (p.key !== ADMIN_KEY) return json_({ ok: false, error: 'auth' });
  const v = sheet_().getDataRange().getValues().slice(1);
  const rows = v.map(r => ({ id: String(r[0]), at: new Date(r[1]).getTime(), name: String(r[2]), status: String(r[3]) }))
                .sort((a, b) => b.at - a.at);
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    const idx = ids.indexOf(String(b.id));

    if (b.action === 'rsvp') {
      const name = String(b.name || '').trim().slice(0, 80);
      if (!b.id || !name || ['Hadir', 'Tidak hadir'].indexOf(b.status) < 0) return json_({ ok: false, error: 'invalid' });
      const row = [String(b.id), new Date(), safe_(name), b.status];
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 4).setValues([row]); else sh.appendRow(row);
      return json_({ ok: true });
    }
    if (b.action === 'delete') {
      if (b.key !== ADMIN_KEY) return json_({ ok: false, error: 'auth' });
      if (idx >= 0) sh.deleteRow(idx + 2);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown' });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}
