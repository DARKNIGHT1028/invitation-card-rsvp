// Google Apps Script backend for the SeJasmin invitation.
// Sheet layout (tab "RSVP"): A Timestamp | B Nama Tetamu | C Kehadiran | D ID (helper)
const SHEET_ID  = '1kyoV80-YOn6bjPnX-1M4u6GWoLWNTY6E';
const ADMIN_KEY = 'CHANGE_THIS_PASSWORD';   // password for index.html#admin
const TAB = 'RSVP';
const HEAD = ['Timestamp', 'Nama Tetamu', 'Kehadiran', 'ID'];

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  if (sh.getRange(1, 1, 1, 4).getValues()[0].join('|') !== HEAD.join('|')) {
    sh.getRange(1, 1, 1, 4).setValues([HEAD]);
  }
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
  const rows = [];
  v.forEach((r, i) => {
    if (!r[1] && !r[2]) return;                       // skip empty rows
    rows.push({
      id: String(r[3] || 'row' + (i + 2)),
      at: r[0] ? new Date(r[0]).getTime() : 0,
      name: String(r[1]),
      status: String(r[2])
    });
  });
  rows.sort((a, b) => b.at - a.at);
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    const ids = n > 0 ? sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0])) : [];
    const idx = ids.indexOf(String(b.id));

    if (b.action === 'rsvp') {
      const name = String(b.name || '').trim().slice(0, 80);
      if (!b.id || !name || ['Hadir', 'Tidak hadir'].indexOf(b.status) < 0) return json_({ ok: false, error: 'invalid' });
      const row = [new Date(), safe_(name), b.status, String(b.id)];
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 4).setValues([row]); else sh.appendRow(row);
      return json_({ ok: true });
    }
    if (b.action === 'delete') {
      if (b.key !== ADMIN_KEY) return json_({ ok: false, error: 'auth' });
      const m = /^row(\d+)$/.exec(String(b.id));       // rows added by hand have no ID
      if (m) { const r = Number(m[1]); if (r >= 2 && r <= sh.getLastRow()) sh.deleteRow(r); }
      else if (idx >= 0) sh.deleteRow(idx + 2);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown' });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}
