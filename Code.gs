/**
 * SeJasmin RSVP backend — JSONP edition.
 * Spreadsheet tab: RSVP
 * A Timestamp | B Nama Tetamu | C Kehadiran | D ID
 */
const SHEET_ID = '1sNT9ROqSQHwxhfApVyL3MVKKxeTLlqHWqI0VYw17Hd0';
function adminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
}
const TAB = 'RSVP';
const HEAD = ['Timestamp', 'Nama Tetamu', 'Kehadiran', 'ID'];

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB);
  if (!sh) sh = ss.insertSheet(TAB);
  if (sh.getRange(1, 1, 1, 4).getValues()[0].join('|') !== HEAD.join('|')) {
    sh.getRange(1, 1, 1, 4).setValues([HEAD]);
  }
  return sh;
}

function output_(data, callback) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function safe_(value) {
  const s = String(value || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/**
 * GET handles JSONP requests from GitHub Pages to avoid browser CORS blocking.
 * Actions: health, rsvp, list, delete.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';

  try {
    if (p.action === 'rsvp') {
      const name = String(p.name || '').trim().slice(0, 80);
      const status = String(p.status || '');
      const id = String(p.id || '');
      if (!id || !name || ['Hadir', 'Tidak hadir'].indexOf(status) < 0) {
        return output_({ok:false, error:'invalid'}, cb);
      }
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const sh = sheet_();
        const n = Math.max(0, sh.getLastRow() - 1);
        const ids = n ? sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0])) : [];
        const idx = ids.indexOf(id);
        const row = [new Date(), safe_(name), status, id];
        if (idx >= 0) sh.getRange(idx + 2, 1, 1, 4).setValues([row]);
        else sh.appendRow(row);
      } finally {
        lock.releaseLock();
      }
      return output_({ok:true}, cb);
    }

    if (p.action === 'list') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const values = sheet_().getDataRange().getValues().slice(1);
      const rows = [];
      values.forEach((r, i) => {
        if (!r[1] && !r[2]) return;
        rows.push({
          id: String(r[3] || 'row' + (i + 2)),
          at: r[0] ? new Date(r[0]).getTime() : 0,
          name: String(r[1]),
          status: String(r[2])
        });
      });
      rows.sort((a, b) => b.at - a.at);
      return output_({ok:true, rows:rows}, cb);
    }

    if (p.action === 'delete') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const id = String(p.id || '');
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const sh = sheet_();
        const n = Math.max(0, sh.getLastRow() - 1);
        const ids = n ? sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0])) : [];
        const idx = ids.indexOf(id);
        const m = /^row(\d+)$/.exec(id);
        if (m) {
          const rowNumber = Number(m[1]);
          if (rowNumber >= 2 && rowNumber <= sh.getLastRow()) sh.deleteRow(rowNumber);
        } else if (idx >= 0) {
          sh.deleteRow(idx + 2);
        }
      } finally {
        lock.releaseLock();
      }
      return output_({ok:true}, cb);
    }

    return output_({ok:true, message:'SeJasmin RSVP API is running.'}, cb);
  } catch (err) {
    return output_({ok:false, error:'server'}, cb);
  }
}

/**
 * Retained for compatibility with direct POST clients.
 * The current website uses JSONP GET requests in order to avoid CORS.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const n = Math.max(0, sh.getLastRow() - 1);
    const ids = n ? sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0])) : [];
    const idx = ids.indexOf(String(b.id));
    if (b.action === 'rsvp') {
      const name = String(b.name || '').trim().slice(0, 80);
      if (!b.id || !name || ['Hadir', 'Tidak hadir'].indexOf(b.status) < 0) {
        return output_({ok:false, error:'invalid'});
      }
      const row = [new Date(), safe_(name), b.status, String(b.id)];
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 4).setValues([row]);
      else sh.appendRow(row);
      return output_({ok:true});
    }
    if (b.action === 'delete') {
      if (b.key !== adminKey_()) return output_({ok:false, error:'auth'});
      if (idx >= 0) sh.deleteRow(idx + 2);
      return output_({ok:true});
    }
    return output_({ok:false, error:'unknown'});
  } catch (err) {
    return output_({ok:false, error:'server'});
  } finally {
    lock.releaseLock();
  }
}

/**
 * Manual connection test. Run this from the Apps Script editor.
 * The account running this script must have Editor access to the spreadsheet.
 */
function testConnection() {
  const sh = sheet_();
  Logger.log('Spreadsheet connected successfully.');
  Logger.log('Sheet name: ' + sh.getName());
  Logger.log('Last row: ' + sh.getLastRow());
}
