/**
 * SeJasmin RSVP -> Google Sheets
 * Deploy as a Web App: Execute as Me; access: Anyone.
 * Keep the spreadsheet itself private.
 */
const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'RSVP';
const HEADERS = ['Timestamp', 'Guest Name', 'Attendance'];

function doGet() {
  return HtmlService.createHtmlOutput('SeJasmin RSVP endpoint is running.');
}

function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  if (String(params.website || '').trim() !== '') return reply_({ok: true});
  if (params.action !== 'submit') return reply_({ok: false, message: 'Permintaan tidak sah.'});

  const name = String(params.name || '').trim().replace(/\s+/g, ' ');
  const status = String(params.status || '').trim();
  if (!name || name.length > 80) return reply_({ok: false, message: 'Sila masukkan nama (maksimum 80 aksara).'});
  if (status !== 'Hadir' && status !== 'Tidak hadir') return reply_({ok: false, message: 'Pilihan kehadiran tidak sah.'});
  if (SPREADSHEET_ID === 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') return reply_({ok: false, message: 'Penganjur belum menetapkan Google Sheet.'});

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    const safeName = /^[=+\-@]/.test(name) ? "'" + name : name;
    sheet.appendRow([new Date(), safeName, status]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    SpreadsheetApp.flush();
    return reply_({ok: true});
  } catch (error) {
    console.error(error);
    return reply_({ok: false, message: 'Tidak dapat menyimpan RSVP sekarang. Sila cuba lagi.'});
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function reply_(payload) {
  const message = JSON.stringify(Object.assign({source: 'sejasmin-rsvp'}, payload)).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.parent.postMessage(' + message + ', "*");</script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}