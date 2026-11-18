/**
 * Digital Den Order Logger: Google Sheets connector.
 *
 * Paste this into Extensions > Apps Script of your order Google Sheet,
 * set ACCESS_KEY to a long random secret, then Deploy > New deployment >
 * Web app (Execute as: Me, Who has access: Anyone).
 * The page sends the key with every request; without it nothing is read or written.
 */
const ACCESS_KEY = 'PUT-A-LONG-RANDOM-SECRET-HERE';
const SHEET_NAME = '';        // blank = first sheet in the file
const MIN_COLS = 30;          // A..AD

function doPost(e) { return handle_(JSON.parse((e.postData && e.postData.contents) || '{}')); }
function doGet(e) { return handle_((e && e.parameter) || {}); }

function handle_(req) {
  try {
    if (!ACCESS_KEY || ACCESS_KEY.indexOf('PUT-') === 0 || req.key !== ACCESS_KEY) return out_({ ok: false, error: 'Wrong or missing access key.' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
    if (req.action === 'read') return out_(Object.assign({ ok: true }, read_(ss, sh)));
    if (req.action === 'write') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        const err = write_(ss, sh, req.changes || []);
        if (err) return out_({ ok: false, conflict: true, error: err });
        SpreadsheetApp.flush();
        return out_(Object.assign({ ok: true }, read_(ss, sh)));
      } finally { lock.releaseLock(); }
    }
    return out_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cellOut_(v, tz) {
  if (v instanceof Date) return { d: Utilities.formatDate(v, tz, 'yyyy-MM-dd') };
  return v;
}

function read_(ss, sh) {
  const tz = ss.getSpreadsheetTimeZone();
  const rows = Math.max(sh.getLastRow(), 2), cols = Math.max(sh.getLastColumn(), MIN_COLS);
  const values = sh.getRange(1, 1, rows, cols).getValues().map(function (r) { return r.map(function (v) { return cellOut_(v, tz); }); });
  return { values: values, name: ss.getName() + ' / ' + sh.getName() };
}

function same_(a, b) {
  const norm = function (v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v.d) return 'D' + v.d;
    return String(v).trim().toUpperCase();
  };
  return norm(a) === norm(b);
}

function serial_(iso) {
  const p = iso.split('-').map(Number);
  return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
}

function write_(ss, sh, changes) {
  const tz = ss.getSpreadsheetTimeZone();
  // 1. conflict check: every cell we change must still hold what the page saw
  for (var i = 0; i < changes.length; i++) {
    const c = changes[i];
    if (c.f) continue;   // formula cells (commission etc.) are recalculated, skip
    const now = cellOut_(sh.getRange(c.row, c.col).getValue(), tz);
    if (!same_(now, c.was)) return 'The sheet was changed from another device while you were filling this in (row ' + c.row + '). Nothing was saved. Please press the button again.';
  }
  // 2. copy formatting for new rows from the row above
  const groups = {};
  changes.forEach(function (c) {
    if (!c.styleRow) return;
    const k = c.row + ':' + c.styleRow;
    const g = groups[k] = groups[k] || { row: c.row, styleRow: c.styleRow, min: c.col, max: c.col };
    g.min = Math.min(g.min, c.col); g.max = Math.max(g.max, c.col);
  });
  Object.keys(groups).forEach(function (k) {
    const g = groups[k], w = g.max - g.min + 1;
    sh.getRange(g.styleRow, g.min, 1, w).copyTo(sh.getRange(g.row, g.min, 1, w), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  // 3. write values
  changes.forEach(function (c) {
    const r = sh.getRange(c.row, c.col);
    if (c.f) { r.setFormula(c.f); return; }
    if (c.date) {
      const fmt = r.getNumberFormat();
      if (!/[dy]/i.test(fmt)) r.setNumberFormat('dd/mm/yyyy');
      r.setValue(serial_(c.date));
      return;
    }
    r.setValue(c.v === undefined ? '' : c.v);
  });
  return null;
}
