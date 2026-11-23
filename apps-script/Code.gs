/**
 * Digital Den Order Logger: Google Sheets connector.
 *
 * Drive layout it expects:
 *   Digital Den/                 (ROOT_FOLDER)
 *     Payment Details            (one file for all years; name contains "Payment")
 *     2026/  Rahul Yadav 2026, Sourav Mondal 2026   (names contain the dealer keyword)
 *     2027/  ...                 (created by the app's "Start new year" button)
 *
 * Set ACCESS_KEY, then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every request must carry the key; without it nothing is read or written.
 *
 * Speed: where the files are is remembered for 6 hours (script cache), so most requests skip the
 * Drive search. If a remembered file can't be opened, the folders are searched again automatically.
 * Writes check and format cells in one read per sheet instead of one read per cell.
 */
const ACCESS_KEY = 'PUT-A-LONG-RANDOM-SECRET-HERE';
const ROOT_FOLDER = 'Digital Den';
const DEALER_KEYWORDS = { rahul: ['rahul'], sourav: ['sourav', 'sourab'] };
const CLIENTS_KEYWORD = 'payment';
const HISTORY_TAB = 'Settled history';
const HISTORY_HEADER = ['Settled on', 'Client', 'Order Date', 'Transaction Amount', 'Given', 'Given Date / Method'];
const MIN_COLS = { rahul: 30, sourav: 30, clients: 8, history: 6 };
const CACHE_KEY = 'dd_files_v1';
const CACHE_SECONDS = 21600;   // 6 hours, the longest the script cache keeps anything

function doPost(e) { return handle_(JSON.parse((e.postData && e.postData.contents) || '{}')); }
function doGet(e) { return handle_((e && e.parameter) || {}); }

function handle_(req) {
  try {
    if (!ACCESS_KEY || ACCESS_KEY.indexOf('PUT-') === 0 || req.key !== ACCESS_KEY) return out_({ ok: false, error: 'Wrong or missing access key.' });
    if (req.action === 'read') {
      const parts = Array.isArray(req.parts) ? req.parts : null;   // optional: only these sheets (the dashboard asks for the dealers only)
      return withContext_(req.year, function (ctx) {
        return out_({ ok: true, data: readAll_(ctx, parts), years: ctx.years, year: ctx.year, missing: ctx.missing });
      });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (req.action === 'write') {
        // reply: 'changed' sends back only the sheets this save touched (the app keeps the rest).
        const only = req.reply === 'changed' ? touched_(req.changes || []) : null;
        return withContext_(req.year, function (ctx) {
          const err = write_(ctx, req.changes || [], only);
          if (err) return out_({ ok: false, conflict: true, error: err });
          SpreadsheetApp.flush();
          return out_({ ok: true, data: readAll_(ctx, only), years: ctx.years, year: ctx.year, missing: ctx.missing });
        });
      } else if (req.action === 'newYear') {
        const made = newYear_(scanContext_(req.year), String(req.to || ''));
        forget_();
        const c2 = context_(made, true);
        SpreadsheetApp.flush();
        return out_({ ok: true, data: readAll_(c2), years: c2.years, year: c2.year, missing: c2.missing });
      }
      return out_({ ok: false, error: 'Unknown action.' });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* Runs fn with the remembered file locations. If a remembered file can't be opened (moved or
   deleted), forget them, search Drive again and run fn once more. Files are opened before
   anything is written, so the second run never repeats a write. */
function withContext_(year, fn) {
  const ctx = context_(year, false);
  try { return fn(ctx); }
  catch (e) {
    if (!(e && e.stale && ctx.fromCache)) throw e;
    forget_();
    return fn(context_(year, true));
  }
}

/* ---------- find the files ---------- */
function root_() {
  const it = DriveApp.getFoldersByName(ROOT_FOLDER);
  if (!it.hasNext()) throw new Error('Folder "' + ROOT_FOLDER + '" was not found in your Google Drive.');
  return it.next();
}
function sheetsIn_(folder) {
  const out = [], it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) out.push(f); }
  return out;
}
function yearFolders_(root) {
  const out = {}, it = root.getFolders();
  while (it.hasNext()) { const f = it.next(); if (/^\d{4}$/.test(f.getName()) && !f.isTrashed()) out[f.getName()] = f; }
  return out;
}
function dealerIds_(files) {
  const ids = {};
  Object.keys(DEALER_KEYWORDS).forEach(function (k) {
    const f = files.filter(function (x) { const n = x.getName().toLowerCase(); return DEALER_KEYWORDS[k].some(function (w) { return n.indexOf(w) >= 0; }); })[0];
    if (f) ids[k] = f.getId();
  });
  return ids;
}

/* Remembered map: { folders: {year: folderId}, clientsId, dealers: {year: {rahul: id, sourav: id}} } */
function cached_() {
  try { const s = CacheService.getScriptCache().get(CACHE_KEY); return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
function remember_(m) { try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(m), CACHE_SECONDS); } catch (e) {} }
function forget_() { try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) {} }

function scanMap_() {
  const root = root_(), folders = yearFolders_(root), m = { folders: {}, clientsId: '', dealers: {} };
  Object.keys(folders).forEach(function (y) { m.folders[y] = folders[y].getId(); });
  const clients = sheetsIn_(root).filter(function (f) { return f.getName().toLowerCase().indexOf(CLIENTS_KEYWORD) >= 0; })[0];
  if (clients) m.clientsId = clients.getId();
  return m;
}

function context_(year, fresh) {
  let m = fresh ? null : cached_(), fromCache = !!m;
  if (m && year && !m.folders[year]) { m = null; fromCache = false; }   // a year we don't know yet: look again
  if (!m) m = scanMap_();
  const years = Object.keys(m.folders).sort();
  if (!years.length) throw new Error('No year folder (like "2026") inside "' + ROOT_FOLDER + '".');
  const y = year && m.folders[year] ? String(year) : years[years.length - 1];
  if (!m.dealers[y]) m.dealers[y] = dealerIds_(sheetsIn_(DriveApp.getFolderById(m.folders[y])));
  remember_(m);
  const sheets = {}, missing = [];
  if (m.clientsId) { sheets.clients = { id: m.clientsId, tab: '' }; sheets.history = { id: m.clientsId, tab: HISTORY_TAB, create: true, header: HISTORY_HEADER }; }
  else missing.push('Payment Details');
  Object.keys(DEALER_KEYWORDS).forEach(function (k) {
    if (m.dealers[y][k]) sheets[k] = { id: m.dealers[y][k], tab: '' }; else missing.push(k + ' ' + y);
  });
  return { years: years, year: y, sheets: sheets, missing: missing, open: {}, shs: {}, tz: {}, fromCache: fromCache };
}

/* Full Drive search with folder objects; only "Start next year" needs this. */
function scanContext_(year) {
  const root = root_(), folders = yearFolders_(root), years = Object.keys(folders).sort();
  if (!years.length) throw new Error('No year folder (like "2026") inside "' + ROOT_FOLDER + '".');
  const y = year && folders[year] ? String(year) : years[years.length - 1];
  const sheets = {}, missing = [];
  const clients = sheetsIn_(root).filter(function (f) { return f.getName().toLowerCase().indexOf(CLIENTS_KEYWORD) >= 0; })[0];
  if (clients) { sheets.clients = { id: clients.getId(), tab: '' }; sheets.history = { id: clients.getId(), tab: HISTORY_TAB, create: true, header: HISTORY_HEADER }; }
  else missing.push('Payment Details');
  const ids = dealerIds_(sheetsIn_(folders[y]));
  Object.keys(DEALER_KEYWORDS).forEach(function (k) { if (ids[k]) sheets[k] = { id: ids[k], tab: '' }; else missing.push(k + ' ' + y); });
  return { root: root, folders: folders, years: years, year: y, sheets: sheets, missing: missing, open: {}, shs: {}, tz: {} };
}

function sheet_(ctx, key, create) {
  const memo = ctx.shs[key];
  if (memo && (memo.sh || !create)) return memo;
  const c = ctx.sheets[key];
  if (!c) throw new Error('No sheet found for ' + key + ' in ' + ctx.year + '.');
  let ss = ctx.open[c.id];
  if (!ss) {
    try { ss = SpreadsheetApp.openById(c.id); }
    catch (e) { const err = new Error('Could not open the ' + key + ' sheet: ' + ((e && e.message) || e)); err.stale = true; throw err; }
    ctx.open[c.id] = ss;
  }
  let sh = c.tab ? ss.getSheetByName(c.tab) : ss.getSheets()[0];
  if (!sh && c.create && create) {
    sh = ss.insertSheet(c.tab, ss.getSheets().length);
    if (c.header) sh.getRange(1, 1, 1, c.header.length).setValues([c.header]).setFontWeight('bold');
  }
  return (ctx.shs[key] = { ss: ss, sh: sh });
}
function tz_(ctx, ss) { const id = ss.getId(); return ctx.tz[id] || (ctx.tz[id] = ss.getSpreadsheetTimeZone()); }

/* ---------- read ---------- */
function cellOut_(v, tz) {
  if (v instanceof Date) return { d: Utilities.formatDate(v, tz, 'yyyy-MM-dd') };
  return v;
}
function readAll_(ctx, parts) {
  const data = {};
  Object.keys(ctx.sheets).forEach(function (k) {
    if (parts && parts.indexOf(k) < 0) return;
    const o = sheet_(ctx, k, false);
    if (!o.sh) { data[k] = { values: [], name: '' }; return; }
    const tz = tz_(ctx, o.ss);
    const raw = o.sh.getDataRange().getValues();             // one call instead of last row + last column + range
    const cols = Math.max(raw.length ? raw[0].length : 0, MIN_COLS[k] || 1);
    while (raw.length < 2) raw.push([]);
    data[k] = {
      name: o.ss.getName(),
      values: raw.map(function (r) {
        const out = new Array(cols);
        for (var i = 0; i < cols; i++) out[i] = i < r.length ? cellOut_(r[i], tz) : '';
        return out;
      })
    };
  });
  return data;
}

/* ---------- write ---------- */
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
/* One read of the smallest block that covers the given cells; returns get(row, col). */
function block_(sh, cells, how) {
  var r0 = Infinity, r1 = 0, c0 = Infinity, c1 = 0;
  cells.forEach(function (c) { r0 = Math.min(r0, c.row); r1 = Math.max(r1, c.row); c0 = Math.min(c0, c.col); c1 = Math.max(c1, c.col); });
  const rng = sh.getRange(r0, c0, r1 - r0 + 1, c1 - c0 + 1);
  const vals = how === 'formats' ? rng.getNumberFormats() : rng.getValues();
  return function (row, col) { return vals[row - r0][col - c0]; };
}
function touched_(changes) {
  const keys = {};
  changes.forEach(function (c) { keys[c.sheet] = true; });
  if (keys.clients || keys.history) { keys.clients = true; keys.history = true; }   // same file
  return Object.keys(keys);
}
function bySheet_(list) {
  const g = {};
  list.forEach(function (c) { (g[c.sheet] = g[c.sheet] || []).push(c); });
  return g;
}
/* changes: cell writes {sheet,row,col,v|f|date,was,styleRow}
            {sheet,op:'insertCols',cols:[1-based columns, inserted before each]}
            {sheet,op:'copyFormat',from:[row,col,rows,cols],to:[row,col]} */
function write_(ctx, changes, only) {
  // Open every sheet the reply will read first, so a file that can't be opened stops us before
  // anything is written.
  Object.keys(ctx.sheets).forEach(function (k) { if (!only || only.indexOf(k) >= 0) sheet_(ctx, k, false); });
  changes.forEach(function (c) { sheet_(ctx, c.sheet, false); });

  // Someone else changed a cell since this device read it? One read per sheet, then compare in order.
  const checked = changes.filter(function (c) { return !(c.op || c.f || c.was === undefined); });
  const getNow = {}, groups0 = bySheet_(checked);
  Object.keys(groups0).forEach(function (k) {
    const o = sheet_(ctx, k, false);
    if (o.sh) getNow[k] = { get: block_(o.sh, groups0[k], 'values'), tz: tz_(ctx, o.ss) };
  });
  for (var i = 0; i < checked.length; i++) {
    const c = checked[i], b = getNow[c.sheet];
    if (!b) continue;
    if (!same_(cellOut_(b.get(c.row, c.col), b.tz), c.was)) return 'The sheet was changed from another device while you were filling this in (' + c.sheet + ' row ' + c.row + '). Nothing was saved. Please try again.';
  }

  changes.filter(function (c) { return c.op === 'insertCols'; }).forEach(function (c) {
    const sh = sheet_(ctx, c.sheet, true).sh;
    c.cols.slice().sort(function (a, b) { return b - a; }).forEach(function (col) { sh.insertColumnBefore(col); });
  });
  changes.filter(function (c) { return c.op === 'copyFormat'; }).forEach(function (c) {
    const sh = sheet_(ctx, c.sheet, true).sh;
    sh.getRange(c.from[0], c.from[1], c.from[2], c.from[3]).copyTo(sh.getRange(c.to[0], c.to[1], c.from[2], c.from[3]), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  const groups = {};
  changes.forEach(function (c) {
    if (c.op || !c.styleRow) return;
    const k = c.sheet + ':' + c.row + ':' + c.styleRow;
    const g = groups[k] = groups[k] || { sheet: c.sheet, row: c.row, styleRow: c.styleRow, min: c.col, max: c.col };
    g.min = Math.min(g.min, c.col); g.max = Math.max(g.max, c.col);
  });
  Object.keys(groups).forEach(function (k) {
    const g = groups[k], w = g.max - g.min + 1, sh = sheet_(ctx, g.sheet, true).sh;
    sh.getRange(g.styleRow, g.min, 1, w).copyTo(sh.getRange(g.row, g.min, 1, w), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });

  // Date cells keep their own date format; read the formats once per sheet (after the style copy above).
  const fmt = {}, dates = bySheet_(changes.filter(function (c) { return !c.op && !c.f && c.date; }));
  Object.keys(dates).forEach(function (k) { fmt[k] = block_(sheet_(ctx, k, true).sh, dates[k], 'formats'); });

  changes.forEach(function (c) {
    if (c.op) return;
    const r = sheet_(ctx, c.sheet, true).sh.getRange(c.row, c.col);
    if (c.f) { r.setFormula(c.f); return; }
    if (c.date) {
      if (!/[dy]/i.test(fmt[c.sheet](c.row, c.col))) r.setNumberFormat('dd/mm/yyyy');
      r.setValue(serial_(c.date));
      return;
    }
    r.setValue(c.v === undefined ? '' : c.v);
  });
  return null;
}

/* ---------- start a new year ----------
   Copies every dealer sheet of the latest year into a new year folder, keeps the layout and the
   row-2 totals, clears the orders and payments, and puts last year's LEFT (AA2) into the GOT column
   of row 2 as a negative "Previous Year Due", the same way the yearly sheets were started by hand. */
function newYear_(ctx, to) {
  if (!/^\d{4}$/.test(to)) throw new Error('Give the new year as 4 digits.');
  if (ctx.folders[to]) throw new Error('The ' + to + ' folder already exists.');
  const from = ctx.years[ctx.years.length - 1];
  if (+to <= +from) throw new Error('The new year must be after ' + from + '.');
  const src = scanContext_(from);
  const folder = ctx.root.createFolder(to);
  Object.keys(DEALER_KEYWORDS).forEach(function (k) {
    if (!src.sheets[k]) return;
    const file = DriveApp.getFileById(src.sheets[k].id);
    let name = file.getName().indexOf(from) >= 0 ? file.getName().split(from).join(to) : file.getName() + ' ' + to;
    const copy = file.makeCopy(name, folder);
    const oldSh = SpreadsheetApp.openById(src.sheets[k].id).getSheets()[0];
    const left = Number(oldSh.getRange('AA2').getValue()) || 0;
    const sh = SpreadsheetApp.openById(copy.getId()).getSheets()[0];
    const last = sh.getLastRow(), cols = Math.max(sh.getLastColumn(), 30);
    if (last >= 3) sh.getRange(3, 1, last - 2, cols).clearContent();
    const row2 = sh.getRange(2, 1, 1, cols).getValues()[0];
    if (!row2.some(function (v) { return String(v).toLowerCase().indexOf('previous year due') >= 0; })) sh.getRange('W2').setValue('Previous Year Due');
    sh.getRange('Y2').setValue(-left);
  });
  return to;
}

/* Run this once from the editor to grant Drive access. */
function authorize() { Logger.log(root_().getName() + " found"); }
