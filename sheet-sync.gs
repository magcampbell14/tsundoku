/**
 * Saves Tsundoku data to this Google Sheet.
 *
 * Setup, once, on a computer:
 *   1. In the Sheet: Extensions > Apps Script. Delete what's there and paste this whole file. Save.
 *   2. Deploy > New deployment > gear icon > Web app.
 *      Execute as: Me.  Who has access: Anyone.  Click Deploy, then Authorize and allow access.
 *   3. Copy the Web app URL and paste it into the app: Settings > Google Sheet backup.
 *
 * The URL works like a password: anyone who has it can read and change your saved data,
 * so keep it to yourself.
 */

const SYNC_TAB = '_sync';
const PART = 45000; // stay under Sheets' 50,000-character cell limit

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'load') return json(load_(p.app));
  return json({ ok: true, hello: 'Sheet sync is running' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad request' }); }
  if (!body.app || typeof body.data !== 'string') return json({ ok: false, error: 'missing data' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    save_(body.app, body.data, body.savedAt || Date.now());
    if (body.views) writeViews_(body.views);
    return json({ ok: true, savedAt: body.savedAt });
  } finally {
    lock.releaseLock();
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function syncSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SYNC_TAB);
  if (!sh) {
    sh = ss.insertSheet(SYNC_TAB);
    sh.getRange(1, 1, 1, 4).setValues([['App', 'Saved at', 'Parts', "Saved data (the apps read this to restore. Please don't edit.)"]]);
  }
  return sh;
}

function rowFor_(sh, app) {
  const last = sh.getLastRow();
  if (last >= 2) {
    const apps = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < apps.length; i++) if (apps[i][0] === app) return i + 2;
  }
  return Math.max(last, 1) + 1;
}

function save_(app, data, savedAt) {
  const sh = syncSheet_();
  const row = rowFor_(sh, app);
  const parts = [];
  for (let i = 0; i < data.length; i += PART) parts.push("'" + data.slice(i, i + PART)); // ' keeps it as plain text
  const old = sh.getLastColumn();
  if (old > 3) sh.getRange(row, 4, 1, old - 3).clearContent();
  if (sh.getMaxColumns() < 3 + parts.length) sh.insertColumnsAfter(sh.getMaxColumns(), 3 + parts.length - sh.getMaxColumns());
  sh.getRange(row, 1, 1, 3).setValues([["'" + app, savedAt, parts.length]]);
  if (parts.length) sh.getRange(row, 4, 1, parts.length).setValues([parts]);
}

function load_(app) {
  const sh = syncSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, data: null };
  const rows = sh.getRange(2, 1, last - 1, Math.max(4, sh.getLastColumn())).getValues();
  for (const r of rows) {
    if (r[0] !== app) continue;
    const n = Number(r[2]) || 0;
    return { ok: true, savedAt: Number(r[1]) || 0, data: r.slice(3, 3 + n).join('') };
  }
  return { ok: true, data: null };
}

/** views = { "Tab name": [[header, ...], [row...], ...] }  Rewrites each tab with the rows given. */
function writeViews_(views) {
  const ss = SpreadsheetApp.getActive();
  Object.keys(views).forEach(function (name) {
    const rows = views[name];
    if (!rows || !rows.length) return;
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.setFrozenRows(1); }
    const w = rows[0].length;
    const clean = rows.map(function (r) {
      return r.slice(0, w).concat(Array(Math.max(0, w - r.length)).fill('')).map(function (v) {
        if (typeof v === 'number' || typeof v === 'boolean') return v;
        const s = v == null ? '' : String(v);
        return s === '' ? '' : "'" + s; // text stays text, never a formula
      });
    });
    const keep = Math.max(sh.getLastRow(), 1);
    sh.getRange(1, 1, keep, Math.max(w, sh.getLastColumn() || 1)).clearContent();
    if (sh.getMaxRows() < clean.length) sh.insertRowsAfter(sh.getMaxRows(), clean.length - sh.getMaxRows());
    sh.getRange(1, 1, clean.length, w).setValues(clean);
  });
}
