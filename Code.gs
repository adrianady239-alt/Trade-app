/**
 * Trade App – server (Google Apps Script) pentru „TRANZACTII + Backtest”.
 * Pune parola ta la TOKEN, apoi: Implementare → Implementare nouă → Aplicație web
 * (Execută ca: Eu, Acces: Oricine). URL-ul /exec + parola se pun în aplicație la Setări → Server.
 */
const TOKEN = 'SCHIMBA_PAROLA';

const ALLOWED = ['🔆LIVE - Executii', '📊 TRADES OF THE  WEEK', 'DAX (500)', 'GBPUSD (500)', 'UK100'];
const DATA_SHEET = 'APP_DATA';

function doGet() {
  return out({ ok: true, msg: 'Trade App server activ. Folosește URL-ul în aplicație.' });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    if (String(p.token || '') !== TOKEN) return out({ ok: false, code: 'bad_token', error: 'Parolă greșită.' });
    switch (p.action) {
      case 'ping':   return out({ ok: true, rev: rev() });
      case 'get':    return out({ ok: true, values: getRange(p.range), rev: rev() });
      case 'update': return out(updateRange(p.range, p.values, false));
      case 'edit':   return out(updateRange(p.range, p.values, true));
      case 'list':   return out({ ok: true, items: listDocs(p.prefix || ''), rev: rev() });
      case 'doc':    return out({ ok: true, item: getDoc(p.id) });
      case 'set':    return out(withLock(() => { putDoc(p.id, p.data, false); return { ok: true }; }));
      case 'merge':  return out(withLock(() => { putDoc(p.id, p.data, true); return { ok: true }; }));
      case 'delete': return out(withLock(() => { delDoc(p.id); return { ok: true }; }));
      default:       return out({ ok: false, code: 'tool_error', error: 'Acțiune necunoscută: ' + p.action });
    }
  } catch (err) {
    return out({ ok: false, code: 'tool_error', error: String(err && err.message || err) });
  }
}

/* ---------- utilitare ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function rev() { return Number(PropertiesService.getScriptProperties().getProperty('rev') || 0); }
function bump() { const p = PropertiesService.getScriptProperties(); p.setProperty('rev', String(rev() + 1)); }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// "'🔆LIVE - Executii'!A7:S600" → { sheet, a1 }
function parseRange(range) {
  range = String(range || '');
  const k = range.lastIndexOf('!');
  if (k < 0) throw new Error('Range invalid: ' + range);
  let name = range.slice(0, k).trim();
  if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
  if (ALLOWED.indexOf(name) < 0) throw new Error('Foaie nepermisă: ' + name);
  const sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Foaia nu există: ' + name);
  return { sh: sh, a1: range.slice(k + 1) };
}

/* ---------- citire / scriere tradeuri ---------- */
function getRange(range) {
  const r = parseRange(range);
  const rng = r.sh.getRange(r.a1);
  const vals = rng.getDisplayValues();
  // DATA (coloana C) se trimite mereu ca zz.ll.aaaa, indiferent de formatul celulei (evită confuzia lună/zi)
  const cIdx = 3 - rng.getColumn();
  if (cIdx >= 0 && cIdx < rng.getNumColumns()) {
    const raw = r.sh.getRange(rng.getRow(), 3, rng.getNumRows(), 1).getValues();
    const tz = ss().getSpreadsheetTimeZone();
    for (let i = 0; i < raw.length; i++) {
      if (raw[i][0] instanceof Date) vals[i][cIdx] = Utilities.formatDate(raw[i][0], tz, 'dd.MM.yyyy');
    }
  }
  let last = vals.length - 1;
  while (last >= 0 && vals[last].every(v => v === '')) last--;
  return vals.slice(0, last + 1).map(row => {
    let e = row.length - 1;
    while (e >= 0 && row[e] === '') e--;
    return row.slice(0, e + 1);
  });
}

function updateRange(range, values, allowOverwrite) {
  if (!Array.isArray(values) || !values.length) throw new Error('Nimic de scris.');
  return withLock(() => {
    const r = parseRange(range);
    const rng = r.sh.getRange(r.a1);
    const row0 = rng.getRow(), col0 = rng.getColumn();
    // la adăugare nu suprascrie rânduri care au deja DATA (coloana C); la editare, da
    if (!allowOverwrite) {
      const dates = r.sh.getRange(row0, 3, values.length, 1).getDisplayValues();
      for (let i = 0; i < dates.length; i++) {
        if (String(dates[i][0]).trim()) return { ok: false, code: 'tool_error', error: 'Rândul ' + (row0 + i) + ' are deja DATA. Nu am scris nimic.' };
      }
    }
    const width = Math.max.apply(null, values.map(v => v.length));
    const fmts = [];
    const fixed = values.map(row => {
      const out = [], f = [];
      for (let c = 0; c < width; c++) {
        const conv = typed(col0 + c, row[c] == null ? '' : row[c]);
        out.push(conv[0]); f.push(conv[1]);
      }
      fmts.push(f);
      return out;
    });
    const block = r.sh.getRange(row0, col0, fixed.length, width);
    block.setNumberFormats(fmts);
    block.setValues(fixed);
    SpreadsheetApp.flush();
    bump();
    return { ok: true, row: row0, rows: fixed.length };
  });
}

// Valori tipizate pe coloană: [valoare, format]. Datele și orele ca număr serial (fără decalaj de fus orar).
function typed(col, v) {
  const s = String(v).trim();
  if (col === 3) {                        // C DATA: zz.ll.aaaa → dată
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return [Date.UTC(+m[3], +m[2] - 1, +m[1]) / 864e5 + 25569, 'dd.mm.yyyy'];
  }
  if (col === 13) {                       // M PROFIT %: "2%" → 0.02
    const m = s.match(/^(-?\d+(?:[.,]\d+)?)%$/);
    if (m) return [parseFloat(m[1].replace(',', '.')) / 100, '0%'];
  }
  if (col === 14 && s !== '' && !isNaN(+s)) return [+s, '0'];   // N RR
  if (col === 15) {                       // O ORA: hh:mm → oră
    const m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (m) return [(+m[1] * 60 + +m[2]) / 1440, 'hh:mm'];
  }
  return [v, '@'];                        // restul: text exact, nu se interpretează
}

/* ---------- lista „De trimis” + setări (foaia ascunsă APP_DATA) ---------- */
function dataSheet() {
  let sh = ss().getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss().insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'json', 'updatedAt']]);
    sh.hideSheet();
  }
  return sh;
}
function allRows() {
  const sh = dataSheet(), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues();
}
function listDocs(prefix) {
  return allRows().filter(r => r[0] && String(r[0]).indexOf(prefix) === 0)
    .map(r => ({ id: String(r[0]), data: JSON.parse(r[1] || '{}') }));
}
function getDoc(id) {
  const r = allRows().find(x => String(x[0]) === id);
  return r ? { id: id, data: JSON.parse(r[1] || '{}') } : null;
}
function putDoc(id, data, merge) {
  if (!id) throw new Error('Lipsește id.');
  const sh = dataSheet(), rows = allRows();
  const i = rows.findIndex(x => String(x[0]) === id);
  let obj = data || {};
  if (merge && i >= 0) obj = Object.assign(JSON.parse(rows[i][1] || '{}'), obj);
  const line = [[id, JSON.stringify(obj), new Date()]];
  if (i >= 0) sh.getRange(i + 2, 1, 1, 3).setValues(line);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues(line);
  bump();
}
function delDoc(id) {
  const rows = allRows(), i = rows.findIndex(x => String(x[0]) === id);
  if (i >= 0) { dataSheet().deleteRow(i + 2); bump(); }
}
