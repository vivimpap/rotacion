/**
 * VolApp compartida.
 *
 * Una hoja de cálculo guarda un registro de operaciones, una por fila, en la
 * pestaña «registro». Si el script está dentro de una hoja, usa esa; si no,
 * crea «VolApp datos» en tu Drive (y otra nueva si desaparece). Cada móvil
 * envía sus operaciones y recibe las de los demás; la app reconstruye el
 * equipo, la plantilla y los partidos a partir de ese registro.
 * No edites la pestaña «registro» a mano: para borrarlo todo, usa
 * «Empezar de cero» dentro de la app.
 */
const SHEET_NAME = 'registro';
const PAGE = 2000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('VolApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * The spreadsheet holding the data: the sheet the script was opened from, or
 * else a "VolApp datos" sheet the script creates in Drive. If that sheet can no
 * longer be opened, a new empty one is created.
 */
function spreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const known = openSaved_(props);
  if (known) return known;
  return withLock_(() => {
    const again = openSaved_(props);
    if (again) return again;
    const ss = SpreadsheetApp.create('VolApp datos');
    props.setProperty('SHEET_ID', ss.getId());
    return ss;
  });
}

function openSaved_(props) {
  const id = props.getProperty('SHEET_ID');
  if (!id) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    return null;
  }
}

function sheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['operación', 'cuándo', 'quién']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Identifies the current data set. It changes when the data is wiped or the
 * spreadsheet is replaced, so phones holding older data know to drop it.
 */
function generation_(ss) {
  const props = PropertiesService.getScriptProperties();
  let gen = props.getProperty('GEN');
  if (!gen) {
    gen = String(Date.now());
    props.setProperty('GEN', gen);
  }
  return ss.getId() + ':' + gen;
}

function withLock_(task) {
  const lock = LockService.getScriptLock();
  const had = lock.hasLock();
  if (!had) lock.waitLock(20000);
  try {
    return task();
  } finally {
    if (!had) lock.releaseLock();
  }
}

/** A row that is not a valid operation (edited by hand, say) is skipped, not fatal. */
function parseOp_(cell) {
  try {
    const op = JSON.parse(String(cell));
    return op && typeof op === 'object' ? op : { t: 'skip' };
  } catch (e) {
    return { t: 'skip' };
  }
}

/** Operations after the first `since` ones, which the phone already has. */
function pull(since) {
  since = Number(since) || 0;
  const sh = sheet_();
  const gen = generation_(sh.getParent());
  const total = Math.max(0, sh.getLastRow() - 1);
  if (since > total) return { gen: gen, reset: true, ver: 0, ops: [], more: false };
  const n = Math.min(total - since, PAGE);
  const ops = n > 0 ? sh.getRange(since + 2, 1, n, 1).getValues().map((r) => parseOp_(r[0])) : [];
  return { gen: gen, ver: since + n, ops: ops, more: since + n < total };
}

/** Appends this phone's operations, then returns everything new since `since`. */
function push(ops, since, gen) {
  withLock_(() => {
    const sh = sheet_();
    // Operations recorded against data that has since been wiped are dropped.
    if (gen && gen !== generation_(sh.getParent())) return;
    if (ops && ops.length) {
      const now = new Date();
      sh.getRange(sh.getLastRow() + 1, 1, ops.length, 3)
        .setValues(ops.map((o) => [JSON.stringify(o), now, o.by || '']));
      SpreadsheetApp.flush();
    }
  });
  return pull(since);
}

/** Deletes every team, player and match for everyone. */
function reset() {
  withLock_(() => {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
    PropertiesService.getScriptProperties().setProperty('GEN', String(Date.now()));
    SpreadsheetApp.flush();
  });
  return pull(0);
}
