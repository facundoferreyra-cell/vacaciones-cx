/**
 * Vacaciones CX — backend en Google Apps Script
 * ------------------------------------------------
 * Guarda todo en las pestañas de este Google Sheet:
 *   Config     → clave/valor (título, período, regla, criterios, abierto)
 *   Equipo     → una fila por persona (id, nombre, turno, área, activo, fecha_ingreso, metrica)
 *   Pedidos    → una fila por opción pedida (opción 1, 2, 3) o una fila "sin fechas definidas"
 *   Forzados   → aprobaciones/rechazos manuales de coordinación
 *   Resultados → la última asignación publicada, tal como la ve el equipo
 *
 * Cómo publicarlo: Implementar → Nueva implementación → App web →
 *   Ejecutar como: yo · Quién tiene acceso: Cualquier usuario → Implementar → copiar la URL.
 * Cada vez que cambies este código: Implementar → Administrar implementaciones → editar → versión nueva.
 */

// ▼▼ CAMBIÁ ESTAS DOS CLAVES ▼▼
const TEAM_CODE  = 'cx2027';       // la que compartís con el equipo
const ADMIN_CODE = 'coordinacion'; // la tuya, para el panel de asignación
// ▲▲ ------------------------ ▲▲

const SH = { config: 'Config', team: 'Equipo', requests: 'Pedidos', overrides: 'Forzados', results: 'Resultados' };
const HEAD = {
  config:    ['clave', 'valor'],
  team:      ['id', 'nombre', 'turno', 'area', 'activo', 'fecha_ingreso', 'metrica'],
  requests:  ['id', 'persona_id', 'opcion', 'desde', 'hasta', 'comentario', 'creado', 'origen', 'sin_fechas'],
  overrides: ['clave_pedido', 'estado', 'actualizado'],
  results:   ['publicado', 'persona_id', 'opcion', 'desde', 'hasta', 'estado', 'motivo']
};

function doGet(e)  { return respond(handle((e && e.parameter) || {})); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (_) {}
  return respond(handle(body));
}
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle(p) {
  try {
    const action = p.action || 'state';
    if (action === 'ping') return { ok: true };
    const isAdmin = p.code === ADMIN_CODE;
    const isTeam = isAdmin || p.code === TEAM_CODE;
    if (!isTeam) return { ok: false, error: 'bad_code' };
    const adminOnly = ['saveConfig', 'saveTeam', 'override', 'publish', 'setOpen', 'import', 'removePerson'];
    if (adminOnly.indexOf(action) >= 0 && !isAdmin) return { ok: false, error: 'admin_only' };

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      ensureSheets();
      switch (action) {
        case 'state':       break;
        case 'setRequest':  setRequest(p, isAdmin); break;
        case 'clearRequest': clearRequest(p, isAdmin); break;
        case 'removePerson': clearRequest(p, true); break;
        case 'saveConfig':  saveConfig(p.config || {}); break;
        case 'saveTeam':    saveTeam(p.team || []); break;
        case 'override':    setOverride(p.key, p.status); break;
        case 'publish':     publish(p.items || []); break;
        case 'setOpen':     setConfigKey('open', !!p.open); break;
        case 'import':      importRows(p.rows || []); break;
        default: return { ok: false, error: 'unknown_action' };
      }
      const st = readState(isAdmin, p.personId);
      st.ok = true; st.isAdmin = isAdmin;
      return st;
    } finally { lock.releaseLock(); }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

// ---------- sheets ----------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(key) {
  let s = ss().getSheetByName(SH[key]);
  if (!s) { s = ss().insertSheet(SH[key]); }
  return s;
}
function ensureSheets() {
  Object.keys(SH).forEach(k => {
    const s = sheet(k);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
      s.setFrozenRows(1);
    } else {
      const cur = s.getRange(1, 1, 1, HEAD[k].length).getValues()[0].map(String);
      if (cur.join('|') !== HEAD[k].join('|')) s.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
    }
    // todo como texto plano para que Sheets no convierta fechas ni números
    s.getRange(1, 1, Math.max(s.getMaxRows(), 2), Math.max(s.getMaxColumns(), HEAD[k].length)).setNumberFormat('@');
  });
  const cfg = readConfig();
  if (cfg.title === undefined) {
    // valores iniciales de la primera vez
    saveConfig({ title: 'Vacaciones', periodStart: '', periodEnd: '', rule: 'shift', maxSimul: 1, maxDays: 0,
      criteria: ['antiguedad', 'metricas', 'llegada'], note: '', open: true });
  }
}
function rows(key) {
  const s = sheet(key);
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, HEAD[key].length).getValues().map(r => r.map(cellStr)).filter(r => r.some(v => v !== ''));
}
function cellStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (v === null || v === undefined) return '';
  return String(v);
}
function writeRows(key, data) {
  const s = sheet(key);
  const n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, s.getMaxColumns()).clearContent();
  if (data.length) s.getRange(2, 1, data.length, HEAD[key].length).setNumberFormat('@').setValues(data);
}
function appendRow(key, row) {
  const s = sheet(key);
  s.getRange(s.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function bool(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === '1' || v === 1; }
function now() { return new Date().toISOString(); }

// ---------- config ----------
function readConfig() {
  const cfg = {};
  rows('config').forEach(r => { if (r[0]) { try { cfg[r[0]] = JSON.parse(r[1]); } catch (_) { cfg[r[0]] = r[1]; } } });
  return cfg;
}
function saveConfig(cfg) {
  const cur = readConfig();
  const keep = ['title', 'periodStart', 'periodEnd', 'rule', 'maxSimul', 'maxDays', 'criteria', 'note', 'open', 'publishedAt'];
  keep.forEach(k => { if (cfg[k] !== undefined) cur[k] = cfg[k]; });
  writeRows('config', Object.keys(cur).map(k => [k, JSON.stringify(cur[k])]));
}
function setConfigKey(k, v) { const o = {}; o[k] = v; saveConfig(o); }

// ---------- equipo ----------
function readTeam(withScores) {
  return rows('team').filter(r => r[0] && r[1]).map(r => {
    const p = { id: r[0], name: r[1], shift: r[2] === 'pm' ? 'pm' : 'am', area: r[3], active: r[4] === '' ? true : bool(r[4]) };
    if (withScores) { p.seniorityDate = r[5] || ''; p.metric = r[6] === '' ? '' : Number(r[6]); }
    return p;
  });
}
function saveTeam(team) {
  writeRows('team', team.filter(p => p && p.id && p.name).map(p => [
    String(p.id), String(p.name), p.shift === 'pm' ? 'pm' : 'am', String(p.area || ''), p.active === false ? 'FALSE' : 'TRUE',
    String(p.seniorityDate || ''), (p.metric === '' || p.metric === null || p.metric === undefined || isNaN(Number(p.metric))) ? '' : String(Number(p.metric))
  ]));
}

// ---------- pedidos ----------
// Cada persona tiene UN pedido: hasta 3 opciones de fechas en orden de preferencia
// (una fila por opción), o una sola fila "sin fechas definidas".
function readRequests() {
  return rows('requests').filter(r => r[0] && r[1]).map(r => ({
    id: r[0], personId: r[1], option: Number(r[2]) || 0, from: r[3], to: r[4], note: r[5], createdAt: r[6], source: r[7] || 'self', flexible: bool(r[8])
  }));
}
function reqRow(r) { return [r.id, r.personId, String(r.option || 0), r.from || '', r.to || '', r.note || '', r.createdAt || '', r.source || 'self', r.flexible ? 'TRUE' : 'FALSE']; }
function validDate(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function newId() { return Utilities.getUuid().slice(0, 8); }
function setRequest(p, isAdmin) {
  const cfg = readConfig();
  if (cfg.open === false && !isAdmin) throw new Error('La postulación está cerrada.');
  if (!p.personId) throw new Error('Falta la persona.');
  const note = String(p.note || '').slice(0, 300);
  const all = readRequests();
  const mine = all.filter(r => r.personId === p.personId);
  // el orden de llegada se conserva: si ya había pedido, se mantiene su fecha original
  const createdAt = mine.length ? mine.map(r => r.createdAt).filter(Boolean).sort()[0] || now() : now();
  const source = p.source || (mine[0] && mine[0].source) || 'self';
  const rest = all.filter(r => r.personId !== p.personId);
  const fresh = [];
  if (p.flexible) {
    fresh.push({ id: newId(), personId: p.personId, option: 0, from: '', to: '', note, createdAt, source, flexible: true });
  } else {
    const opts = (p.options || []).filter(o => o && (o.from || o.to)).slice(0, 3);
    if (!opts.length) throw new Error('Cargá al menos una opción de fechas.');
    opts.forEach((o, i) => {
      if (!validDate(o.from) || !validDate(o.to)) throw new Error('La opción ' + (i + 1) + ' tiene una fecha incompleta.');
      if (o.to < o.from) throw new Error('En la opción ' + (i + 1) + ' la fecha "hasta" es anterior a "desde".');
      fresh.push({ id: newId(), personId: p.personId, option: i + 1, from: o.from, to: o.to, note: i === 0 ? note : '', createdAt, source, flexible: false });
    });
  }
  writeRows('requests', rest.concat(fresh).map(reqRow));
}
function clearRequest(p, isAdmin) {
  const cfg = readConfig();
  if (cfg.open === false && !isAdmin) throw new Error('La postulación está cerrada.');
  if (!p.personId) throw new Error('Falta la persona.');
  writeRows('requests', readRequests().filter(r => r.personId !== p.personId).map(reqRow));
}
function importRows(list) {
  // list: [{personId, options:[{from,to}], note}] — reemplaza el pedido de cada persona
  list.forEach(r => { try { setRequest({ personId: r.personId, options: r.options || [], note: r.note || '', source: 'import' }, true); } catch (_) {} });
}

// ---------- forzados y resultados ----------
function readOverrides() {
  const o = {};
  rows('overrides').forEach(r => { if (r[0] && (r[1] === 'approved' || r[1] === 'rejected')) o[r[0]] = r[1]; });
  return o;
}
function setOverride(key, status) {
  if (!key) return;
  const o = readOverrides();
  if (status === 'approved' || status === 'rejected') o[key] = status; else delete o[key];
  const ts = now();
  writeRows('overrides', Object.keys(o).map(k => [k, o[k], ts]));
}
function publish(items) {
  const ts = now();
  writeRows('results', items.map(it => [ts, it.personId, String(it.option || 0), it.from, it.to, it.status, it.reason || '']));
  setConfigKey('publishedAt', ts);
}
function readResults() {
  const cfg = readConfig();
  const list = rows('results').filter(r => r[1]).map(r => ({ personId: r[1], option: Number(r[2]) || 0, from: r[3], to: r[4], status: r[5], reason: r[6] }));
  if (!cfg.publishedAt) return null;
  return { publishedAt: cfg.publishedAt, items: list };
}

// ---------- estado completo ----------
// El equipo solo recibe SU propio pedido y SU resultado (personId elegido en la página);
// los pedidos del resto solo viajan a coordinación.
function readState(isAdmin, personId) {
  const cfg = readConfig();
  const onlyMine = list => isAdmin ? list : list.filter(r => personId && r.personId === personId);
  const results = readResults();
  const st = {
    config: {
      title: cfg.title || 'Vacaciones', periodStart: cfg.periodStart || '', periodEnd: cfg.periodEnd || '', rule: cfg.rule || 'shift',
      maxSimul: cfg.maxSimul || 1, maxDays: cfg.maxDays || 0, criteria: cfg.criteria || ['antiguedad', 'metricas', 'llegada'],
      note: cfg.note || '', open: cfg.open !== false
    },
    team: readTeam(isAdmin),
    requests: onlyMine(readRequests()),
    results: results ? { publishedAt: results.publishedAt, items: onlyMine(results.items) } : null,
    serverTime: now()
  };
  if (isAdmin) st.overrides = readOverrides();
  return st;
}
