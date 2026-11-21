/**
 * ZONA NADADORES CN SALOU — Google Apps Script
 *
 * La web lee en directo el Excel de marcas y mínimas, y guarda usuarios, valoraciones,
 * asistencia y entrenamientos en una hoja aparte ("Zona Nadadores CN Salou – Datos web")
 * que se crea al ejecutar instalar().
 *
 * Este script es el servidor de la app (https://cnsalou.github.io/zona-nadadores/).
 */

const CFG = {
  MARCAS_ID: '1cHUn_mu_7fuc_RyTimqZW_Kb08Y--DYYm886-wVZCXc',   // Excel de mínimas y mejores marcas
  PROG_ID: '11XhbiyIoxIy-zU0B29RDvOIw7tru7JI_rH3_FIe_wIk',     // Programación de entrenamientos
  PINES_ID: '1iJe4cOKEiB2u2B_rPLBWYHlhUr__Td2sv8R9ETLm3v8',    // Usuarios y PIN (CONFIDENCIAL): aquí se cambian los PIN
  TITULO: 'Zona Nadadores CN Salou',
  CACHE_SEG: 600,      // los cálculos del Excel se guardan 10 min (se borran solos al editar el Excel)
  SESION_DIAS: 30,     // días que dura la sesión en el móvil antes de pedir el PIN otra vez
  ZONA: 'Europe/Madrid'
};

const T_USERS = 'Usuarios', T_EVAL = 'Valoraciones', T_ATT = 'Asistencia', T_TRAIN = 'Entrenamientos';

// Tiempos base World Aquatics (sexo|piscina|prueba) para calcular puntos
const BASE = {"M|25|50L":19.9,"M|50|50L":20.91,"M|25|100L":44.84,"M|50|100L":46.4,"M|25|200L":98.61,"M|50|200L":102.0,"M|25|400L":212.25,"M|50|400L":219.96,"M|25|800L":440.46,"M|50|800L":452.12,"M|25|1500L":846.88,"M|50|1500L":870.67,"M|25|50E":22.11,"M|50|50E":23.55,"M|25|100E":48.33,"M|50|100E":51.6,"M|25|200E":105.63,"M|50|200E":111.92,"M|25|50B":24.95,"M|50|50B":25.95,"M|25|100B":55.28,"M|50|100B":56.88,"M|25|200B":120.16,"M|50|200B":125.48,"M|25|50M":21.32,"M|50|50M":22.27,"M|25|100M":47.71,"M|50|100M":49.45,"M|25|200M":106.85,"M|50|200M":110.34,"M|25|100S":49.28,"M|25|200S":108.88,"M|50|200S":112.69,"M|25|400S":234.81,"M|50|400S":242.5,"F|25|50L":22.83,"F|50|50L":23.61,"F|25|100L":50.25,"F|50|100L":51.71,"F|25|200L":110.31,"F|50|200L":112.23,"F|25|400L":230.25,"F|50|400L":234.18,"F|25|800L":477.42,"F|50|800L":484.12,"F|25|1500L":908.24,"F|50|1500L":920.48,"F|25|50E":25.23,"F|50|50E":26.86,"F|25|100E":54.02,"F|50|100E":57.13,"F|25|200E":118.04,"F|50|200E":123.14,"F|25|50B":28.37,"F|50|50B":29.16,"F|25|100B":62.36,"F|50|100B":64.13,"F|25|200B":132.5,"F|50|200B":137.55,"F|25|50M":23.94,"F|50|50M":24.43,"F|25|100M":52.71,"F|50|100M":54.6,"F|25|200M":119.32,"F|50|200M":121.81,"F|25|100S":55.11,"F|25|200S":121.63,"F|50|200S":125.7,"F|25|400S":255.48,"F|50|400S":263.65};

/* =========================================================================
 *  PÁGINA WEB
 * ========================================================================= */

// La app (GitHub Pages) llama aquí con POST {fn, args}. Respuesta: {ok:true, data} o {ok:false, error}.
const API = { login: login, getData: getData, saveTrain: saveTrain, saveAtt: saveAtt, saveEval: saveEval, delEval: delEval, getProg: getProg };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API[req.fn];
    if (!fn) throw new Error('Función desconocida.');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: CFG.TITULO })).setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================================
 *  MENÚ EN EL EXCEL E INSTALACIÓN
 * ========================================================================= */

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Zona Nadadores')
      .addItem('Instalar (solo la primera vez)', 'instalar')
      .addItem('Crear usuarios para nadadores nuevos', 'generarUsuarios')
      .addItem('Abrir hoja de datos de la web', 'abrirDatos')
      .addItem('Refrescar la web ahora', 'limpiarCache')
      .addToUi();
  } catch (e) {}
}

// Al editar el Excel, la web deja de usar los datos guardados y vuelve a leerlo.
function onEdit(e) {
  try { CacheService.getScriptCache().remove('D'); } catch (err) {}
}

function alEditarExcel() { try { CacheService.getScriptCache().remove('D'); } catch (e) {} }

// Usuarios y PIN de la web anterior (archivo privado en tu Drive; solo se lee al instalar)
const USUARIOS_FILE_ID = '1UmgkWE1hr1Ksuzxo_cTdp3SsVdulLGne';

function instalar() {
  const USUARIOS_INICIALES = JSON.parse(DriveApp.getFileById(USUARIOS_FILE_ID).getBlob().getDataAsString());
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('DATA_ID')) {
    try { SpreadsheetApp.openById(P.getProperty('DATA_ID')); avisar('Ya estaba instalado. Hoja de datos:\n' + SpreadsheetApp.openById(P.getProperty('DATA_ID')).getUrl()); return; }
    catch (e) { /* la hoja se borró: se vuelve a crear */ }
  }
  const ss = SpreadsheetApp.create(CFG.TITULO + ' – Datos web');
  const first = ss.getSheets()[0];
  const mk = (name, head, widths) => {
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(head.length, 8)).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e3f0f7');
    sh.setFrozenRows(1);
    (widths || []).forEach((w, i) => sh.setColumnWidth(i + 1, w));
    return sh;
  };
  const us = mk(T_USERS, ['Nombre', 'Usuario', 'PIN nuevo', 'Hash (no tocar)', 'Rol'], [240, 180, 100, 120, 100]);
  mk(T_EVAL, ['Id', 'Usuario', 'Competición', 'Fecha', 'Cómo ha ido', 'Aspectos positivos', 'Aspectos a mejorar'], [120, 160, 260, 100, 320, 240, 240]);
  mk(T_ATT, ['Fecha', 'Usuario', 'Estado (p = asistió, f = no asistió)'], [100, 180, 110]);
  mk(T_TRAIN, ['Fecha', 'Usuario', 'Dureza (1-5, 0 = no asistió)', '(sin uso)', 'Comentario', 'Sensaciones (1 = muy malas, 5 = muy buenas)'], [100, 180, 150, 60, 320, 160]);
  ss.deleteSheet(first);
  const rows = USUARIOS_INICIALES.lista.map(r => [r[0], r[1], '', r[2], r[3]]);
  us.getRange(2, 1, rows.length, 5).setValues(rows);
  us.getRange('D:D').setFontColor('#999999');
  us.getRange(1, 3).setNote('Escribe aquí un PIN para dar de alta o cambiar el PIN de alguien. Si hay PIN nuevo, la columna Hash se ignora.');

  P.setProperties({
    DATA_ID: ss.getId(),
    SALT: USUARIOS_INICIALES.salt,
    SECRET: Utilities.getUuid() + Utilities.getUuid()
  });
  // Al editar el Excel de marcas, la app vuelve a leerlo (sirve aunque el script no esté dentro del Excel)
  try {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alEditarExcel').forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('alEditarExcel').forSpreadsheet(CFG.MARCAS_ID).onEdit().create();
  } catch (e) { Logger.log('No se pudo crear el aviso de edición: ' + e); }
  const nuevos = generarUsuarios(true);
  avisar('Instalación hecha.\n\nHoja de datos (solo para ti, no la compartas):\n' + ss.getUrl() +
    (nuevos.length ? '\n\nUsuarios nuevos creados: ' + nuevos.map(n => n.join(' · ')).join('\n') : '') +
    '\n\nSiguiente paso: Implementar → Nueva implementación → Aplicación web (Ejecutar como: Yo · Acceso: Cualquier usuario).');
}

function generarUsuarios(silencio) {
  const D = computeData();
  const sh = pinesSheet();
  const users = usersRows();
  const have = {}; users.forEach(u => have[normName(u.name)] = 1);
  const seen = {}; users.forEach(u => seen[u.user] = 1);
  const out = [];
  D.athletes.forEach(a => {
    if (have[normName(a.name)]) return;
    const t = a.name.trim().split(/\s+/);
    let base = slug(t[0]) + (t[1] ? '.' + slug(t[1]) : ''), u = base, i = 2;
    while (seen[u]) u = base + (i++);
    seen[u] = 1;
    const pin = String(100000 + Math.floor(Math.random() * 900000));
    sh.appendRow([a.name, a.cat || '', a.year || '', u, pin]);
    out.push([a.name, u, 'PIN ' + pin]);
  });
  if (silencio !== true) avisar(out.length ? 'Usuarios creados (el PIN queda en el documento «Usuarios y PIN»):\n\n' + out.map(n => n.join(' · ')).join('\n') : 'No hay nadadores nuevos: todos los del Excel ya tienen usuario.');
  return out;
}

function abrirDatos() { avisar('Hoja de datos de la web:\n' + dataSS().getUrl()); }

function limpiarCache() { CacheService.getScriptCache().remove('D'); avisar('Listo: la web leerá el Excel de nuevo en la próxima visita.'); }

function avisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* =========================================================================
 *  LOGIN Y SESIONES
 * ========================================================================= */

function login(user, pin) {
  user = String(user || '').trim().toLowerCase();
  pin = String(pin || '').trim();
  if (!user || !pin) return { error: 'Escribe usuario y PIN.' };
  const cache = CacheService.getScriptCache(), fk = 'fail_' + user;
  const fails = +(cache.get(fk) || 0);
  if (fails >= 8) return { error: 'Demasiados intentos fallidos. Espera 10 minutos.' };
  const u = usersRows().find(x => x.user === user);
  let ok = false;
  if (u) ok = !!u.pin && u.pin === pin;
  if (!ok) { cache.put(fk, String(fails + 1), 600); return { error: 'Usuario o PIN incorrecto. Revisa que no haya espacios.' }; }
  cache.remove(fk);
  const exp = Date.now() + CFG.SESION_DIAS * 864e5;
  return { token: sign(user + '|' + exp) };
}

function auth(token, needCoach) {
  if (!token || typeof token !== 'string' || token.indexOf('.') < 0) throw new Error('SESION');
  const parts = token.split('.');
  if (hmac(parts[0]) !== parts[1]) throw new Error('SESION');
  const pl = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString().split('|');
  if (+pl[1] < Date.now()) throw new Error('SESION');
  const u = usersRows().find(x => x.user === pl[0]);
  if (!u) throw new Error('SESION');
  const s = { user: u.user, name: u.name, role: u.role === 'entrenador' ? 'coach' : 'athlete' };
  if (needCoach && s.role !== 'coach') throw new Error('Solo el entrenador puede hacer esto.');
  return s;
}

function sign(payload) { const p = Utilities.base64EncodeWebSafe(payload, Utilities.Charset.UTF_8); return p + '.' + hmac(p); }
function hmac(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, prop('SECRET'))).replace(/=+$/, ''); }
function sha256hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function prop(k) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  if (!v) throw new Error('La web no está instalada todavía: ejecuta la función instalar().');
  return v;
}

/* =========================================================================
 *  DATOS PARA LA PÁGINA
 * ========================================================================= */

function getData(token) {
  const s = auth(token);
  const D = buildData();
  const byName = {};
  usersRows().forEach(u => { if (u.role !== 'entrenador') byName[normName(u.name)] = u.user; });
  D.athletes.forEach(a => { a.user = byName[normName(a.name)] || null; a.id = a.user || ('x-' + slug(a.name)); });
  const st = store();
  const base = { role: s.role, user: s.user, name: s.name, mintab: D.mintab, comps: D.comps, ligas: D.ligas, built: D.built };
  if (s.role === 'coach') {
    return Object.assign(base, { athletes: D.athletes, evals: st.evals, train: st.train, att: st.att });
  }
  const me = D.athletes.find(a => a.user === s.user);
  if (!me) return { error: 'Tu usuario existe, pero en el Excel no hay ninguna marca a nombre de "' + s.name + '". Avisa al entrenador.' };
  return Object.assign(base, {
    athletes: [me],
    evals: { [me.id]: st.evals[me.id] || [] },
    train: { [me.id]: st.train[me.id] || {} },
    att: groupAtt(st.att, me, D.athletes),
    rank: teamRank(st.att, st.train, D.athletes, me.id)
  });
}

// Para un nadador: solo su propia asistencia y una marca _g en los días en que su grupo tuvo lista (para calcular los salowis).
function groupAtt(att, me, athletes) {
  const grp = {}; athletes.forEach(a => { if (a.cat === me.cat) grp[a.id] = 1; });
  const out = {};
  Object.keys(att).forEach(d => {
    const day = att[d];
    if (!Object.keys(day).some(k => grp[k])) return;
    out[d] = { _g: 1 };
    if (day[me.id]) out[d][me.id] = day[me.id];
  });
  return out;
}

// Posición de un nadador en el equipo por Salowis de cada mes (empates comparten posición: 1, 1, 3…).
// Salowi = «Asistió» en la lista del entrenador + entreno valorado por el nadador (dureza, sensaciones y comentario).
function teamRank(att, train, athletes, id) {
  const valid = (k, d) => { const v = (train[k] || {})[d]; return !!(v && v.r > 0 && v.s > 0 && String(v.n || '').trim()); };
  const today = Utilities.formatDate(new Date(), CFG.ZONA, 'yyyy-MM-dd');
  const months = {};
  Object.keys(att).forEach(d => {
    const w = new Date(d + 'T12:00:00Z').getUTCDay();
    if (w < 1 || w > 4 || d > today) return;
    const mk = d.slice(0, 7), c = months[mk] || (months[mk] = {});
    Object.keys(att[d]).forEach(k => { if (att[d][k] === 'p' && valid(k, d)) c[k] = (c[k] || 0) + 1; });
  });
  const out = {};
  Object.keys(months).forEach(mk => {
    const c = months[mk], mine = c[id] || 0, vals = athletes.map(a => c[a.id] || 0);
    out[mk] = { pos: 1 + vals.filter(v => v > mine).length, tie: vals.filter(v => v === mine).length - 1, n: athletes.length, got: mine };
  });
  return out;
}

function buildData() {
  const c = CacheService.getScriptCache();
  const z = c.get('D');
  if (z) {
    try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(z), 'application/x-gzip')).getDataAsString()); } catch (e) {}
  }
  const D = computeData();
  try {
    const b = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(D), 'application/json')).getBytes());
    if (b.length < 95000) c.put('D', b, CFG.CACHE_SEG);
  } catch (e) {}
  return D;
}

// Lee el Excel y calcula marcas corregidas, puntos y distancia a la mínima (igual que refresh.py)
function computeData() {
  const ss = SpreadsheetApp.openById(CFG.MARCAS_ID);
  const rowsOf = name => { const sh = ss.getSheetByName(name); return sh ? sh.getDataRange().getDisplayValues() : []; };

  // Mínimas: invierno (CC d'Hivern) y verano (CC d'Estiu). Clave temporada|categoría|sexo|año|prueba|piscina
  const mins = {}, mintab = { junior: [], infantil: [], junior_v: [], infantil_v: [] };
  [['inv', 'junior', 'Mínimas Catalunya Junior'], ['inv', 'infantil', 'Mínimas Catalunya Infantil'],
   ['ver', 'junior', 'Mínimas Junior Verano'], ['ver', 'infantil', 'Mínimas Infantil Verano']].forEach(([se, cat, sh]) => {
    rowsOf(sh).slice(1).forEach(r => {
      const t = tsec(r[5]);
      if (!r[0] || t === null) return;
      const sex = r[1].trim().toUpperCase(), yr = parseInt(r[2], 10), ev = r[3].trim().toUpperCase(), pool = parseInt(r[4], 10);
      mins[[se, cat, sex, yr, ev, pool].join('|')] = t;
      mintab[se === 'inv' ? cat : cat + '_v'].push([sex, yr, ev, pool, fmt(t), r[6] || null]);
    });
  });
  const minInfo = (mn, ct) => ({ min: fmt(mn), ms: mn, pct: round2((ct - mn) / mn * 100), ok: ct <= mn + 1e-9, gap: round2(ct - mn) });

  const athletes = {};
  // small = prebenjamín/benjamín: sí cuentan las pruebas de 25 m y aparecen también los nadadores aún sin marcas
  const load = (sheet, cat, hasMin, small) => {
    rowsOf(sheet).slice(1).forEach(r => {
      const t = tsec(r[4]);
      if (!r[0] || !r[0].trim()) return;
      const name = r[0].trim().replace(/\s+/g, ' '), sex = r[1].trim().toUpperCase(), yr = parseInt(r[2], 10),
        ev = r[3].trim().toUpperCase(), pool = parseInt(r[5], 10), ch = r[6].trim();
      if (t === null) { if (small && !athletes[name]) athletes[name] = { name: name, sex: sex, year: yr, cat: cat, marks: [] }; return; }
      if (!small && /^25[LEBMS]$/.test(ev)) return; // no se muestran pruebas de 25 m (salvo prebenjamín/benjamín)
      const a = athletes[name] || (athletes[name] = { name: name, sex: sex, year: yr, cat: cat, marks: [] });
      const corr = ch === 'Manual' ? (ev.indexOf('50') === 0 ? 0.29 : 0.19) : 0, ct = t + corr;
      const m = { ev: ev, t: fmt(t), s: t, pool: pool, ch: ch, date: dstr(r[7]), place: r[8] || null };
      const b = BASE[sex + '|' + pool + '|' + ev];
      if (b) m.pts = Math.floor(1000 * Math.pow(b / ct, 3) + 1e-9);
      if (hasMin) {
        const mn = mins[['inv', cat, sex, yr, ev, pool].join('|')];
        if (mn) Object.assign(m, minInfo(mn, ct));                 // invierno: campos min, ms, pct, ok, gap
        const mv = mins[['ver', cat, sex, yr, ev, pool].join('|')];
        if (mv) m.v = minInfo(mv, ct);                              // verano: en m.v
      }
      a.marks.push(m);
    });
  };
  load('Junior - Marcas y mínimas', 'junior', true);
  load('Infantil - Marcas y mínimas', 'infantil', true);
  load('Alevín - Marcas', 'alevin', false);
  load('Prebenjamín y Benjamín - Marcas', 'benjamin', false, true);
  load('Máster - Marcas', 'master', false);

  const order = { junior: 0, infantil: 1, alevin: 2, benjamin: 3, master: 4 };
  const list = Object.keys(athletes).map(k => athletes[k])
    .sort((a, b) => order[a.cat] - order[b.cat] || a.name.localeCompare(b.name, 'es'));
  list.forEach(a => a.marks.sort((x, y) => (x.date || '').localeCompare(y.date || '')));

  const cal = rowsOf('Calendario normativas');
  const comps = [], ligas = [];
  for (let i = 4; i <= 16 && i < cal.length; i++) {
    const r = cal[i]; if (!r[0]) continue;
    comps.push({ name: r[0], norm: r[1], cat: r[2], venue: r[3], pool: r[4], start: dstr(r[5]), end: dstr(r[6]), mfrom: dstr(r[7]), mto: dstr(r[8]),
      iopen: dstr(r[9]), iclose: dstr(r[10]), prov: dstr(r[11]), defi: dstr(r[12]), notes: r[13] || null });
  }
  for (let i = 20; i <= 31 && i < cal.length; i++) {
    const r = cal[i]; if (!r[0]) continue;
    ligas.push({ name: r[0], norm: r[1], cat: r[2], j: r[3], start: dstr(r[5]), end: dstr(r[6]), deadline: dstr(r[10]), pub: dstr(r[11]), notes: r[13] || null });
  }
  // Bloque C: competiciones del calendario del club (filas 35-60)
  for (let i = 34; i < cal.length && i < 60; i++) {
    const r = cal[i]; if (!r[0] || !dstr(r[5])) continue;
    comps.push({ name: r[0], norm: '', origin: r[1] || 'Club', cat: r[2], venue: r[3], pool: r[4], start: dstr(r[5]), end: dstr(r[6]) || dstr(r[5]),
      day: r[7] || null, time: r[8] || null, notes: r[13] || null });
  }
  return { athletes: list, mintab: mintab, comps: comps, ligas: ligas, built: Utilities.formatDate(new Date(), CFG.ZONA, 'dd/MM/yyyy HH:mm') };
}

/* =========================================================================
 *  VALORACIONES, ASISTENCIA Y ENTRENAMIENTOS (hoja de datos)
 * ========================================================================= */

function store() {
  const ss = dataSS();
  const read = n => { const sh = ss.getSheetByName(n); const lr = sh.getLastRow(); return lr > 1 ? sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues() : []; };
  const evals = {}, train = {}, att = {};
  read(T_EVAL).forEach(r => {
    if (!r[0] || !r[1]) return;
    (evals[r[1]] = evals[r[1]] || []).push({ id: String(r[0]), comp: String(r[2]), date: normDate(r[3]), general: String(r[4]), pos: String(r[5]), mej: String(r[6]) });
  });
  Object.keys(evals).forEach(k => evals[k].sort((a, b) => (b.date || '').localeCompare(a.date || '') || b.id.localeCompare(a.id)));
  read(T_TRAIN).forEach(r => {
    const d = normDate(r[0]); if (!d || !r[1]) return;
    (train[r[1]] = train[r[1]] || {})[d] = { r: parseInt(r[2], 10) || 0, s: parseInt(r[5], 10) || 0, n: String(r[4] || '') };
  });
  read(T_ATT).forEach(r => {
    const d = normDate(r[0]); const v = String(r[2]).trim().toLowerCase(); if (!d || !r[1] || !/^[pf]$/.test(v)) return;
    (att[d] = att[d] || {})[r[1]] = v;
  });
  return { evals: evals, train: train, att: att };
}

function saveTrain(token, date, val) {
  const s = auth(token);
  if (s.role !== 'athlete') throw new Error('El entrenador solo puede consultar los entrenamientos.');
  checkDate(date);
  let row = null;
  if (val) {
    const r = Math.max(0, Math.min(5, parseInt(val.r, 10) || 0)), f = Math.max(0, Math.min(5, parseInt(val.s, 10) || 0));
    row = [date, s.user, String(r), '', clean(val.n, 1000), String(f)];
  }
  withLock(() => {
    const sh = dataSS().getSheetByName(T_TRAIN);
    if (sh.getRange(1, 3).getValue() !== 'Dureza (1-5, 0 = no asistió)') {
      sh.getRange(1, 3, 1, 4).setValues([['Dureza (1-5, 0 = no asistió)', '(sin uso)', 'Comentario', 'Sensaciones (1 = muy malas, 5 = muy buenas)']]).setFontWeight('bold').setBackground('#e3f0f7');
      sh.getRange('F:F').setNumberFormat('@');
    }
    upsert(sh, [[date, s.user, row]]);
  });
  return { days: store().train[s.user] || {} };
}

function saveAtt(token, date, aids, val) {
  auth(token, true);
  checkDate(date);
  if (val !== null && !/^[pf]$/.test(val)) throw new Error('Valor no válido.');
  withLock(() => upsert(dataSS().getSheetByName(T_ATT), (aids || []).map(a => [date, String(a), val === null ? null : [date, String(a), val]])));
  return { att: store().att };
}

function saveEval(token, aid, ev) {
  auth(token, true);
  checkDate(ev.date);
  const id = 'e' + Date.now();
  withLock(() => dataSS().getSheetByName(T_EVAL).appendRow([id, String(aid), clean(ev.comp, 200), ev.date, clean(ev.general, 4000), clean(ev.pos, 4000), clean(ev.mej, 4000)]));
  return { items: store().evals[aid] || [] };
}

function delEval(token, aid, id) {
  auth(token, true);
  withLock(() => {
    const sh = dataSS().getSheetByName(T_EVAL), lr = sh.getLastRow();
    if (lr < 2) return;
    const ids = sh.getRange(2, 1, lr - 1, 1).getDisplayValues();
    for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) sh.deleteRow(i + 2);
  });
  return { items: store().evals[aid] || [] };
}

// Actualiza o borra filas identificadas por (columna A, columna B)
function upsert(sh, items) {
  const lr = sh.getLastRow();
  const keys = lr > 1 ? sh.getRange(2, 1, lr - 1, 2).getValues() : [];
  const idx = {};
  keys.forEach((r, i) => idx[normDate(r[0]) + '|' + r[1]] = i + 2);
  const del = [];
  items.forEach(([a, b, row]) => {
    const k = a + '|' + b, at = idx[k];
    if (row === null) { if (at) del.push(at); }
    else if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); idx[k] = sh.getLastRow(); }
  });
  del.sort((x, y) => y - x).forEach(r => sh.deleteRow(r));
}

/* =========================================================================
 *  PROGRAMACIÓN DE ENTRENAMIENTOS (solo entrenador)
 * ========================================================================= */

// Devuelve la pestaña con su formato (colores, negrita, alineación, celdas combinadas) para mostrarla en solo lectura.
function getProg(token, tab, rows) {
  auth(token, true);
  const ss = SpreadsheetApp.openById(CFG.PROG_ID);
  const sheets = ss.getSheets();
  let sh = tab ? ss.getSheetByName(tab) : null;
  if (!sh) {
    const wk = isoWeek(new Date());
    sh = sheets.find(s => new RegExp('^S' + wk + '\\b').test(s.getName())) || sheets[0];
  }
  const maxR = Math.max(1, Math.min(parseInt(rows, 10) || 200, sh.getMaxRows())), maxC = Math.min(12, sh.getMaxColumns());
  const rg = sh.getRange(1, 1, maxR, maxC);
  const disp = rg.getDisplayValues();
  let nr = disp.length; while (nr > 0 && disp[nr - 1].every(x => x === '')) nr--;
  let nc = 0; disp.slice(0, nr).forEach(r => r.forEach((x, j) => { if (x !== '') nc = Math.max(nc, j + 1); }));
  if (!nr || !nc) return { tabs: sheets.map(s => s.getName()), tab: sh.getName(), gid: sh.getSheetId(), vals: [] };
  const r2 = sh.getRange(1, 1, nr, nc);
  const cut = a => a.slice(0, nr).map(r => r.slice(0, nc));
  const bg = r2.getBackgrounds().map(r => r.map(c => (c === '#ffffff' ? '' : c)));
  const fc = r2.getFontColorObjects().map(r => r.map(c => { try { const h = c.asRgbColor().asHexString(); return h === '#000000' ? '' : h; } catch (e) { return ''; } }));
  const b = r2.getFontWeights().map(r => r.map(w => (w === 'bold' ? 1 : 0)));
  const al = r2.getHorizontalAlignments();
  const merges = r2.getMergedRanges().map(m => [m.getRow() - 1, m.getColumn() - 1, m.getNumRows(), m.getNumColumns()]);
  const widths = []; for (let c = 1; c <= nc; c++) widths.push(sh.getColumnWidth(c));
  return { tabs: sheets.map(s => s.getName()), tab: sh.getName(), gid: sh.getSheetId(), vals: cut(disp), bg: bg, fc: fc, b: b, al: al, merges: merges, widths: widths };
}

/* =========================================================================
 *  UTILIDADES
 * ========================================================================= */

function dataSS() {
  const id = prop('DATA_ID');
  return SpreadsheetApp.openById(id);
}

// Lee el documento único de usuarios y PIN. Columnas por nombre de cabecera:
// Nombre | Categoría | Año | Usuario | PIN  (Categoría «entrenador» = cuenta de entrenador)
function pinesSheet() { return SpreadsheetApp.openById(CFG.PINES_ID).getSheets()[0]; }

function usersRows() {
  const v = pinesSheet().getDataRange().getDisplayValues();
  const h = v[0].map(x => normName(x));
  const col = n => h.indexOf(normName(n));
  const cN = col('Nombre'), cC = col('Categoría'), cU = col('Usuario'), cP = col('PIN');
  return v.slice(1).filter(r => r[cU] && r[cU].trim()).map(r => {
    const cat = String(r[cC] || '').trim().toLowerCase();
    return { name: String(r[cN] || '').replace(/\s*\(entrenador\)\s*$/i, '').trim(), user: r[cU].trim().toLowerCase(),
      pin: String(r[cP] || '').trim(), hash: '', role: cat === 'entrenador' ? 'entrenador' : 'nadador' };
  });
}

function withLock(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

function tsec(s) {
  s = String(s == null ? '' : s).trim().replace(',', '.');
  if (!s) return null;
  const p = s.split(':');
  if (p.some(x => x === '' || isNaN(+x))) return null;
  let t = 0; p.forEach(x => t = t * 60 + (+x));
  return t > 0 ? round2(t) : null;
}
function fmt(s) { const m = Math.floor(s / 60), x = s - m * 60; return ('0' + m).slice(-2) + ':' + ('0' + x.toFixed(2)).slice(-5); }
function round2(x) { return Math.round(x * 100) / 100; }
function dstr(x) {
  x = String(x == null ? '' : x).trim();
  let m = x.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if (/^\d{4}-\d{2}-\d{2}$/.test(x)) return x;
  return null;
}
function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.ZONA, 'yyyy-MM-dd');
  return dstr(v);
}
function checkDate(d) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d))) throw new Error('Fecha no válida.'); }
function clean(s, max) { s = String(s == null ? '' : s).slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function slug(s) { return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z]/g, ''); }
function normName(s) { return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day);
  return Math.ceil(((t - new Date(Date.UTC(t.getUTCFullYear(), 0, 1))) / 864e5 + 1) / 7);
}
