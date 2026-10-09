/**
 * TRIMESTRES IVA — FuturMovil
 * Backend API (Apps Script) para https://puntofibra.github.io/trimestre/
 * Hoja: trimestre  ·  Pestañas: clave, Tareas, Archivos
 * Claves válidas: cualquier celda de la columna A de la pestaña "clave".
 * Implementar como App web → Ejecutar como: Yo · Acceso: Cualquier usuario.
 */

const SHEET_ID = '1lQwinJ1teQQx1Hf3VveZiaagdeJ40w-q9TiP1nXdViE';
const H_TAREAS = ['id', 'trimestre', 'titulo', 'notas', 'icono', 'hecho', 'fechaHecho', 'orden', 'creado'];
const H_ARCH = ['id', 'tareaId', 'trimestre', 'nombre', 'fileId', 'url', 'mime', 'size', 'fecha'];

/* ---------- Entrada ---------- */
function doGet(e) {
  return out_({ ok: true, app: 'trimestre', hora: new Date().toISOString() });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'JSON inválido' }); }
  try {
    if (!claveValida_(req.pass)) return out_({ ok: false, error: 'CLAVE', msg: 'Clave incorrecta' });
    const fn = ACCIONES[req.accion];
    if (!fn) return out_({ ok: false, error: 'Acción desconocida: ' + req.accion });
    const lectura = ['login', 'datos'].indexOf(req.accion) >= 0;
    if (lectura) return out_(fn(req));
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return out_(fn(req)); } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

const ACCIONES = {
  login: () => ({ ok: true }),
  datos: datos_,
  add: add_,
  edit: edit_,
  del: del_,
  toggle: toggle_,
  reorder: reorder_,
  copiar: copiar_,
  subir: subir_,
  borrarArchivo: borrarArchivo_,
  carpeta: (r) => ({ ok: true, url: carpetaTri_(r.tri, true).getUrl() })
};

/* ---------- Claves ---------- */
function claveValida_(pass) {
  pass = String(pass || '').trim();
  if (!pass) return false;
  const sh = ss_().getSheetByName('clave');
  if (!sh || sh.getLastRow() < 1) return false;
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
  return vals.some(r => String(r[0]).trim() !== '' && String(r[0]).trim() === pass);
}

/* ---------- Lectura ---------- */
function datos_(r) {
  const tri = r.tri;
  const tareas = filas_('Tareas', H_TAREAS);
  const archivos = filas_('Archivos', H_ARCH);
  const resumen = {};
  tareas.forEach(t => {
    const k = t.trimestre;
    resumen[k] = resumen[k] || { total: 0, hechas: 0 };
    resumen[k].total++;
    if (t.hecho) resumen[k].hechas++;
  });
  const carpeta = tri ? carpetaTri_(tri, false) : null;
  return {
    ok: true,
    tri: tri,
    tareas: tareas.filter(t => t.trimestre === tri).sort((a, b) => a.orden - b.orden),
    archivos: archivos.filter(a => a.trimestre === tri),
    resumen: resumen,
    carpeta: carpeta ? carpeta.getUrl() : ''
  };
}

/* ---------- Tareas ---------- */
function add_(r) {
  const sh = hoja_('Tareas', H_TAREAS);
  const titulo = String(r.titulo || '').trim();
  if (!titulo) throw new Error('Falta el título');
  const max = filas_('Tareas', H_TAREAS).filter(t => t.trimestre === r.tri).reduce((m, t) => Math.max(m, t.orden), 0);
  const t = { id: uid_(), trimestre: r.tri, titulo: titulo, notas: String(r.notas || ''), icono: r.icono || '📄', hecho: false, fechaHecho: '', orden: max + 1, creado: new Date() };
  sh.appendRow(H_TAREAS.map(h => t[h]));
  return { ok: true, tarea: limpia_(t) };
}

function edit_(r) {
  return actualiza_(r.id, t => {
    if (r.titulo !== undefined) t.titulo = String(r.titulo).trim() || t.titulo;
    if (r.notas !== undefined) t.notas = String(r.notas);
    if (r.icono !== undefined) t.icono = r.icono;
  });
}

function toggle_(r) {
  return actualiza_(r.id, t => {
    t.hecho = !!r.hecho;
    t.fechaHecho = r.hecho ? new Date() : '';
  });
}

function reorder_(r) {
  const sh = hoja_('Tareas', H_TAREAS);
  const data = sh.getDataRange().getValues();
  const pos = {};
  (r.ids || []).forEach((id, i) => pos[String(id)] = i + 1);
  const col = H_TAREAS.indexOf('orden');
  for (let i = 1; i < data.length; i++) {
    if (pos[String(data[i][0])]) sh.getRange(i + 1, col + 1).setValue(pos[String(data[i][0])]);
  }
  return { ok: true };
}

function del_(r) {
  const sh = hoja_('Tareas', H_TAREAS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(r.id)) { sh.deleteRow(i + 1); break; }
  }
  // Archivos de la tarea: a la papelera de Drive (recuperables) y fuera de la lista
  const sa = hoja_('Archivos', H_ARCH);
  const da = sa.getDataRange().getValues();
  for (let i = da.length - 1; i >= 1; i--) {
    if (String(da[i][1]) === String(r.id)) {
      try { DriveApp.getFileById(da[i][4]).setTrashed(true); } catch (e) {}
      sa.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

function copiar_(r) {
  const todas = filas_('Tareas', H_TAREAS);
  const yaHay = todas.filter(t => t.trimestre === r.to).map(t => t.titulo.toLowerCase());
  const origen = todas.filter(t => t.trimestre === r.from).sort((a, b) => a.orden - b.orden);
  const sh = hoja_('Tareas', H_TAREAS);
  let orden = todas.filter(t => t.trimestre === r.to).reduce((m, t) => Math.max(m, t.orden), 0);
  let n = 0;
  origen.forEach(o => {
    if (yaHay.indexOf(o.titulo.toLowerCase()) >= 0) return;
    orden++; n++;
    sh.appendRow([uid_(), r.to, o.titulo, o.notas, o.icono, false, '', orden, new Date()]);
  });
  return { ok: true, copiadas: n };
}

function actualiza_(id, fn) {
  const sh = hoja_('Tareas', H_TAREAS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const t = {};
      H_TAREAS.forEach((h, j) => t[h] = data[i][j]);
      fn(t);
      sh.getRange(i + 1, 1, 1, H_TAREAS.length).setValues([H_TAREAS.map(h => t[h])]);
      return { ok: true, tarea: limpia_(t) };
    }
  }
  throw new Error('Tarea no encontrada');
}

/* ---------- Archivos ---------- */
function subir_(r) {
  const tarea = filas_('Tareas', H_TAREAS).find(t => t.id === String(r.tareaId));
  if (!tarea) throw new Error('Tarea no encontrada');
  const carpeta = carpetaTri_(tarea.trimestre, true);
  const bytes = Utilities.base64Decode(r.data);
  const nombre = (tarea.titulo + ' - ' + r.nombre).replace(/[\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(bytes, r.mime || 'application/octet-stream', nombre);
  const f = carpeta.createFile(blob);
  f.setDescription('Trimestre ' + tarea.trimestre + ' · ' + tarea.titulo);
  const a = { id: uid_(), tareaId: tarea.id, trimestre: tarea.trimestre, nombre: r.nombre, fileId: f.getId(), url: f.getUrl(), mime: r.mime || '', size: bytes.length, fecha: new Date() };
  hoja_('Archivos', H_ARCH).appendRow(H_ARCH.map(h => a[h]));
  return { ok: true, archivo: limpia_(a), carpeta: carpeta.getUrl() };
}

function borrarArchivo_(r) {
  const sh = hoja_('Archivos', H_ARCH);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(r.id)) {
      try { DriveApp.getFileById(data[i][4]).setTrashed(true); } catch (e) {}
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Archivo no encontrado');
}

/** Carpeta "2026 T3" dentro de la carpeta donde está el Sheet. */
function carpetaTri_(tri, crear) {
  const m = /^(\d{4})-T([1-4])$/.exec(String(tri || ''));
  if (!m) throw new Error('Trimestre inválido');
  const nombre = m[1] + ' T' + m[2];
  const padre = DriveApp.getFileById(SHEET_ID).getParents().next();
  const it = padre.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  return crear ? padre.createFolder(nombre) : null;
}

/* ---------- Utilidades ---------- */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function hoja_(nombre, cab) {
  const ss = ss_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1e1b4b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
  }
  return sh;
}

function filas_(nombre, cab) {
  const sh = hoja_(nombre, cab);
  const data = sh.getDataRange().getValues();
  const res = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const o = {};
    cab.forEach((h, j) => o[h] = data[i][j]);
    res.push(limpia_(o));
  }
  return res;
}

function limpia_(o) {
  const r = {};
  Object.keys(o).forEach(k => {
    let v = o[k];
    if (v instanceof Date) v = v.toISOString();
    if (k === 'id' || k === 'tareaId') v = String(v);
    if (k === 'hecho') v = v === true || v === 'TRUE' || v === 'true';
    if (k === 'orden' || k === 'size') v = Number(v) || 0;
    r[k] = v;
  });
  return r;
}

// Siempre empieza por letra: un id solo de cifras (p. ej. 12345678 o 1e234567) la hoja lo convertía en número
function uid_() { return 't' + Utilities.getUuid().replace(/-/g, '').slice(0, 9); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Instalación (ejecutar una vez desde el editor) ---------- */
function setup() {
  const ss = ss_();
  let c = ss.getSheetByName('clave');
  if (!c) { c = ss.insertSheet('clave'); c.getRange('B1').setValue('← Escribe en la columna A las claves válidas (una por celda)'); }
  hoja_('Tareas', H_TAREAS);
  hoja_('Archivos', H_ARCH);
  const h1 = ss.getSheetByName('Hoja 1');
  if (h1 && h1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h1);

  // Tareas iniciales del 3er trimestre 2026
  const tri = '2026-T3';
  if (!filas_('Tareas', H_TAREAS).some(t => t.trimestre === tri)) {
    const sh = hoja_('Tareas', H_TAREAS);
    [
      ['🏠', 'Facturas del alquiler del bajo', 'Las 3 mensualidades del trimestre'],
      ['📱', 'Facturas comisiones Simyo', ''],
      ['📡', 'Facturas comisiones O2', ''],
      ['🛒', 'Facturas de Disashop', ''],
      ['👤', 'Facturas cliente Wilmer', ''],
      ['👤', 'Facturas cliente Johanna', '']
    ].forEach((x, i) => sh.appendRow([uid_(), tri, x[1], x[2], x[0], false, '', i + 1, new Date()]));
  }
  carpetaTri_(tri, true);
  Logger.log('Listo. Carpeta: ' + carpetaTri_(tri, false).getUrl());
}

/** Repara ids que la hoja convirtió en número (ejecutar una vez). */
function repararIds() {
  const st = hoja_('Tareas', H_TAREAS), sa = hoja_('Archivos', H_ARCH);
  st.getRange('A:C').setNumberFormat('@'); sa.getRange('A:C').setNumberFormat('@');
  const dt = st.getDataRange().getValues(), da = sa.getDataRange().getValues();
  const mapa = {};
  let n = 0;
  for (let i = 1; i < dt.length; i++) {
    if (dt[i][0] !== '' && typeof dt[i][0] !== 'string') {
      const nuevo = uid_(); mapa[String(dt[i][0])] = nuevo;
      st.getRange(i + 1, 1).setValue(nuevo); n++;
    }
  }
  for (let i = 1; i < da.length; i++) {
    const k = String(da[i][1]);
    if (mapa[k]) sa.getRange(i + 1, 2).setValue(mapa[k]);
    if (da[i][0] !== '' && typeof da[i][0] !== 'string') sa.getRange(i + 1, 1).setValue(uid_());
  }
  Logger.log('Ids reparados: ' + n);
}
