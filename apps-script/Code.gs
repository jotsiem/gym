/**
 * Gym v2 — servidor de la web app (Apps Script ligado al Sheet "Gym v2").
 * Pestañas que usa: Log, Resumen, Sesiones, Ejercicios, Rutinas, Config, Dashboard, Semanal.
 */

// Columnas de Log (1-based)
var LOG = { id: 1, sesion: 2, fecha: 3, est: 4, orden: 5, ej: 6, nombre: 7, variante: 8, serie: 9, tramo: 10, tipo: 11,
            peso: 12, reps: 13, carga: 14, e1rm: 15, notas: 16, origen: 17, texto: 18, borrado: 19, volumen: 20, client: 21 };
var TZ = 'Europe/Madrid';
var DIAS_HISTORIAL = 400;
var FILAS_RECIENTES = 2500;   // ~16 meses de series: suficiente para sugerencias e historial, mucho más rápido que leer todo el Log
var COL_NOTA_EJ = 14;         // Ejercicios!N: nota fijada por ejercicio
var COL_SES_DUR = 10, COL_SES_DESC = 11;   // Sesiones!J y K: duración (min) y descanso medio (s)

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Gym')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black');
}

// ------------------------------------------------------------------ API (única puerta de entrada)
/** Funciones que la app puede llamar. Cualquier otra se rechaza. */
function apiFns_() {
  return { getEstado: getEstado_, guardarSerie: guardarSerie_, borrarSerie: borrarSerie_, editarSerie: editarSerie_,
           crearEjercicio: crearEjercicio_, getProgreso: getProgreso_, getHistorial: getHistorial_, descartarSesion: descartarSesion_,
           guardarNotaEjercicio: guardarNotaEjercicio_, cerrarSesion: cerrarSesion_, getEstancados: getEstancados_,
           guardarPesoCorporal: guardarPesoCorporal_ };
}
/** Comprueba la clave (si está configurada) y ejecuta la función pedida. req = { fn, args, key } */
function despachar_(req) {
  var clave = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (clave && String(req.key || '').toUpperCase() !== clave) throw new Error('CLAVE');
  var f = apiFns_()[req.fn];
  if (!f) throw new Error('Función no permitida: ' + req.fn);
  return JSON.parse(JSON.stringify(f.apply(null, req.args || [])));
}
/** Llamadas desde la app servida por Apps Script (google.script.run). */
function rpc(req) { return despachar_(req); }
/** Llamadas desde la app instalada fuera de Google (fetch POST). */
function doPost(e) {
  var out;
  try { out = { ok: true, result: despachar_(JSON.parse(e.postData.contents)) }; }
  catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gym')
    .addItem('Instalar gráficos y copias semanales', 'instalar')
    .addItem('Hacer copia de seguridad ahora', 'backupSemanal')
    .addItem('Recalcular desde el Log', 'recalcularDesdeLog')
    .addSeparator()
    .addItem('Conectar el móvil (clave de la app)', 'conectarMovil')
    .addToUi();
}

// ------------------------------------------------------------------ lectura
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(n) { return ss_().getSheetByName(n); }
function hoyIso_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function iso_(d) { return Object.prototype.toString.call(d) === '[object Date]' ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d); }
/** Últimas `k` filas (o menos) de una hoja: devuelve { inicio, valores }. */
function cola_(sh, n, k, cols) { var ini = Math.max(2, n - k + 1); return n < 2 ? { inicio: 2, valores: [] } : { inicio: ini, valores: sh.getRange(ini, 1, n - ini + 1, cols).getValues() }; }

function config_() {
  var v = sh_('Config').getDataRange().getValues(), c = {};
  v.slice(1).forEach(function (r) { c[r[0]] = r[1]; });
  return {
    pesoCorporal: Number(c.peso_corporal_kg) || 75,
    repsMin: Number(c.reps_min_defecto) || 8,
    repsMax: Number(c.reps_max_defecto) || 12,
    incremento: Number(c.incremento_defecto_kg) || 2.5,
    rotacion: String(c.rotacion || 'Push, Pull').split(',').map(function (s) { return s.trim(); })
  };
}

function ejercicios_() {
  var sh = sh_('Ejercicios'), n = sh.getLastRow();
  var v = sh.getRange(1, 1, n, COL_NOTA_EJ).getValues();
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return { id: r[0], nombre: r[1], grupo: r[2], grupo2: r[3], equipo: r[4], tipoCarga: r[5],
             repsMin: Number(r[6]) || 8, repsMax: Number(r[7]) || 12, ultima: r[9] ? iso_(r[9]) : '',
             mejor: r[10] === '' ? null : Number(r[10]), activo: activo_(r[9] || null), nota: String(r[COL_NOTA_EJ - 1] || '') };
  });
}

function rutinas_() {
  var v = sh_('Rutinas').getDataRange().getValues(), out = {};
  v.slice(1).forEach(function (r) {
    if (!r[0]) return;
    (out[r[0]] = out[r[0]] || []).push({ ej: r[2], series: Number(r[5]) || 3 });
  });
  return out;
}

/** Filas de Log de los últimos DIAS_HISTORIAL días, no borradas. */
function logReciente_() {
  var sh = sh_('Log'), n = sh.getLastRow();
  if (n < 2) return [];
  var v = cola_(sh, n, FILAS_RECIENTES, LOG.client).valores;
  var lim = new Date(); lim.setDate(lim.getDate() - DIAS_HISTORIAL);
  return v.filter(function (r) { return r[LOG.borrado - 1] !== true && r[LOG.fecha - 1] >= lim; }).map(function (r) {
    return { id: r[0], sesion: r[1], fecha: iso_(r[2]), orden: r[4], ej: r[5], variante: r[7], serie: r[8], tramo: r[9],
             tipo: r[10], peso: r[11] === '' ? null : Number(r[11]), reps: r[12] === '' ? null : Number(r[12]), notas: r[15] };
  });
}

/** Historial por ejercicio en el formato de Pred.gs (sin la sesión de hoy). */
function historiales_(filas, hoy) {
  var h = {};
  filas.forEach(function (x) {
    if (x.fecha === hoy || x.tramo !== 1 || x.tipo === 'calentamiento') return;
    var e = (h[x.ej] = h[x.ej] || {}); (e[x.fecha] = e[x.fecha] || []).push({ serie: x.serie, peso: x.peso, reps: x.reps });
  });
  var out = {};
  Object.keys(h).forEach(function (ej) {
    out[ej] = Object.keys(h[ej]).sort().map(function (f) {
      return { fecha: f, sets: h[ej][f].sort(function (a, b) { return a.serie - b.serie; }) };
    });
  });
  return out;
}

/** Estado completo para la app: catálogo, rutinas, sesiones de hoy, historial reciente y modelo de descanso. */
function getEstado_() {
  var hoy = hoyIso_(), cfg = config_(), ejs = ejercicios_(), filas = logReciente_();
  var hist = historiales_(filas, hoy);
  var hoySets = filas.filter(function (x) { return x.fecha === hoy; });
  var shS = sh_('Sesiones'), nS = shS.getLastRow();
  var ses = cola_(shS, nS, 60, COL_SES_DESC).valores.filter(function (r) { return r[0]; });
  var ultimas = ses.map(function (r) {
    return { id: r[0], fecha: iso_(r[1]), tipo: r[3], volumen: Number(r[6]) || 0, duracion: Number(r[COL_SES_DUR - 1]) || null };
  });
  var deHoy = ultimas.filter(function (s) { return s.fecha === hoy && s.tipo !== 'Descartada'; });
  var resumenHist = {};
  Object.keys(hist).forEach(function (ej) { resumenHist[ej] = hist[ej].slice(-6); });
  return { hoy: hoy, cfg: cfg, ejercicios: ejs, rutinas: rutinas_(), historial: resumenHist,
           hoySets: hoySets, sesionesHoy: deHoy, sesionHoy: deHoy.length ? deHoy[deHoy.length - 1] : null,
           ultimasSesiones: ultimas, siguiente: siguienteRutina_(ultimas, cfg.rotacion),
           historialSesiones: agruparSesiones_(filas, 10), modeloDescanso: modeloDescanso_(filas) };
}

/**
 * Cuánto cambian las reps entre dos series seguidas con el mismo peso según el descanso entre ellas
 * (aprendido de las notas "descanso m:ss" que guarda la app). Tramos: <90 s, 90–150, 150–210, >210.
 */
function modeloDescanso_(filas) {
  var LIM = [90, 150, 210], tramos = [[], [], [], []], todos = [];
  var prev = {};
  filas.forEach(function (x) {
    if (x.tramo !== 1 || x.tipo === 'calentamiento') return;
    var k = x.sesion + '|' + x.ej, p = prev[k]; prev[k] = x;
    var m = /descanso (\d+):(\d\d)/.exec(x.notas || ''); if (!p || !m || p.peso !== x.peso || x.reps == null || p.reps == null) return;
    var s = Number(m[1]) * 60 + Number(m[2]); if (s < 20 || s > 600) return;
    var t = s < LIM[0] ? 0 : s < LIM[1] ? 1 : s < LIM[2] ? 2 : 3;
    tramos[t].push(x.reps - p.reps); todos.push(s);
  });
  todos.sort(function (a, b) { return a - b; });
  return { limites: LIM, tipico: todos.length ? todos[Math.floor(todos.length / 2)] : null,
           tramos: tramos.map(function (d) { return { n: d.length, media: d.length ? d.reduce(function (a, b) { return a + b; }, 0) / d.length : 0 }; }) };
}

function siguienteRutina_(ultimas, rot) {
  for (var i = ultimas.length - 1; i >= 0; i--) {
    var k = rot.indexOf(ultimas[i].tipo);
    if (k >= 0) return rot[(k + 1) % rot.length];
  }
  return rot[0];
}

// ------------------------------------------------------------------ escritura
/**
 * Guarda una serie (o un tramo de dropset/rest-pause). Idempotente por clientId (reintentos sin cobertura).
 * s = { clientId, sesion?, fecha?, ej, variante, peso, reps, tramo, tipo, rutina, notas }
 * sesion: 'S2026-11-01' para la primera sesión del día, 'S2026-11-01-2' para la segunda, etc.
 */
function guardarSerie_(s) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sh_('Log'), n = sh.getLastRow();
    var fecha = s.fecha || hoyIso_();
    var sid = /^S\d{4}-\d{2}-\d{2}(-\d+)?$/.test(s.sesion || '') ? s.sesion : 'S' + fecha;
    var bloque = cola_(sh, n, 300, LOG.client), ultimas = bloque.valores;
    for (var i = ultimas.length - 1; i >= 0; i--) {   // idempotencia: solo hace falta mirar las últimas filas
      if (s.clientId && ultimas[i][LOG.client - 1] === s.clientId) return { ok: true, duplicado: true, id: ultimas[i][0] };
    }
    var deSes = ultimas.filter(function (r) { return r[LOG.sesion - 1] === sid && r[LOG.borrado - 1] !== true; });
    var delEj = deSes.filter(function (r) { return r[LOG.ej - 1] === s.ej; });
    var tramo = Number(s.tramo) || 1;
    var serie = tramo > 1 && delEj.length ? Math.max.apply(null, delEj.map(function (r) { return r[LOG.serie - 1]; }))
                                          : delEj.filter(function (r) { return r[LOG.tramo - 1] === 1; }).length + 1;
    var orden;
    if (delEj.length) orden = delEj[0][LOG.orden - 1];
    else { var vistos = {}; deSes.forEach(function (r) { vistos[r[LOG.ej - 1]] = 1; }); orden = Object.keys(vistos).length + 1; }
    var ej = ejercicios_().filter(function (e) { return e.id === s.ej; })[0];
    if (!ej) throw new Error('Ejercicio desconocido: ' + s.ej);
    var peso = s.peso === null || s.peso === '' || s.peso === undefined ? '' : Number(s.peso);
    var carga = cargaEfectiva_(ej.tipoCarga, peso, config_().pesoCorporal);
    var tipo = s.tipo || (tramo > 1 ? 'drop' : 'normal');
    if (tramo > 1) {   // el 1er tramo de la serie pasa a ser drop/rest-pause también
      delEj.forEach(function (r) {
        if (r[LOG.serie - 1] === serie && r[LOG.tramo - 1] === 1 && r[LOG.tipo - 1] === 'normal') sh.getRange(bloque.inicio + ultimas.indexOf(r), LOG.tipo).setValue(tipo);
      });
    }
    var id = 'L' + Utilities.formatString('%05d', n) + '-' + String(s.clientId || Date.now()).slice(-4);
    var fila = [id, sid, fechaDate_(fecha), false, orden, s.ej, ej.nombre, s.variante || '', serie, tramo, tipo, peso, Number(s.reps), carga,
                '', s.notas || '', 'app', '', false, 0, s.clientId || ''];
    fila[LOG.e1rm - 1] = e1rmFila_(fila); fila[LOG.volumen - 1] = volFila_(fila);
    sh.getRange(n + 1, 1, 1, fila.length).setValues([fila]);
    asegurarSesion_(sid, fecha, s.rutina);
    actualizarDerivados_(sid, [s.ej]);
    return { ok: true, id: id, serie: serie, tramo: tramo, orden: orden, sesion: sid, e1rm: fila[LOG.e1rm - 1] || null };
  } finally { lock.releaseLock(); }
}

function cargaEfectiva_(tipoCarga, peso, bw) {
  if (tipoCarga === 'lastre') return bw + (Number(peso) || 0);
  if (tipoCarga === 'asistencia') return bw - (Number(peso) || 0);
  if (tipoCarga === 'reps' || tipoCarga === 'segs') return '';
  return peso;
}
function fechaDate_(iso) { var p = iso.split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }

function asegurarSesion_(sid, fecha, rutina) {
  var sh = sh_('Sesiones'), n = sh.getLastRow();
  var b = cola_(sh, n, 30, 4), k = -1;
  b.valores.forEach(function (r, i) { if (r[0] === sid) k = i; });
  if (k >= 0) {   // ya existe: si se había descartado, se reactiva con la rutina nueva
    if (b.valores[k][3] === 'Descartada') sh.getRange(b.inicio + k, 4).setValue(rutina || '');
    return;
  }
  sh.getRange(n + 1, 1, 1, 9).setValues([[sid, fechaDate_(fecha), false, rutina || '', 0, 0, 0, 'app', '']]);
}

/** Borrado lógico (Deshacer o desde Historial). */
function borrarSerie_(id) {
  var f = buscarFilaLog_(id); if (!f) return { ok: false, error: 'No encontrada' };
  var sh = sh_('Log'), r = sh.getRange(f, 1, 1, LOG.client).getValues()[0];
  r[LOG.borrado - 1] = true;
  sh.getRange(f, LOG.e1rm).setValue(''); sh.getRange(f, LOG.borrado).setValue(true); sh.getRange(f, LOG.volumen).setValue(0);
  actualizarDerivados_(r[LOG.sesion - 1], [r[LOG.ej - 1]]);
  return { ok: true };
}
/** Fila (1-based) de un id del Log, buscando primero en las últimas filas. */
function buscarFilaLog_(id) {
  var sh = sh_('Log'), n = sh.getLastRow(), b = cola_(sh, n, 2000, 1);
  for (var i = b.valores.length - 1; i >= 0; i--) if (b.valores[i][0] === id) return b.inicio + i;
  var v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var j = v.length - 1; j >= 0; j--) if (v[j][0] === id) return j + 2;
  return null;
}

/** Descarta (borrado lógico) todas las series de una sesión y la marca como descartada. Acepta id de sesión o fecha. */
function descartarSesion_(sesionOFecha) {
  var sid = /^S/.test(sesionOFecha) ? sesionOFecha : 'S' + sesionOFecha, sh = sh_('Log'), n = sh.getLastRow();
  var b = cola_(sh, n, 400, LOG.client), ejs = {};
  b.valores.forEach(function (r, i) {
    if (r[LOG.sesion - 1] !== sid || r[LOG.borrado - 1] === true) return;
    ejs[r[LOG.ej - 1]] = 1;
    sh.getRange(b.inicio + i, LOG.e1rm).setValue(''); sh.getRange(b.inicio + i, LOG.borrado).setValue(true); sh.getRange(b.inicio + i, LOG.volumen).setValue(0);
  });
  var ss = sh_('Sesiones'), v = cola_(ss, ss.getLastRow(), 30, 1);
  v.valores.forEach(function (r, i) { if (r[0] === sid) ss.getRange(v.inicio + i, 4).setValue('Descartada'); });
  if (Object.keys(ejs).length) actualizarDerivados_(sid, Object.keys(ejs));
  return { ok: true };
}

/** Corrige peso y reps de una serie ya guardada. */
function editarSerie_(id, peso, reps) {
  var f = buscarFilaLog_(id); if (!f) return { ok: false, error: 'No encontrada' };
  var sh = sh_('Log'), r = sh.getRange(f, 1, 1, LOG.client).getValues()[0];
  var ej = ejercicios_().filter(function (e) { return e.id === r[LOG.ej - 1]; })[0];
  var p = peso === '' || peso === null ? '' : Number(peso);
  r[LOG.peso - 1] = p; r[LOG.reps - 1] = Number(reps); r[LOG.carga - 1] = cargaEfectiva_(ej.tipoCarga, p, config_().pesoCorporal);
  sh.getRange(f, LOG.peso, 1, 4).setValues([[p, Number(reps), r[LOG.carga - 1], e1rmFila_(r)]]);
  sh.getRange(f, LOG.volumen).setValue(volFila_(r));
  actualizarDerivados_(r[LOG.sesion - 1], [r[LOG.ej - 1]]);
  return { ok: true };
}

/** Alta de un ejercicio nuevo desde la app. e = { nombre, grupo, grupo2, equipo, tipoCarga } */
function crearEjercicio_(e) {
  var sh = sh_('Ejercicios'), n = sh.getLastRow();
  var base = String(e.nombre).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 30);
  var existentes = sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return r[0]; });
  var id = base, k = 2;
  while (existentes.indexOf(id) >= 0) id = base + '_' + (k++);
  var r = n + 1;
  sh.getRange(r, 1, 1, 13).setValues([[id, e.nombre, e.grupo, e.grupo2 || '', e.equipo || '', e.tipoCarga || 'total', 8, 12,
    0, '', '', 'Sí', 'Creado desde la app']]);
  return { id: id, nombre: e.nombre, grupo: e.grupo, grupo2: e.grupo2 || '', equipo: e.equipo || '', tipoCarga: e.tipoCarga || 'total',
           repsMin: 8, repsMax: 12, ultima: '', mejor: null, activo: true, nota: '' };
}

/** Serie temporal de un ejercicio para la pantalla Progreso. */
function getProgreso_(ej) {
  var v = sh_('Resumen').getDataRange().getValues().slice(1);
  var pts = v.filter(function (r) { return r[2] === ej; }).map(function (r) {
    return { fecha: iso_(r[1]), e1rm: r[4] === '' ? null : Number(r[4]), top: r[5] === '' ? null : Number(r[5]), vol: Number(r[7]) || 0 };
  });
  return { ej: ej, puntos: pts };
}

/** Últimas sesiones con sus series, para Historial. */
function getHistorial_(nSesiones) { return agruparSesiones_(logReciente_(), nSesiones || 10); }
function agruparSesiones_(filas, nSesiones) {
  var por = {};
  filas.forEach(function (x) { (por[x.sesion] = por[x.sesion] || []).push(x); });
  var ids = Object.keys(por).sort().slice(-nSesiones).reverse();
  return ids.map(function (sid) { return { sesion: sid, fecha: por[sid][0].fecha, sets: por[sid] }; });
}

/** Guarda la nota fijada de un ejercicio (Ejercicios!N). */
function guardarNotaEjercicio_(id, texto) {
  var sh = sh_('Ejercicios'), n = sh.getLastRow();
  if (!sh.getRange(1, COL_NOTA_EJ).getValue()) sh.getRange(1, COL_NOTA_EJ).setValue('nota_fija');
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) { sh.getRange(i + 2, COL_NOTA_EJ).setValue(String(texto || '').slice(0, 300)); return { ok: true }; }
  }
  throw new Error('Ejercicio desconocido: ' + id);
}

/** Al terminar: guarda duración (min) y descanso medio (s) en la fila de la sesión. Acepta id de sesión o fecha. */
function cerrarSesion_(sesionOFecha, datos) {
  var sh = sh_('Sesiones'), n = sh.getLastRow(), sid = /^S/.test(sesionOFecha) ? sesionOFecha : 'S' + sesionOFecha;
  if (!sh.getRange(1, COL_SES_DUR).getValue()) sh.getRange(1, COL_SES_DUR, 1, 2).setValues([['duracion_min', 'descanso_medio_s']]);
  var b = cola_(sh, n, 30, 1);
  for (var i = b.valores.length - 1; i >= 0; i--) {
    if (b.valores[i][0] === sid) {
      sh.getRange(b.inicio + i, COL_SES_DUR, 1, 2).setValues([[datos.duracionMin || '', datos.descansoMedio || '']]);
      return { ok: true };
    }
  }
  return { ok: false };
}

/** Ajustes → peso corporal: actualiza Config y guarda el cambio con fecha en la hoja PesoCorporal. */
function guardarPesoCorporal_(kg) {
  kg = Number(kg); if (!(kg >= 30 && kg <= 250)) throw new Error('Peso no válido');
  var c = sh_('Config'), v = c.getRange(1, 1, c.getLastRow(), 1).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === 'peso_corporal_kg') c.getRange(i + 1, 2).setValue(kg);
  var h = sh_('PesoCorporal') || ss_().insertSheet('PesoCorporal');
  if (h.getLastRow() === 0) h.appendRow(['fecha', 'peso_kg']);
  h.appendRow([fechaDate_(hoyIso_()), kg]);
  return { ok: true, pesoCorporal: kg };
}

/**
 * Ejercicios activos (usados en los últimos 45 días, con 5+ sesiones) cuyo e1RM no supera su mejor marca
 * en las últimas `ventana` sesiones.
 */
function getEstancados_(ventana) {
  ventana = ventana || 4;
  var v = sh_('Resumen').getDataRange().getValues().slice(1), por = {};
  v.forEach(function (r) { if (r[2] && r[4] !== '') (por[r[2]] = por[r[2]] || []).push({ fecha: iso_(r[1]), e1rm: Number(r[4]) }); });
  var lim = Utilities.formatDate(new Date(Date.now() - 45 * 864e5), TZ, 'yyyy-MM-dd'), out = [];
  Object.keys(por).forEach(function (ej) {
    var p = por[ej]; if (p.length < ventana + 1 || p[p.length - 1].fecha < lim) return;
    var antes = p.slice(0, -ventana), ult = p.slice(-ventana);
    var mejorAntes = Math.max.apply(null, antes.map(function (x) { return x.e1rm; }));
    var mejorUlt = Math.max.apply(null, ult.map(function (x) { return x.e1rm; }));
    if (mejorUlt <= mejorAntes) {
      var iRec = -1; antes.forEach(function (x, i) { if (x.e1rm === mejorAntes) iRec = i; });
      out.push({ ej: ej, sesiones: p.length - 1 - iRec, mejor: mejorAntes, fechaMejor: antes[iRec].fecha, ultimo: p[p.length - 1].e1rm });
    }
  });
  return out.sort(function (a, b) { return b.sesiones - a.sesiones; });
}

// ------------------------------------------------------------------ instalación, gráficos y copias
function soloPropietario_() {
  if (!Session.getActiveUser().getEmail()) throw new Error('Solo el propietario puede hacer esto desde la hoja.');
}

function instalar() {
  soloPropietario_();
  crearGraficos_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupSemanal') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  SpreadsheetApp.getActive().toast('Gráficos creados y copia semanal programada (lunes 4:00).', 'Gym', 8);
}

function crearGraficos_() {
  var d = sh_('Dashboard'), s = sh_('Semanal');
  d.getCharts().forEach(function (c) { d.removeChart(c); });
  s.getCharts().forEach(function (c) { s.removeChart(c); });
  d.insertChart(d.newChart().asLineChart()
    .addRange(d.getRange('C9:C209')).addRange(d.getRange('D9:E209'))
    .setNumHeaders(1).setOption('title', 'Fuerza estimada (e1RM, kg)').setOption('legend', { position: 'bottom' })
    .setOption('series', { 1: { lineDashStyle: [4, 4] } }).setOption('interpolateNulls', true)
    .setPosition(3, 10, 0, 0).setOption('width', 720).setOption('height', 320).build());
  d.insertChart(d.newChart().asLineChart()
    .addRange(d.getRange('C9:C209')).addRange(d.getRange('F9:F209'))
    .setNumHeaders(1).setOption('title', 'Peso de la serie más pesada (kg)').setOption('legend', { position: 'none' })
    .setOption('interpolateNulls', true).setPosition(21, 10, 0, 0).setOption('width', 720).setOption('height', 280).build());
  s.insertChart(s.newChart().asColumnChart()
    .addRange(s.getRange('A4:A30')).addRange(s.getRange('C4:N30'))
    .setNumHeaders(1).setStacked().setOption('title', 'Volumen semanal por grupo (kg × reps)')
    .setOption('legend', { position: 'right' }).setPosition(4, 16, 0, 0).setOption('width', 900).setOption('height', 380).build());
}

function backupSemanal() {
  var carpeta = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('BACKUP_FOLDER_ID') || '1uFQeefKCFzmB5E_T2lwJwlNWAl5D0sDj');
  var f = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  ['Log', 'Sesiones', 'Ejercicios'].forEach(function (n) {
    var v = sh_(n).getDataRange().getDisplayValues();
    var csv = v.map(function (r) { return r.map(function (c) { c = String(c); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(','); }).join('\n');
    carpeta.createFile('Gym v2 - ' + n + ' - ' + f + '.csv', csv, MimeType.CSV);
  });
}

/** Menú Gym → Conectar el móvil: crea (una vez) la clave de la app y la muestra con la URL de la API. */
function conectarMovil() {
  soloPropietario_();
  var props = PropertiesService.getScriptProperties(), clave = props.getProperty('API_KEY');
  if (!clave) {
    var abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now()), s = '';
    for (var i = 0; i < 16; i++) s += abc.charAt((raw[i] + 256) % 32) + ((i % 4 === 3 && i < 15) ? '-' : '');
    clave = s; props.setProperty('API_KEY', clave);
  }
  var url = ScriptApp.getService().getUrl() || '(implementa primero la aplicación web)';
  var html = '<div style="font:14px Arial;line-height:1.5">' +
    '<p>Clave de la app (escríbela una vez en el móvil):</p>' +
    '<p style="font:700 22px monospace;letter-spacing:1px;background:#f3f4f6;padding:10px;border-radius:8px;text-align:center">' + clave + '</p>' +
    '<p>URL de la API (va en <code>config.js</code>):</p><p style="word-break:break-all;background:#f3f4f6;padding:8px;border-radius:8px">' + url + '</p>' +
    '<p style="color:#666">La URL correcta termina en <b>/exec</b>. Si la clave se filtra, borra la propiedad API_KEY en Configuración del proyecto y vuelve a abrir este menú para crear otra.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(360), 'Conectar el móvil');
}

/** Inyecta en la página el mismo motor de predicción que usa el servidor (Pred.gs). */
function predSource() {
  return [sugerirSerie, sugerirProgresion, ajusteDescanso, e1rm, ultimaSesion, caidaTipica, pesosUsados, esNum, limitar, redondear, fmt, fechaCorta]
    .map(String).join('\n');
}
