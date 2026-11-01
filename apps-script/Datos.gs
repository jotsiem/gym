/**
 * Datos derivados calculados por el script (sustituyen a las fórmulas MAXIFS/SUMIFS que recorrían el Log entero).
 *
 *  - Log:        G nombre, O e1rm y T volumen se escriben como valores al guardar.
 *  - Resumen:    una fila por sesión × ejercicio (A..H), se actualiza al guardar/editar/borrar.
 *  - Sesiones:   E ejercicios, F series, G volumen como valores.
 *  - Ejercicios: I series efectivas, J última vez, K mejor e1RM, L activo como valores.
 *
 * reconstruir_() lo recalcula todo desde el Log (menú Gym → Recalcular desde el Log): sirve como migración y
 * como red de seguridad si se edita el Log a mano.
 */

var RES_COLS = 8;   // Resumen A..H: sesion_id, fecha, ejercicio_id, grupo, mejor_e1rm, peso_top_kg, series_efectivas, volumen_kg

/** e1RM (Epley) de una fila del Log, o '' si no cuenta. */
function e1rmFila_(r) {
  var c = r[LOG.carga - 1], reps = r[LOG.reps - 1];
  if (r[LOG.borrado - 1] === true || r[LOG.tipo - 1] === 'calentamiento' || Number(r[LOG.tramo - 1]) !== 1) return '';
  if (typeof c !== 'number' || c <= 0 || typeof reps !== 'number' || reps > 15) return '';
  return Math.round(c * (1 + reps / 30) * 10 + 1e-9) / 10;   // +1e-9: redondea 29,75 a 29,8 como la hoja
}
/** Volumen (kg × reps) de una fila del Log. */
function volFila_(r) {
  var c = r[LOG.carga - 1], reps = r[LOG.reps - 1];
  if (r[LOG.borrado - 1] === true || r[LOG.tipo - 1] === 'calentamiento') return 0;
  return (typeof c === 'number' && typeof reps === 'number') ? c * reps : 0;
}
/** Agrega las filas del Log de una sesión y ejercicio en una fila de Resumen (o null si no queda nada vivo). */
function filaResumen_(filas, sid, ej, grupo) {
  var mejor = '', top = 0, series = 0, vol = 0, fecha = null, vivas = 0;
  filas.forEach(function (r) {
    if (r[LOG.sesion - 1] !== sid || r[LOG.ej - 1] !== ej) return;
    fecha = r[LOG.fecha - 1];
    if (r[LOG.borrado - 1] === true) return;
    vivas++;
    var e = e1rmFila_(r); if (e !== '' && (mejor === '' || e > mejor)) mejor = e;
    var p = r[LOG.peso - 1];
    if (r[LOG.tipo - 1] !== 'calentamiento' && typeof p === 'number' && p > top) top = p;
    if (Number(r[LOG.tramo - 1]) === 1 && r[LOG.tipo - 1] !== 'calentamiento') series++;
    vol += volFila_(r);
  });
  if (!vivas) return null;
  return [sid, fecha, ej, grupo || '', mejor, top, series, vol];
}
function grupos_() { var m = {}; ejercicios_().forEach(function (e) { m[e.id] = e.grupo; }); return m; }

/**
 * Tras guardar/editar/borrar en la sesión `sid`: recalcula la fila de Resumen de cada ejercicio tocado,
 * la fila de Sesiones y las estadísticas de esos ejercicios en Ejercicios.
 */
function actualizarDerivados_(sid, ejs) {
  var log = sh_('Log'), n = log.getLastRow();
  var filas = cola_(log, n, 600, LOG.client).valores;
  var g = grupos_();
  // Resumen
  var res = sh_('Resumen'), nr = res.getLastRow(), cr = cola_(res, nr, 200, RES_COLS);
  ejs.forEach(function (ej) {
    var nueva = filaResumen_(filas, sid, ej, g[ej]), k = -1;
    cr.valores.forEach(function (r, i) { if (r[0] === sid && r[2] === ej) k = i; });
    if (k >= 0) {
      if (nueva) { res.getRange(cr.inicio + k, 1, 1, RES_COLS).setValues([nueva]); cr.valores[k] = nueva; }
      else { res.deleteRow(cr.inicio + k); cr = cola_(res, res.getLastRow(), 200, RES_COLS); }
    } else if (nueva) {
      res.getRange(res.getLastRow() + 1, 1, 1, RES_COLS).setValues([nueva]);
      cr = cola_(res, res.getLastRow(), 200, RES_COLS);
    }
  });
  // Sesiones
  var tot = [0, 0, 0];
  cr.valores.forEach(function (r) { if (r[0] === sid) { tot[0]++; tot[1] += Number(r[6]) || 0; tot[2] += Number(r[7]) || 0; } });
  var ses = sh_('Sesiones'), cs = cola_(ses, ses.getLastRow(), 30, 1);
  cs.valores.forEach(function (r, i) { if (r[0] === sid) ses.getRange(cs.inicio + i, 5, 1, 3).setValues([[tot[0], tot[1], Math.round(tot[2])]]); });
  // Ejercicios
  actualizarEstadisticasEj_(ejs);
}

/** I series efectivas, J última vez, K mejor e1RM, L activo de los ejercicios indicados (desde Resumen). */
function actualizarEstadisticasEj_(ejs) {
  var res = sh_('Resumen'), nr = res.getLastRow();
  var v = nr > 1 ? res.getRange(2, 1, nr - 1, RES_COLS).getValues() : [];
  var st = {};
  ejs.forEach(function (id) { st[id] = { series: 0, ultima: null, mejor: '' }; });
  v.forEach(function (r) {
    var s = st[r[2]]; if (!s) return;
    s.series += Number(r[6]) || 0;
    if (!s.ultima || r[1] > s.ultima) s.ultima = r[1];
    if (r[4] !== '' && (s.mejor === '' || r[4] > s.mejor)) s.mejor = r[4];
  });
  var sh = sh_('Ejercicios'), n = sh.getLastRow(), ids = sh.getRange(2, 1, n - 1, 1).getValues();
  ids.forEach(function (r, i) {
    var s = st[r[0]]; if (!s) return;
    sh.getRange(i + 2, 9, 1, 4).setValues([[s.series, s.ultima || '', s.mejor, activo_(s.ultima) ? 'Sí' : 'No']]);
  });
}
function activo_(fecha) { return !!fecha && (Date.now() - new Date(fecha).getTime()) <= 180 * 864e5; }

// ------------------------------------------------------------------ reconstrucción completa / migración
/** Menú Gym → Recalcular desde el Log. */
function recalcularDesdeLog() {
  soloPropietario_();
  var t = Date.now(), r = reconstruir_();
  SpreadsheetApp.getActive().toast('Recalculado en ' + Math.round((Date.now() - t) / 1000) + ' s: ' + r.log + ' filas de Log, ' +
    r.resumen + ' de Resumen, ' + r.sesiones + ' sesiones.', 'Gym', 10);
}

function reconstruir_() {
  var bw = config_().pesoCorporal, ejs = ejercicios_(), info = {};
  ejs.forEach(function (e) { info[e.id] = e; });
  // 1) Log: nombre, carga (si falta), e1rm y volumen como valores
  var log = sh_('Log'), n = log.getLastRow();
  var v = log.getRange(2, 1, n - 1, LOG.client).getValues();
  var colG = [], colN = [], colO = [], colT = [];
  v.forEach(function (r) {
    var e = info[r[LOG.ej - 1]] || {};
    if (r[LOG.carga - 1] === '' && r[LOG.peso - 1] !== '' && e.tipoCarga) r[LOG.carga - 1] = cargaEfectiva_(e.tipoCarga, r[LOG.peso - 1], bw);
    colG.push([e.nombre || '']); colN.push([r[LOG.carga - 1]]); colO.push([e1rmFila_(r)]); colT.push([volFila_(r)]);
  });
  log.getRange(2, LOG.nombre, n - 1, 1).setValues(colG);
  log.getRange(2, LOG.carga, n - 1, 1).setValues(colN);
  log.getRange(2, LOG.e1rm, n - 1, 1).setValues(colO);
  log.getRange(2, LOG.volumen, n - 1, 1).setValues(colT);
  // 2) Resumen: una fila por sesión × ejercicio, en orden de fecha y de ejercicio
  var orden = [], vistos = {};
  v.map(function (r, i) { return { r: r, i: i }; })
    .sort(function (a, b) { return (a.r[LOG.fecha - 1] - b.r[LOG.fecha - 1]) || String(a.r[1]).localeCompare(String(b.r[1])) || ((a.r[4] || 99) - (b.r[4] || 99)) || (a.i - b.i); })
    .forEach(function (x) { var k = x.r[1] + '|' + x.r[5]; if (!vistos[k]) { vistos[k] = 1; orden.push([x.r[1], x.r[5]]); } });
  var porSes = {};
  v.forEach(function (r) { (porSes[r[1]] = porSes[r[1]] || []).push(r); });
  var filasRes = [];
  orden.forEach(function (k) { var f = filaResumen_(porSes[k[0]], k[0], k[1], (info[k[1]] || {}).grupo); if (f) filasRes.push(f); });
  var res = sh_('Resumen');
  res.getRange(1, 1, 1, RES_COLS + 1).setValues([['sesion_id', 'fecha', 'ejercicio_id', 'grupo', 'mejor_e1rm', 'peso_top_kg', 'series_efectivas', 'volumen_kg', '']]);
  if (res.getLastRow() > 1) res.getRange(2, 1, res.getLastRow() - 1, Math.max(RES_COLS + 1, res.getLastColumn())).clearContent();
  if (filasRes.length) res.getRange(2, 1, filasRes.length, RES_COLS).setValues(filasRes);
  res.getRange(2, 2, Math.max(1, filasRes.length), 1).setNumberFormat('dd/mm/yyyy');
  // 3) Sesiones: E, F, G como valores (y filas que falten)
  var tot = {};
  filasRes.forEach(function (r) { var t = tot[r[0]] = tot[r[0]] || [0, 0, 0, r[1]]; t[0]++; t[1] += r[6]; t[2] += r[7]; });
  var ses = sh_('Sesiones'), ns = ses.getLastRow(), sv = ns > 1 ? ses.getRange(2, 1, ns - 1, 7).getValues() : [], hay = {};
  var efg = sv.map(function (r) { hay[r[0]] = 1; var t = tot[r[0]] || [0, 0, 0]; return [t[0], t[1], Math.round(t[2])]; });
  if (efg.length) ses.getRange(2, 5, efg.length, 3).setValues(efg);
  Object.keys(tot).filter(function (s) { return !hay[s]; }).sort().forEach(function (s) {
    var t = tot[s]; ses.appendRow([s, t[3], false, '', t[0], t[1], Math.round(t[2]), 'recalculado', '']);
  });
  // 4) Ejercicios: I..L como valores
  actualizarEstadisticasEj_(ejs.map(function (e) { return e.id; }));
  // 5) Dashboard: la tabla del ejercicio elegido sale con FILTER (sin la columna auxiliar de Resumen)
  var d = sh_('Dashboard');
  if (d) {
    d.getRange('C10:D209').clearContent(); d.getRange('F10:H209').clearContent();
    d.getRange('C10').setFormula('=IFERROR(FILTER({Resumen!B2:B,Resumen!E2:E},Resumen!C2:C=$C$4),"")');
    d.getRange('F10').setFormula('=IFERROR(FILTER(Resumen!F2:H,Resumen!C2:C=$C$4),"")');
  }
  return { log: n - 1, resumen: filasRes.length, sesiones: Object.keys(tot).length };
}
