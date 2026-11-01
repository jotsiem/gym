/**
 * Motor de predicción (funciones puras, sin dependencias de Google).
 * Validado contra el histórico (5.153 series): peso exacto en el 76,5 % y peso exacto con ±1 rep en el 64,7 %.
 *
 * Una "sesión" de un ejercicio es: { fecha: 'yyyy-mm-dd', sets: [{ serie, peso, reps }] }
 * con solo las series efectivas (1er tramo, sin calentamiento), ordenadas por serie.
 */

/** Sugerencia para la serie número `k` (1, 2, 3...) de hoy. */
function sugerirSerie(historial, hoy, k, cfg) {
  cfg = cfg || {};
  var repsMax = cfg.repsMax || 12, repsMin = cfg.repsMin || 8;
  var L = ultimaSesion(historial);
  var prevHoy = hoy.length ? hoy[hoy.length - 1] : null;
  if (!L) {
    // Ejercicio sin historial: repetir lo de hoy o dejar vacío.
    return prevHoy ? { peso: prevHoy.peso, reps: prevHoy.reps, motivo: 'Igual que la serie anterior' }
                   : { peso: null, reps: null, motivo: 'Sin historial: introduce peso y reps' };
  }
  var Lk = L.sets[k - 1], Lprev = L.sets[k - 2];
  if (k === 1 || !prevHoy) {
    if (!Lk) Lk = L.sets[0];
    return { peso: Lk.peso, reps: Lk.reps, motivo: 'Como el ' + fechaCorta(L.fecha) };
  }
  if (!Lk) {
    // Hoy haces más series que la última vez: misma carga y caída típica de reps.
    var d = caidaTipica(historial);
    return { peso: prevHoy.peso, reps: limitar(prevHoy.reps + d, 1, repsMax), motivo: 'Serie extra: caída típica de ' + d + ' rep' };
  }
  // Series 2+: mismo peso que la serie anterior de hoy (acierta el 76,5 % frente al 72,5 % de copiar la última sesión).
  var peso = esNum(prevHoy.peso) ? prevHoy.peso : Lk.peso;
  var reps = Lk.reps;
  if (Lprev && prevHoy.peso === Lprev.peso && esNum(prevHoy.reps) && esNum(Lprev.reps)) {
    reps = Lk.reps + (prevHoy.reps - Lprev.reps);   // hoy vas mejor/peor que la última vez
  }
  return { peso: peso, reps: limitar(reps, 1, Math.max(repsMax, Lk.reps)), motivo: 'Como el ' + fechaCorta(L.fecha) + ', ajustado a tu serie anterior' };
}

/** Propuesta de progresión (doble progresión) para mostrar como botón aparte. */
function sugerirProgresion(historial, cfg) {
  cfg = cfg || {};
  var repsMax = cfg.repsMax || 12, incremento = cfg.incremento || 2.5;
  var L = ultimaSesion(historial);
  if (!L || !L.sets.length || !esNum(L.sets[0].peso)) return null;
  var W = L.sets[0].peso;
  var aW = L.sets.filter(function (s) { return s.peso === W; });
  if (!aW.length || !aW.every(function (s) { return s.reps >= repsMax; })) return null;
  var usados = pesosUsados(historial).filter(function (p) { return p > W; });
  var W2 = usados.length ? usados[0] : redondear(W + incremento);
  var e = W * (1 + aW[0].reps / 30);
  var r = limitar(Math.round(30 * (e / W2 - 1)), 1, repsMax);
  return { peso: W2, reps: r, motivo: 'Llegaste a ' + repsMax + ' en todas las series con ' + fmt(W) + ' kg' };
}

/**
 * Ajuste de reps por el descanso actual frente a tu descanso típico, con el modelo aprendido en el servidor
 * (modeloDescanso_). Devuelve 0 mientras no haya al menos 15 muestras en los dos tramos comparados.
 */
function ajusteDescanso(m, seg) {
  if (!m || !m.tipico || seg === null || seg === undefined) return 0;
  var t = function (s) { return s < m.limites[0] ? 0 : s < m.limites[1] ? 1 : s < m.limites[2] ? 2 : 3; };
  var a = m.tramos[t(seg)], b = m.tramos[t(m.tipico)];
  if (!a || !b || a.n < 15 || b.n < 15) return 0;
  return Math.max(-2, Math.min(1, Math.round(a.media - b.media)));
}

/** e1RM de Epley. */
function e1rm(carga, reps) {
  if (!esNum(carga) || !esNum(reps) || carga <= 0 || reps > 15) return null;
  return Math.round(carga * (1 + reps / 30) * 10 + 1e-9) / 10;
}

function ultimaSesion(historial) {
  for (var i = historial.length - 1; i >= 0; i--) if (historial[i].sets && historial[i].sets.length) return historial[i];
  return null;
}
function caidaTipica(historial) {
  var d = [];
  historial.slice(-10).forEach(function (s) {
    for (var i = 1; i < s.sets.length; i++) if (s.sets[i].peso === s.sets[i - 1].peso) d.push(s.sets[i].reps - s.sets[i - 1].reps);
  });
  if (!d.length) return 0;
  d.sort(function (a, b) { return a - b; });
  return d[Math.floor(d.length / 2)];
}
function pesosUsados(historial) {
  var m = {};
  historial.forEach(function (s) { s.sets.forEach(function (x) { if (esNum(x.peso)) m[x.peso] = 1; }); });
  return Object.keys(m).map(Number).sort(function (a, b) { return a - b; });
}
function esNum(x) { return typeof x === 'number' && !isNaN(x); }
function limitar(x, lo, hi) { return Math.max(lo, Math.min(hi, Math.round(x))); }
function redondear(x) { return Math.round(x * 4) / 4; }
function fmt(x) { return String(x).replace('.', ','); }
function fechaCorta(iso) { var p = String(iso).split('-'); return p[2] + '/' + p[1]; }

if (typeof module !== 'undefined') module.exports = { sugerirSerie: sugerirSerie, sugerirProgresion: sugerirProgresion, e1rm: e1rm, ajusteDescanso: ajusteDescanso };
