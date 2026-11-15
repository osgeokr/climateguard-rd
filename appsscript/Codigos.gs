/**
 * ClimateGuard - Codigos de practica y cupos de envio (servidor v3.11)
 * ===========================================================================
 * Archivo NUEVO del mismo proyecto (junto a Code.gs, Sesion.gs y Amenazas.gs).
 *
 * Que hace
 *  - Un administrador crea un "codigo de practica" para cada capacitacion. Con
 *    ese codigo (escrito o leido de una QR) cualquier participante entra sin
 *    cuenta de Google: el servidor le crea una cuenta de practica
 *    (c3-007@practica.cg, "Participante 07") y le entrega la misma clave de
 *    sesion que usa el resto de la app (Sesion.gs).
 *  - El codigo NO limita el numero de participantes. La proteccion de Drive
 *    viene de tres cupos de registros nuevos (las ediciones no cuentan):
 *        lim_dia     por cuenta y dia
 *        lim_cuenta  por cuenta, acumulado
 *        tope_total  suma de todas las cuentas del codigo
 *  - El codigo deja de admitir cuentas NUEVAS al pasar la fecha "fin" o si el
 *    administrador lo pausa. Las cuentas ya creadas siguen entrando y enviando
 *    (dentro de los cupos); nada se borra solo.
 *  - Las cuentas de Google siguen con registro abierto y reciben solo el
 *    cupo diario (GOOGLE_LIM_DIA). No tienen cupo acumulado (GOOGLE_LIM_CUENTA
 *    = 0), para no bloquear a quienes ya enviaron muchos registros. Los
 *    administradores (ADMIN_EMAILS) no tienen cupo.
 *
 * Pestana "Codigos" (se crea sola). Se puede editar a mano: fechas, cupos y
 * estado se leen en cada peticion.
 *
 * Acciones GET (JSONP), todas llegan por sesionGet():
 *  - action=codigo&codigo=XXXX-XXXX&dev=...   entrar con un codigo (sin login)
 *  - action=codigos                           lista con uso      (solo admin)
 *  - action=codigocrear&nombre=..&dias=..     crear              (solo admin)
 *  - action=codigoestado&id=c3&estado=..      activo / pausado   (solo admin)
 * Envio (POST): sesionPost() llama a codigosCupo(body) antes de guardar.
 *
 * Instalacion
 *  1. Archivo nuevo "Codigos" en el editor de Apps Script: pegar este codigo.
 *  2. Reemplazar Sesion.gs por la version v3.11 (dos lineas nuevas marcadas).
 *  3. Implementar > Gestionar implementaciones > Editar > Nueva version.
 *  Code.gs no cambia.
 */

var HEAD_COD = ['id', 'codigo', 'nombre', 'inicio', 'fin', 'lim_dia', 'lim_cuenta', 'tope_total',
                'estado', 'cuentas', 'creado', 'creado_por'];
var COD_DOM = '@practica.cg';
/* Valores por defecto de un codigo nuevo (0 = sin limite) */
var COD_DEF = { dias: 15, lim_dia: 30, lim_cuenta: 100, tope_total: 2000 };
/* Cupos de las cuentas de Google (0 = sin limite). No aplican a ADMIN_EMAILS. */
var GOOGLE_LIM_DIA = 30;
var GOOGLE_LIM_CUENTA = 0;
/* Sin I, O, 0 ni 1, para que el codigo se pueda dictar y escribir */
var COD_ALFA = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function _codSheet(ss) { return _sheet(ss, 'Codigos', HEAD_COD); }
function _codNorm(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function _codNum(v) { var n = Number(v); return (isNaN(n) || n < 0) ? 0 : Math.floor(n); }
function _codDia(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function _codHoy() { return _codDia(new Date()); }
function _codPad(n) { return n < 10 ? '00' + n : (n < 100 ? '0' + n : String(n)); }
function _codBonito(c) { c = _codNorm(c); return c.length === 8 ? c.slice(0, 4) + '-' + c.slice(4) : c; }

/* 'c3' si el email es de una cuenta de practica; '' si no. */
function _codIdDeEmail(email) {
  var m = /^(c\d+)-\d+@practica\.cg$/.exec(String(email || '').toLowerCase());
  return m ? m[1] : '';
}

function _codFilas(sh) {
  var last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, HEAD_COD.length).getValues().forEach(function (v, i) {
    if (!String(v[0] || '')) return;
    var o = { fila: i + 2 };
    HEAD_COD.forEach(function (h, k) { o[h] = v[k]; });
    o.id = String(o.id).toLowerCase();
    out.push(o);
  });
  return out;
}

function _codNuevoTexto() {
  for (var t = 0; t < 20; t++) {
    var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + Utilities.getUuid(), Utilities.Charset.UTF_8);
    var s = '';
    for (var i = 0; i < 8; i++) s += COD_ALFA.charAt(((raw[i] + 256) % 256) % COD_ALFA.length);
    /* con letras y sin forma de numero (12E45678): Sheets no lo convierte */
    if (/[A-Z]/.test(s) && !/^\d+E\d+$/.test(s)) return s;
  }
  return 'CG' + String(Date.now()).slice(-6);
}

/* Registros por usuario. Una sola lectura de "Observaciones" (id, usuario, recibido). */
function _codLeerObs(ss) {
  var sh = _sheet(ss, 'Observaciones', HEAD_OBS), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, 15).getValues();
}

/* ── Entrar con un codigo ─────────────────────────────────────────────────── */
function _codEntrar(P, cb) {
  var cod = _codNorm(P.codigo);
  if (cod.length < 6 || cod.length > 16) return _json({ ok: false, error: 'codigo' }, cb);
  var sid = getSheetId(); if (!sid) return _json({ ok: false, error: 'no_sheet' }, cb);
  var ss = SpreadsheetApp.openById(sid);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (_) { return _json({ ok: false, error: 'busy' }, cb); }
  try {
    var sh = _codSheet(ss), c = null;
    _codFilas(sh).forEach(function (x) { if (_codNorm(x.codigo) === cod) c = x; });
    if (!c) return _json({ ok: false, error: 'codigo' }, cb);
    var dev = String(P.dev || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 64);
    var shU = _usuarios(ss), lastU = shU.getLastRow(), ahora = new Date().toISOString();
    var email = '', name = '', nuevo = false, fin = _ymd(c.fin);
    /* El mismo telefono vuelve a su cuenta (p. ej. tras cerrar sesion), tambien
       con el codigo vencido o pausado: no se crea nada nuevo. */
    if (dev && lastU >= 2) {
      var vals = shU.getRange(2, 1, lastU - 1, 6).getValues();
      for (var i = 0; i < vals.length; i++) {
        var em = String(vals[i][0] || '').toLowerCase();
        if (String(vals[i][5] || '') === 'dev:' + dev && _codIdDeEmail(em) === c.id) {
          if (String(vals[i][3] || '').toLowerCase() === 'blocked') return _json({ ok: false, error: 'blocked' }, cb);
          email = em; name = String(vals[i][1] || '');
          shU.getRange(i + 2, 8).setValue(ahora);
          break;
        }
      }
    }
    if (!email) {
      if (String(c.estado || '').toLowerCase() !== 'activo') return _json({ ok: false, error: 'pausado' }, cb);
      var hoy = _codHoy(), ini = _ymd(c.inicio);
      if (ini && hoy < ini) return _json({ ok: false, error: 'aun_no', inicio: ini }, cb);
      if (fin && hoy > fin) return _json({ ok: false, error: 'vencido', fin: fin }, cb);
    }
    if (!email) {
      var n = _codNum(c.cuentas) + 1;
      sh.getRange(c.fila, 10).setValue(n);
      email = c.id + '-' + _codPad(n) + COD_DOM;
      name = 'Participante ' + (n < 10 ? '0' + n : n);
      shU.appendRow([email, name, '', 'active', 'practica', (dev ? 'dev:' + dev : ''), ahora, ahora, 0, '', '']);
      nuevo = true;
    }
    return _json({ ok: true, session: crearSesion(email), email: email, name: name, practica: true,
      nuevo: nuevo, grupo: String(c.nombre || ''), fin: fin }, cb);
  } finally { try { lock.releaseLock(); } catch (_) {} }
}

/* ── Administracion ───────────────────────────────────────────────────────── */
function _codAdmin(P, cb) {
  var u = verifyIdToken(P.id_token);
  if (!u) return { err: _json({ ok: false, error: 'auth' }, cb) };
  if (!_esAdmin(u.email)) return { err: _json({ ok: false, error: 'not_admin' }, cb) };
  var sid = getSheetId(); if (!sid) return { err: _json({ ok: false, error: 'no_sheet' }, cb) };
  return { u: u, ss: SpreadsheetApp.openById(sid) };
}

function _codListar(P, cb) {
  var a = _codAdmin(P, cb); if (a.err) return a.err;
  var filas = _codFilas(_codSheet(a.ss)), uso = {};
  _codLeerObs(a.ss).forEach(function (v) {
    var id = _codIdDeEmail(v[1]); if (id) uso[id] = (uso[id] || 0) + 1;
  });
  var hoy = _codHoy();
  return _json({ ok: true, hoy: hoy, def: COD_DEF, codigos: filas.map(function (c) {
    var fin = _ymd(c.fin);
    return { id: c.id, codigo: _codBonito(c.codigo), nombre: String(c.nombre || ''), inicio: _ymd(c.inicio), fin: fin,
      lim_dia: _codNum(c.lim_dia), lim_cuenta: _codNum(c.lim_cuenta), tope_total: _codNum(c.tope_total),
      estado: String(c.estado || '').toLowerCase(), vencido: !!(fin && hoy > fin),
      cuentas: _codNum(c.cuentas), registros: uso[c.id] || 0 };
  }).reverse() }, cb);
}

function _codCrear(P, cb) {
  var a = _codAdmin(P, cb); if (a.err) return a.err;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (_) { return _json({ ok: false, error: 'busy' }, cb); }
  try {
    var sh = _codSheet(a.ss), filas = _codFilas(sh), max = 0, usados = {};
    filas.forEach(function (c) {
      var m = /^c(\d+)$/.exec(c.id); if (m) max = Math.max(max, Number(m[1]));
      usados[_codNorm(c.codigo)] = 1;
    });
    var cod; do { cod = _codNuevoTexto(); } while (usados[cod]);
    var num = function (v, d) { return (v === undefined || v === '') ? d : _codNum(v); };
    var dias = Math.min(366, Math.max(1, num(P.dias, COD_DEF.dias)));
    var ini = new Date(), fin = new Date(ini.getTime() + dias * 86400000);
    var id = 'c' + (max + 1);
    var nombre = String(P.nombre || '').replace(/[<>]/g, '').trim().slice(0, 60) || ('Practica ' + (max + 1));
    sh.appendRow([id, cod, nombre, _codDia(ini), _codDia(fin),
      num(P.lim_dia, COD_DEF.lim_dia), num(P.lim_cuenta, COD_DEF.lim_cuenta), num(P.tope, COD_DEF.tope_total),
      'activo', 0, new Date().toISOString(), a.u.email]);
    /* Texto plano: Sheets no debe reinterpretar el codigo ni las fechas */
    try { sh.getRange(sh.getLastRow(), 2, 1, 4).setNumberFormat('@'); } catch (_) {}
    return _json({ ok: true, id: id, codigo: _codBonito(cod), nombre: nombre, fin: _codDia(fin) }, cb);
  } finally { try { lock.releaseLock(); } catch (_) {} }
}

function _codEstado(P, cb) {
  var a = _codAdmin(P, cb); if (a.err) return a.err;
  var id = String(P.id || '').toLowerCase(), est = String(P.estado || '').toLowerCase();
  if (est !== 'activo' && est !== 'pausado') return _json({ ok: false, error: 'estado' }, cb);
  var sh = _codSheet(a.ss), c = null;
  _codFilas(sh).forEach(function (x) { if (x.id === id) c = x; });
  if (!c) return _json({ ok: false, error: 'no_existe' }, cb);
  sh.getRange(c.fila, 9).setValue(est);
  return _json({ ok: true, id: id, estado: est }, cb);
}

/* Gancho en sesionGet. Devuelve una respuesta, o null para seguir. */
function codigosGet(P, cb) {
  var a = P && P.action;
  if (a === 'codigo') return _codEntrar(P, cb);
  if (a === 'codigos') return _codListar(P, cb);
  if (a === 'codigocrear') return _codCrear(P, cb);
  if (a === 'codigoestado') return _codEstado(P, cb);
  return null;
}

/* ── Cupos de envio ───────────────────────────────────────────────────────── */
/* Gancho en sesionPost, antes de guardar. Devuelve una respuesta de rechazo
   ({ok:false, error:'cupo_dia'|'cupo_cuenta'|'cupo_total', libres:N}) o null
   para seguir. El lote se acepta o se rechaza entero: la app reintenta con
   solo "libres" registros nuevos. Las ediciones de registros ya guardados no
   cuentan. Ante un error inesperado deja pasar el envio (no interrumpe la
   clase) y lo anota en el registro de ejecuciones. */
function codigosCupo(body) {
  try {
    var feats = (body && body.observaciones && body.observaciones.features) || [];
    if (!feats.length) return null;
    var user = verifyIdToken(body.idToken);
    if (!user || !user.email) return null;              // Code.gs respondera 'auth'
    var email = String(user.email).toLowerCase();
    if (_esAdmin(email)) return null;
    var sid = getSheetId(); if (!sid) return null;
    var ss = SpreadsheetApp.openById(sid);

    var cid = _codIdDeEmail(email), lim;
    if (cid) {
      var c = null;
      _codFilas(_codSheet(ss)).forEach(function (x) { if (x.id === cid) c = x; });
      lim = c ? { dia: _codNum(c.lim_dia), cuenta: _codNum(c.lim_cuenta), tope: _codNum(c.tope_total) }
              : { dia: COD_DEF.lim_dia, cuenta: COD_DEF.lim_cuenta, tope: COD_DEF.tope_total };
    } else {
      lim = { dia: GOOGLE_LIM_DIA, cuenta: GOOGLE_LIM_CUENTA, tope: 0 };
    }
    if (!lim.dia && !lim.cuenta && !lim.tope) return null;

    var hoy = _codHoy(), ids = {}, nHoy = 0, nCuenta = 0, nTotal = 0;
    _codLeerObs(ss).forEach(function (v) {
      var usr = String(v[1] || '').toLowerCase();
      if (cid && _codIdDeEmail(usr) === cid) nTotal++;
      if (usr !== email) return;
      ids[String(v[0])] = 1; nCuenta++;
      var rec = v[14], d = (rec instanceof Date) ? rec : (rec ? new Date(rec) : null);
      if (d && !isNaN(d.getTime()) && _codDia(d) === hoy) nHoy++;
    });

    var nuevos = 0, vistos = {};
    feats.forEach(function (f) {
      var id = String((f && f.properties && f.properties.id) || '');
      if (!id || ids[id] || vistos[id]) return;
      vistos[id] = 1; nuevos++;
    });
    if (!nuevos) return null;

    var libres = Infinity, motivo = '';
    var probar = function (limite, usados, nombre) {
      if (!limite) return;
      var r = Math.max(0, limite - usados);
      if (r < libres) { libres = r; motivo = nombre; }
    };
    probar(lim.dia, nHoy, 'cupo_dia');
    probar(lim.cuenta, nCuenta, 'cupo_cuenta');
    probar(lim.tope, nTotal, 'cupo_total');
    if (nuevos <= libres) return null;
    return _json({ ok: false, error: motivo, libres: libres });
  } catch (e) {
    console.warn('codigosCupo: ' + e);
    return null;
  }
}
