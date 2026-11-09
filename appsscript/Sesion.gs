/**
 * ClimateGuard - Sesion propia (servidor v3.9)
 * ===========================================================================
 * Archivo NUEVO en el mismo proyecto de Apps Script que Codigo.gs (comparten
 * funciones y constantes: _json, _esAdmin, _usuarios, _userRow, _touchUser,
 * _certificarPuntos, _sheet, _usuarioAvatar, getSheetId, HEAD_OBS...).
 *
 * Que hace
 *  - La app verifica la identidad con Google UNA sola vez (action=login) y
 *    recibe una clave de sesion firmada (HMAC-SHA256) que no caduca.
 *  - Despues la app envia ?session=... (GET) o {session:...} (POST) en lugar
 *    del id_token de Google, que caduca cada hora.
 *  - El administrador puede cerrar la sesion de una cuenta (logoutuser) o de
 *    todos los participantes (logoutall). Las cuentas de ADMIN_EMAILS no se
 *    ven afectadas por logoutall.
 *  - action=avatar devuelve la foto de credencial guardada en Drive, para
 *    restaurarla en un telefono nuevo.
 *
 * Seguridad
 *  - La clave secreta se crea sola en Propiedades del script (SESSION_SECRET).
 *    Si se borra o cambia, TODAS las sesiones dejan de valer (cada persona
 *    vuelve a iniciar sesion con Google una vez).
 *  - El estado del usuario (active/blocked/pending) se sigue comprobando en
 *    cada peticion, igual que antes.
 *  - Si se borra la fila del usuario en Usuarios (p. ej. resetuser), su
 *    sesion deja de valer.
 *
 * Instalacion: ver los 3 cambios marcados con "v3.9" en Codigo.gs.
 */

var SES_PREFIX = 'cgs1.';

function _sesProps() { return PropertiesService.getScriptProperties(); }

function _sesSecret() {
  var p = _sesProps(), s = p.getProperty('SESSION_SECRET');
  if (s) return s;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (_) {}
  try {
    s = p.getProperty('SESSION_SECRET');
    if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SESSION_SECRET', s); }
  } finally { try { lock.releaseLock(); } catch (_) {} }
  return s;
}

function _sesB64(x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ''); }
function _sesUnB64(s) {
  s = String(s || ''); while (s.length % 4) s += '=';
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString();
}
function _sesVer(email) { return Number(_sesProps().getProperty('SV:' + email) || 0); }
function _sesEpoch() { return Number(_sesProps().getProperty('SV_ALL') || 0); }
function _sesFirma(email, uv, ep) {
  return _sesB64(Utilities.computeHmacSha256Signature(email + '|' + uv + '|' + ep, _sesSecret()));
}

/* Crea la clave de sesion de un email ya verificado por Google. */
function crearSesion(email) {
  email = String(email || '').toLowerCase();
  var uv = _sesVer(email), ep = _esAdmin(email) ? 0 : _sesEpoch();
  return SES_PREFIX + _sesB64(email) + '.' + uv + '.' + ep + '.' + _sesFirma(email, uv, ep);
}

/* null = clave falsa o ilegible; {valid:false} = cerrada por el admin. */
function _sesParse(tok) {
  tok = String(tok || '');
  if (tok.indexOf(SES_PREFIX) !== 0) return null;
  var parts = tok.slice(SES_PREFIX.length).split('.');
  if (parts.length !== 4) return null;
  var email;
  try { email = _sesUnB64(parts[0]).toLowerCase(); } catch (_) { return null; }
  var uv = Number(parts[1]), ep = Number(parts[2]);
  if (!email || isNaN(uv) || isNaN(ep)) return null;
  if (_sesFirma(email, uv, ep) !== parts[3]) return null;
  var valid = (uv === _sesVer(email)) && (_esAdmin(email) || ep === _sesEpoch());
  return { email: email, valid: valid };
}

/* Lo llama verifyIdToken cuando recibe una clave de sesion.
   Devuelve el mismo formato que verifyIdToken (con nombre y foto del padron,
   para que _touchUser no los borre). */
function verifySession(tok) {
  var ps = _sesParse(tok);
  if (!ps || !ps.valid) return null;
  var id = getSheetId(); if (!id) return null;
  var sh = _usuarios(SpreadsheetApp.openById(id));
  var row = _userRow(sh, ps.email);
  if (row < 0) return null;
  var v = sh.getRange(row, 1, 1, 6).getValues()[0];
  return { email: ps.email, name: String(v[1] || ''), picture: String(v[2] || ''), sub: String(v[5] || ''), session: true };
}

/* Gancho al inicio de doGet. Devuelve una respuesta, o null para seguir. */
function sesionGet(e, cb) {
  var P = (e && e.parameter) || {};
  var a = P.action;

  if (P.session && !P.id_token) {
    var ps = _sesParse(P.session);
    if (!ps) return _json({ ok: false, error: 'auth' }, cb);
    if (!ps.valid) return _json({ ok: false, error: 'revoked' }, cb);
    P.id_token = P.session;           // las acciones existentes la aceptan via verifyIdToken
  }

  // v3.10: amenazas (Amenazas.gs)
  if (a === 'amenazas') return amenazasGet(P, cb);
  if (a === 'resetuser') {
    var rz = verifyIdToken(P.id_token);
    var rem = String(P.email || '').toLowerCase().trim();
    if (rz && _esAdmin(rz.email) && rem.indexOf('@') > 0) amenazasReset(rem);
    // sigue en Code.gs, que reinicia el resto de la cuenta
  }

  // Primer inicio de sesion: id_token de Google -> clave de sesion propia
  if (a === 'login') {
    if (String(P.id_token || '').indexOf(SES_PREFIX) === 0) return _json({ ok: false, error: 'auth' }, cb);
    var u = verifyIdToken(P.id_token);
    if (!u) return _json({ ok: false, error: 'auth' }, cb);
    var sid = getSheetId(), ss = sid ? SpreadsheetApp.openById(sid) : null;
    var st = ss ? String(_touchUser(ss, u)).toLowerCase() : 'active';
    if (st === 'blocked' || st === 'pending') return _json({ ok: false, error: st, status: st }, cb);
    var pr = ss ? _certificarPuntos(_sheet(ss, 'Observaciones', HEAD_OBS), ss, u.email) : {};
    return _json({ ok: true, session: crearSesion(u.email), email: u.email, status: st,
      puntos: pr.puntos, nivel: pr.nivel, avatar: !!(ss && _usuarioAvatar(ss, u.email)) }, cb);
  }

  // Foto de credencial guardada en Drive (para restaurarla en el telefono)
  if (a === 'avatar') {
    var ua = verifyIdToken(P.id_token);
    if (!ua) return _json({ ok: false, error: 'auth' }, cb);
    var sid2 = getSheetId(); if (!sid2) return _json({ ok: false, error: 'no_sheet' }, cb);
    var url = _usuarioAvatar(SpreadsheetApp.openById(sid2), ua.email);
    var fid = (String(url || '').match(/[-\w]{25,}/) || [])[0];
    if (!fid) return _json({ ok: true, dataB64: '' }, cb);
    try {
      var blob = DriveApp.getFileById(fid).getBlob();
      return _json({ ok: true, mime: blob.getContentType(), dataB64: Utilities.base64Encode(blob.getBytes()) }, cb);
    } catch (ef) { return _json({ ok: false, error: 'not_found' }, cb); }
  }

  // Cerrar sesiones (solo admin)
  if (a === 'logoutuser' || a === 'logoutall') {
    var ad = verifyIdToken(P.id_token);
    if (!ad) return _json({ ok: false, error: 'auth' }, cb);
    if (!_esAdmin(ad.email)) return _json({ ok: false, error: 'not_admin' }, cb);
    var props = _sesProps();
    if (a === 'logoutuser') {
      var em = String(P.email || '').toLowerCase().trim();
      if (!em || em.indexOf('@') < 0) return _json({ ok: false, error: 'no_email' }, cb);
      props.setProperty('SV:' + em, String(_sesVer(em) + 1));
      return _json({ ok: true, email: em }, cb);
    }
    props.setProperty('SV_ALL', String(_sesEpoch() + 1));
    var n = 0, sid3 = getSheetId();
    if (sid3) {
      var shU = _usuarios(SpreadsheetApp.openById(sid3)), last = shU.getLastRow();
      if (last >= 2) shU.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
        var x = String(r[0] || '').toLowerCase(); if (x && !_esAdmin(x)) n++; });
    }
    return _json({ ok: true, count: n }, cb);
  }
  return null;
}

/* Gancho en doPost, justo despues de leer el cuerpo. */
function sesionPost(body) {
  if (body && body.session && !body.idToken) {
    var ps = _sesParse(body.session);
    if (!ps) return _json({ ok: false, error: 'auth' });
    if (!ps.valid) return _json({ ok: false, error: 'revoked' });
    body.idToken = body.session;
  }
  guardarAmenazas(body);   // v3.10 (Amenazas.gs): nunca interrumpe el envio
  return null;
}
