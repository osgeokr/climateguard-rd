/**
 * ClimateGuard - Registro de amenazas (servidor v3.10)
 * ===========================================================================
 * Archivo NUEVO del mismo proyecto (junto a Code.gs y Sesion.gs).
 *
 * La app v1.11 envia cada amenaza como una observacion mas (foto, coordenadas,
 * notas: se guardan en "Observaciones" como siempre) con estas propiedades
 * extra: tipo_registro='amenaza', amenaza, amenaza_sub, iucn_amenaza,
 * condicion, magnitud, accion. Este archivo guarda esas propiedades en la
 * pestana "Amenazas" (una fila por registro, mismo id que en Observaciones)
 * sin modificar Code.gs.
 *
 *  - guardarAmenazas(body): lo llama sesionPost() al recibir un envio.
 *  - action=amenazas: lectura para el sitio web (propias; el admin puede pedir
 *    las de un usuario con &email=... o todas con &email=*).
 *  - action=resetuser: antes de que Code.gs reinicie la cuenta, se borran
 *    tambien sus filas de "Amenazas".
 *
 * Codigos: clasificacion de amenazas de la UICN (Threats Classification
 * Scheme 3.3). Las claves (amenaza, amenaza_sub, condicion...) son las del
 * catalogo AMEN de la app.
 */

var HEAD_AM = ['id', 'usuario', 'amenaza', 'amenaza_sub', 'iucn_amenaza', 'condicion', 'magnitud', 'accion',
               'lat', 'lng', 'fecha', 'notas', 'recibido'];

function _amSheet(ss) { return _sheet(ss, 'Amenazas', HEAD_AM); }

/* Guarda/actualiza las amenazas de un envio. Nunca interrumpe el envio. */
function guardarAmenazas(body) {
  try {
    var feats = (body && body.observaciones && body.observaciones.features) || [];
    var am = feats.filter(function (f) { return f && f.properties && f.properties.tipo_registro === 'amenaza' && f.properties.id; });
    if (!am.length) return;
    var user = verifyIdToken(body.idToken);
    if (!user) return;
    var sid = getSheetId(); if (!sid) return;
    var ss = SpreadsheetApp.openById(sid);
    var shU = _usuarios(ss), row = _userRow(shU, user.email);
    if (row > 0) {
      var st = String(shU.getRange(row, 4).getValue() || 'active').toLowerCase();
      if (st === 'blocked' || st === 'pending') return;
    }
    var sh = _amSheet(ss), lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (_) {}
    try {
      var ahora = new Date().toISOString();
      am.forEach(function (f) {
        var p = f.properties, c = (f.geometry && f.geometry.coordinates) || [null, null];
        _upsert(sh, String(p.id), [String(p.id), user.email, String(p.amenaza || ''), String(p.amenaza_sub || ''),
          String(p.iucn_amenaza || ''), String(p.condicion || ''), String(p.magnitud || ''), String(p.accion || ''),
          (c[1] == null ? '' : c[1]), (c[0] == null ? '' : c[0]), String(p.fecha || ''), String(p.notas || ''), ahora]);
      });
    } finally { try { lock.releaseLock(); } catch (_) {} }
  } catch (e) { console.warn('guardarAmenazas: ' + e); }
}

/* action=amenazas (GET): propias, o de otro usuario / todas si es admin. */
function amenazasGet(P, cb) {
  var u = verifyIdToken(P.id_token);
  if (!u) return _json({ ok: false, error: 'auth' }, cb);
  var sid = getSheetId(); if (!sid) return _json({ ok: false, error: 'no_sheet' }, cb);
  var quien = u.email, pedido = String(P.email || '').toLowerCase().trim();
  if (pedido && pedido !== u.email) {
    if (!_esAdmin(u.email)) return _json({ ok: false, error: 'not_admin' }, cb);
    quien = pedido;
  }
  var sh = _amSheet(SpreadsheetApp.openById(sid)), last = sh.getLastRow(), out = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, HEAD_AM.length).getValues().forEach(function (v) {
      if (quien !== '*' && String(v[1]).toLowerCase() !== quien) return;
      var o = {}; HEAD_AM.forEach(function (h, i) { o[h] = v[i]; }); out.push(o);
    });
  }
  return _json({ ok: true, amenazas: out }, cb);
}

/* Borra las filas de "Amenazas" de un email (lo usa resetuser). */
function amenazasReset(email) {
  try {
    var sid = getSheetId(); if (!sid) return 0;
    var sh = _amSheet(SpreadsheetApp.openById(sid)), last = sh.getLastRow(), n = 0;
    if (last < 2) return 0;
    var v = sh.getRange(2, 2, last - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]).toLowerCase() === email) { sh.deleteRow(i + 2); n++; }
    }
    return n;
  } catch (e) { return 0; }
}
