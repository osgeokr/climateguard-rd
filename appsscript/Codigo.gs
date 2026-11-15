/**
 * ClimateGuard - Receptor de datos + Autenticacion Google (Apps Script Web App)
 * ===========================================================================
 * v2 (login real): en lugar de un TOKEN publico compartido, cada envio debe
 * traer un ID token de Google (OpenID Connect). El servidor lo verifica con
 * el endpoint tokeninfo de Google, obtiene el email VERIFICADO del usuario y
 * solo entonces guarda los datos. Ademas mantiene un padron "Usuarios" con
 * estado por usuario (active / blocked / pending) y limita la tasa de envios.
 *
 * v3.9 (sesion): clave de sesion propia sin caducidad; action=login/avatar/
 *   logoutuser/logoutall. Codigo en Sesion.gs (archivo aparte del proyecto).
 * v3.8 (app v1.9): envio robusto. LockService; la fila solo se actualiza si
 *   el id es del MISMO usuario (si no, error 'owner'); ids de demostracion
 *   (demo-*) se rechazan y no suman puntos; una sola lectura de la hoja
 *   (indice) en vez de recorrerla por cada registro; respuesta con resultado
 *   por registro (items) que la app v1.9 lee para marcar 'enviado'. Se
 *   elimina la consulta publica ?usuario= y el token compartido antiguo.
 *   Nuevas columnas: origen, app_ver. Ejecutar actualizarEsquema una vez.
 * v3.7 (admin): nuevo action=resetuser (solo admin) para reiniciar una
 *   cuenta de prueba: borra sus filas en las 5 pestanas (Usuarios opcional)
 *   y envia su carpeta de Drive a la papelera.
 * v3.6 (admin): action=stats devuelve recorridos por usuario; nuevo
 *   action=userdata (solo admin) con observaciones/recorridos de un usuario.
 * v3.5 (puntos): _certificarPuntos ahora tambien cuenta los recorridos
 *   (10 + 5 si distancia >= 200 m), igual que la app.
 * v3.4 (web-proxy): archivos PRIVADOS; el sitio web lee imagenes/tracks via
 *   action=file (verifica login y que el archivo sea del usuario o admin).
 * v3.3 (web-prep): UNA carpeta por usuario; foto_url/geo_url en las hojas;
 *   avatar de usuario a Drive; endpoints de lectura mine/stats/asistlist/users.
 * v3.2 (asistencia): offline-first; se registra la fecha del escaneo (cliente).
 * v3.1 (asistencia): QR corto (prefijo del hash) + coincidencia por prefijo.
 * v3 (asistencia, ago-2026): un administrador (ADMIN_EMAILS) puede registrar
 * la asistencia de los miembros escaneando la QR de su credencial. La QR
 * contiene el SHA-256 del email (ID anonimo). El servidor recalcula ese hash
 * para cada email del padron, identifica al miembro y anota una fila en la
 * pestana "Asistencia" (una por miembro y dia).
 *
 * Modelo de acceso: REGISTRO ABIERTO (③). Cualquiera con cuenta Google puede
 * entrar y queda registrado. Para "cerrar" luego a ciertos dominios, basta
 * rellenar ALLOWED_DOMAINS/ALLOWED_EMAILS (sin volver a desplegar la app).
 *
 * ▓▓▓ PASOS (una sola vez) ▓▓▓
 *  1. Pega TODO este codigo en el editor de Apps Script y guarda (💾).
 *  2. Ejecuta  actualizarEsquema  una vez (crea las pestanas Usuarios y
 *     Asistencia y pone al dia los encabezados en tu hoja ClimateGuard_DB).
 *     [Si aun no tienes hoja, ejecuta primero  crearBaseDeDatos ].
 *  3. Implementar > Gestionar implementaciones > (lapiz ✏️ Editar) >
 *     Version: "Nueva version" > Implementar.
 *        - Ejecutar como: Yo
 *        - Quien tiene acceso: Cualquier usuario   ← IMPRESCINDIBLE
 *     La URL /exec NO cambia si reutilizas el mismo proyecto.
 */

// ── CONFIGURACION ──────────────────────────────────────────────────────────
const CLIENT_ID   = '918300818571-39mnd09t4q7k1kro6olhiafjme803k5b.apps.googleusercontent.com';
const CARPETA_ID  = '1A1VbIRSmD4TlPTF2yQUw44nLzcZ3sDvn';   // carpeta Drive fotos/geojson

// Registro abierto por defecto (listas vacias). Para restringir el acceso,
// rellena los dominios permitidos (ej. 'knps.or.kr') y/o correos sueltos.
// En cuanto una de las dos listas tenga algo, los que NO coincidan quedan
// en estado 'pending' (deben ser aprobados a mano en la hoja Usuarios).
const ALLOWED_DOMAINS = [];   // p.ej. ['knps.or.kr','koica.go.kr']
const ALLOWED_EMAILS  = [];   // p.ej. ['maria@gmail.com']

// Administradores: estos correos pueden registrar asistencia (accion=attend).
// El control es por email VERIFICADO del ID token, no se puede falsificar
// desde el cliente. Anade mas correos separados por coma si hace falta.
const ADMIN_EMAILS = ['bhyu@knps.or.kr'];

// Limite de tasa por usuario (filas de Observaciones recibidas por ventana).
const RATE_PER_HOUR = 200;
const RATE_PER_DAY  = 1000;

const BONO        = 1;                    // los puntos se "reconocen" al enviar
const NIVELES_MIN = [0, 100, 250, 500, 1000];

const HEAD_OBS = ['id','usuario','especie','clase','iucn','cant','notas','lat','lng','accuracy','fecha','updated_at','foto','submissionId','recibido','estado','altitude','alt_accuracy','foto_url','origen','app_ver'];
const COL_OBS_ESTADO = 16, COL_OBS_FOTOURL = 19;   // 1-based, fijos (no dependen del largo del encabezado)
const COL_REC_GEOURL = 10;
const HEAD_REC = ['id','usuario','name','points','distance_m','start','end','submissionId','recibido','geo_url','app_ver'];
function _esDemo(id){ return /^demo-/i.test(String(id||'')); }
const HEAD_PTS = ['usuario','puntos','nivel','observaciones','actualizado'];
const HEAD_USR = ['email','name','picture','status','role','sub','created_at','last_login','submissions','folder_id','avatar_url'];
const HEAD_ASIST = ['fecha','email','nombre','recibido','registrado_por'];   // v3 asistencia

// ── CREACION DE LA BASE ─────────────────────────────────────────────────────
function crearBaseDeDatos() {
  var ss = SpreadsheetApp.create('ClimateGuard_DB');
  var o = ss.getSheets()[0]; o.setName('Observaciones'); o.appendRow(HEAD_OBS);
  ss.insertSheet('Recorridos').appendRow(HEAD_REC);
  ss.insertSheet('Puntos').appendRow(HEAD_PTS);
  ss.insertSheet('Usuarios').appendRow(HEAD_USR);
  ss.insertSheet('Asistencia').appendRow(HEAD_ASIST);
  try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CARPETA_ID)); } catch (e) {}
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('✅ Base creada. SHEET_ID = ' + ss.getId());
  Logger.log('🔗 ' + ss.getUrl());
  return ss.getUrl();
}

/* Ejecuta UNA vez si ya tenias la hoja: crea Usuarios/Asistencia y encabezados. */
function actualizarEsquema(){
  var id=getSheetId(); if(!id){Logger.log('Ejecuta primero crearBaseDeDatos()');return;}
  var ss=SpreadsheetApp.openById(id);
  var shO=ss.getSheetByName('Observaciones'); if(shO) shO.getRange(1,1,1,HEAD_OBS.length).setValues([HEAD_OBS]);
  var shR=ss.getSheetByName('Recorridos'); if(shR) shR.getRange(1,1,1,HEAD_REC.length).setValues([HEAD_REC]);
  _sheet(ss,'Usuarios',HEAD_USR);
  _sheet(ss,'Asistencia',HEAD_ASIST);
  Logger.log('Esquema actualizado (Usuarios + Asistencia + encabezados).');
  marcarDemo();
}

function verificar() {
  var id = getSheetId();
  if (!id) { Logger.log('⚠️ Ejecuta primero crearBaseDeDatos()'); return; }
  var ss = SpreadsheetApp.openById(id);
  Logger.log('✅ ' + ss.getName() + ' | ' + ss.getUrl());
}
function getSheetId() {
  return PropertiesService.getScriptProperties().getProperty('SHEET_ID') || '';
}

// ── AUTENTICACION: verificar el ID token de Google ─────────────────────────
function verifyIdToken(idToken) {
  if (!idToken) return null;
  if (String(idToken).indexOf('cgs1.') === 0) return verifySession(idToken);   // v3.9: sesion propia (Sesion.gs)
  try {
    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var p = JSON.parse(resp.getContentText());
    if (p.aud !== CLIENT_ID) return null;                          // token para OTRA app
    if (!(p.email_verified === true || p.email_verified === 'true')) return null;
    if (p.exp && (Number(p.exp) * 1000) < Date.now()) return null; // expirado
    return { email: String(p.email || '').toLowerCase(), name: p.name || '', picture: p.picture || '', sub: p.sub || '' };
  } catch (e) { return null; }
}

// ── PADRON DE USUARIOS ─────────────────────────────────────────────────────
function _usuarios(ss){ return _sheet(ss, 'Usuarios', HEAD_USR); }
function _userRow(sh, email){
  var last = sh.getLastRow(); if (last < 2) return -1;
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for (var i=0;i<ids.length;i++){ if (String(ids[i][0]).toLowerCase() === email) return i+2; }
  return -1;
}
function _statusForEmail(email){
  if (!ALLOWED_DOMAINS.length && !ALLOWED_EMAILS.length) return 'active';   // registro abierto
  var dom = (email.split('@')[1] || '').toLowerCase();
  if (ALLOWED_DOMAINS.map(function(d){return d.toLowerCase();}).indexOf(dom) >= 0) return 'active';
  if (ALLOWED_EMAILS.map(function(e){return e.toLowerCase();}).indexOf(email) >= 0) return 'active';
  return 'pending';
}
/* Crea o actualiza al usuario y devuelve su estado (active/blocked/pending). */
function _touchUser(ss, u){
  var sh = _usuarios(ss); var row = _userRow(sh, u.email); var now = new Date().toISOString();
  if (row < 0){
    var st = _statusForEmail(u.email);
    sh.appendRow([u.email, u.name, u.picture, st, '', u.sub, now, now, 0, '', '']);
    return st;
  }
  var st2 = String(sh.getRange(row,4).getValue() || 'active');
  sh.getRange(row,2).setValue(u.name);
  sh.getRange(row,3).setValue(u.picture);
  sh.getRange(row,8).setValue(now);
  return st2;
}
function _bumpSubmissions(ss, email){
  var sh = _usuarios(ss); var row = _userRow(sh, email);
  if (row > 0){ var c = Number(sh.getRange(row,9).getValue() || 0); sh.getRange(row,9).setValue(c+1); }
}
/* Limite de tasa: cuenta filas de Observaciones de este email en la ventana. */
function _rateOk(shO, email){
  var last = shO.getLastRow(); if (last < 2) return true;
  var w = shO.getLastColumn();
  var vals = shO.getRange(2,1,last-1,w).getValues();
  var now = Date.now(), h = 0, d = 0;
  for (var i=0;i<vals.length;i++){
    if (String(vals[i][1]).toLowerCase() !== email) continue;    // col B = usuario (email)
    var rec = vals[i][14];                                        // col O = recibido (ISO)
    var t = rec ? Date.parse(rec) : 0; if (!t) continue;
    if (now - t < 3600e3) h++;
    if (now - t < 86400e3) d++;
  }
  return h < RATE_PER_HOUR && d < RATE_PER_DAY;
}

// ── ASISTENCIA (v3) ─────────────────────────────────────────────────────────
function _esAdmin(email){
  return ADMIN_EMAILS.map(function(e){return e.toLowerCase();}).indexOf(String(email||'').toLowerCase()) >= 0;
}
/* SHA-256 en hex (minusculas). Debe coincidir con la QR de la credencial,
   que codifica el hash del email en minusculas y sin espacios. */
function _sha256hex(str){
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i=0;i<raw.length;i++){ var b=(raw[i]+256)%256; hex += ('0'+b.toString(16)).slice(-2); }
  return hex;
}
/* Busca en el padron el miembro cuyo email produce ese hash. */
function _memberByHash(ss, hash){
  hash = String(hash||'').toLowerCase();
  if (!/^[0-9a-f]{12,64}$/.test(hash)) return null;
  var sh = _usuarios(ss); var last = sh.getLastRow(); if (last < 2) return null;
  var vals = sh.getRange(2,1,last-1,2).getValues();   // email, name
  for (var i=0;i<vals.length;i++){
    var email = String(vals[i][0]||'').toLowerCase().trim(); if (!email) continue;
    if (_sha256hex(email).indexOf(hash) === 0) return { email: email, name: (String(vals[i][1]||'') || email) };
  }
  return null;
}
function _ymd(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v||'').slice(0,10);
}
/* Anota asistencia (una fila por miembro y dia). Devuelve true si YA estaba. */
function _marcarAsistencia(ss, date, email, name, admin){
  var sh = _sheet(ss, 'Asistencia', HEAD_ASIST);
  var last = sh.getLastRow();
  if (last >= 2){
    var vals = sh.getRange(2,1,last-1,2).getValues();   // fecha, email
    for (var i=0;i<vals.length;i++){
      if (_ymd(vals[i][0]) === date && String(vals[i][1]||'').toLowerCase() === email) return true;
    }
  }
  sh.appendRow([date, email, name, new Date().toISOString(), admin]);
  return false;
}

// ── RECEPCION DE DATOS ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return _json({ ok:false, error:'sin datos' });
    var body = JSON.parse(e.postData.contents);
    var _sesP = sesionPost(body); if (_sesP) return _sesP;   // v3.9: sesion propia (Sesion.gs)

    // 1) Autenticacion: ID token de Google (unico metodo desde v3.8)
    var user = verifyIdToken(body.idToken);
    if (!user || !user.email) return _json({ ok:false, error:'auth' });
    var usuario = user.email;
    var sid = String(body.submissionId || '');
    var appVer = String(body.ver || '');

    var sheetId = getSheetId();
    var ss = sheetId ? SpreadsheetApp.openById(sheetId) : null;
    if (!ss) return _json({ ok:false, error:'no_sheet' });

    // 2) Padron: registrar/actualizar y comprobar estado
    var st = String(_touchUser(ss, user)).toLowerCase();
    if (st === 'blocked') return _json({ ok:false, error:'blocked' });
    if (st === 'pending') return _json({ ok:false, error:'pending' });

    // 3) Seccion critica: evita filas duplicadas si llegan dos envios a la vez
    var lock = LockService.getScriptLock();
    try { lock.waitLock(25000); } catch (eL) { return _json({ ok:false, error:'busy' }); }
    try {
      var shO = _sheet(ss, 'Observaciones', HEAD_OBS);
      var shR = _sheet(ss, 'Recorridos', HEAD_REC);
      var ixO = _indice(shO), ixR = _indice(shR);   // UNA lectura por hoja

      if (!_rateOkIdx(ixO, usuario)) return _json({ ok:false, error:'rate' });

      var carpeta = _carpetaUsuario(ss, usuario);
      var fotosByName = {};
      (body.fotos || []).forEach(function(ff){ if (ff && ff.name) fotosByName[ff.name] = ff; });
      if (carpeta && body.avatar && body.avatar.dataB64) {
        try {
          var avBlob = Utilities.newBlob(Utilities.base64Decode(body.avatar.dataB64), body.avatar.mime || 'image/jpeg', 'avatar.jpg');
          var avFile = _upsertFile(carpeta, 'avatar.jpg', avBlob);
          _setUsuarioAvatar(ss, usuario, avFile.getUrl());
        } catch (eav) {}
      }

      var items = [];
      var obsF = (body.observaciones && body.observaciones.features) || [];
      var recF = (body.recorridos && body.recorridos.features) || [];
      var ahora = new Date().toISOString();

      obsF.forEach(function(ft){
        var p = ft.properties || {}, g = ft.geometry || {}, c = (g.coordinates || [null, null]);
        var id = String(p.id || '');
        if (!id) { items.push({ tipo:'obs', id:id, saved:false, error:'no_id' }); return; }
        if (_esDemo(id)) { items.push({ tipo:'obs', id:id, saved:false, error:'demo' }); return; }
        var prev = ixO.byId[id];
        if (prev && prev.usuario && prev.usuario !== usuario) { items.push({ tipo:'obs', id:id, saved:false, error:'owner' }); return; }
        var estado = prev ? (prev.row[COL_OBS_ESTADO-1] || 'valido') : 'valido';
        var fotoUrl = prev ? (prev.row[COL_OBS_FOTOURL-1] || '') : '';
        var fotoOk = !p.foto;
        if (carpeta && p.foto && fotosByName[p.foto]) {
          try {
            var fx = fotosByName[p.foto];
            var mime = fx.mime || 'image/jpeg';
            var ext = mime === 'image/webp' ? '.webp' : (mime === 'image/png' ? '.png' : '.jpg');
            var fb = Utilities.newBlob(Utilities.base64Decode(fx.dataB64), mime, 'obs_' + id + ext);
            fotoUrl = _upsertFile(carpeta, 'obs_' + id + ext, fb, 'obs_' + id + '.').getUrl();
            fotoOk = true;
          } catch (e2) { fotoOk = false; }
        } else if (p.foto && fotoUrl) { fotoOk = true; }   // foto ya guardada (edicion sin reenviar foto)
        var fila = [id, usuario, p.especie || '', p.clase || '', p.iucn || '',
          p.cant || 1, p.notas || '', c[1], c[0], (p.accuracy == null ? '' : p.accuracy), p.fecha || '', p.updated_at || '',
          p.foto || '', sid, ahora, estado,
          (p.altitude == null ? '' : p.altitude), (p.alt_accuracy == null ? '' : p.alt_accuracy), fotoUrl,
          p.origen || '', appVer];
        _upsertIdx(shO, ixO, id, usuario, fila);
        items.push({ tipo:'obs', id:id, saved:true, foto:fotoOk, foto_url:fotoUrl });
      });

      recF.forEach(function(ft){
        var p = ft.properties || {};
        var id = String(p.id || '');
        if (!id) { items.push({ tipo:'trk', id:id, saved:false, error:'no_id' }); return; }
        if (_esDemo(id)) { items.push({ tipo:'trk', id:id, saved:false, error:'demo' }); return; }
        var prev = ixR.byId[id];
        if (prev && prev.usuario && prev.usuario !== usuario) { items.push({ tipo:'trk', id:id, saved:false, error:'owner' }); return; }
        var geoUrl = prev ? (prev.row[COL_REC_GEOURL-1] || '') : '';
        if (carpeta) {
          try {
            var gb = Utilities.newBlob(JSON.stringify({ type:'FeatureCollection', features:[ft] }), 'application/geo+json', 'track_' + id + '.geojson');
            geoUrl = _upsertFile(carpeta, 'track_' + id + '.geojson', gb).getUrl();
          } catch (e4) {}
        }
        _upsertIdx(shR, ixR, id, usuario, [id, usuario, p.name || '', p.points || '', p.distance_m || '',
          p.start || '', p.end || '', sid, ahora, geoUrl, appVer]);
        items.push({ tipo:'trk', id:id, saved:true, geo_url:geoUrl });
      });

      var r = _certificarPuntos(shO, ss, usuario);
      _bumpSubmissions(ss, usuario);
      return _json({ ok:true, usuario:usuario, items:items, puntos:r.puntos, nivel:r.nivel, server:'3.8',
        carpeta:(carpeta ? carpeta.getName() : ''), sheet:true });
    } finally { try { lock.releaseLock(); } catch (eR) {} }
  } catch (err) {
    return _json({ ok:false, error:String(err) });
  }
}

// ── GET: login/estado (JSONP), consulta de puntos y asistencia ──────────────
function doGet(e) {
  var cb = e && e.parameter && e.parameter.callback;
  try {
    var _ses = sesionGet(e, cb); if (_ses) return _ses;   // v3.9: login/avatar/logout + ?session= (Sesion.gs)
    var action = e && e.parameter && e.parameter.action;
    var sheetId = getSheetId();
    var ss = sheetId ? SpreadsheetApp.openById(sheetId) : null;

    // Verificacion de sesion al iniciar: la app envia el id_token y recibe el estado
    if (action === 'me') {
      var user = verifyIdToken(e.parameter.id_token);
      if (!user) return _json({ ok:false, error:'auth' }, cb);
      var st = ss ? _touchUser(ss, user) : 'active';
      var puntos = null, nivel = null;
      if (ss) { var pr = _certificarPuntos(_sheet(ss,'Observaciones',HEAD_OBS), ss, user.email); puntos = pr.puntos; nivel = pr.nivel; }
      return _json({ ok:true, email:user.email, name:user.name, picture:user.picture, status:st, puntos:puntos, nivel:nivel, admin:_esAdmin(user.email), server:'3.8' }, cb);
    }

    // Registro de asistencia (solo administrador): la app escanea la QR de la
    // credencial (hash del email) y la envia aqui con el id_token del admin.
    if (action === 'attend') {
      var admin = verifyIdToken(e.parameter.id_token);
      if (!admin) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(admin.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var miembro = _memberByHash(ss, e.parameter.hash);
      if (!miembro) return _json({ ok:false, error:'unknown' }, cb);
      var fecha = String(e.parameter.date||'').slice(0,10); // fecha del escaneo (cliente)
      if(!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var dup = _marcarAsistencia(ss, fecha, miembro.email, miembro.name, admin.email);
      return _json({ ok:true, name:miembro.name, email:miembro.email, dup:dup, date:fecha }, cb);
    }

    // Datos propios del usuario (para el sitio web)
    if (action === 'mine') {
      var um = verifyIdToken(e.parameter.id_token);
      if (!um) return _json({ ok:false, error:'auth' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var em = um.email;
      var obs = _rowsFor(_sheet(ss,'Observaciones',HEAD_OBS), 1, em, HEAD_OBS);
      var rec = _rowsFor(_sheet(ss,'Recorridos',HEAD_REC), 1, em, HEAD_REC);
      var pm = _certificarPuntos(_sheet(ss,'Observaciones',HEAD_OBS), ss, em);
      return _json({ ok:true, email:em, name:um.name, puntos:pm.puntos, nivel:pm.nivel, avatar:_usuarioAvatar(ss,em), observaciones:obs, recorridos:rec }, cb);
    }
    // Estadisticas por usuario (solo admin)
    if (action === 'stats') {
      var as = verifyIdToken(e.parameter.id_token);
      if (!as) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(as.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var shPs = _sheet(ss,'Puntos',HEAD_PTS); var lp = shPs.getLastRow(); var stats=[];
      var _trk = {};
      var shRc = _sheet(ss,'Recorridos',HEAD_REC); var lrc = shRc.getLastRow();
      if (lrc>=2){ var vrc = shRc.getRange(2,1,lrc-1,HEAD_REC.length).getValues();
        for (var kc=0;kc<vrc.length;kc++){ var uc=String(vrc[kc][1]||'').toLowerCase(); if(uc && !_esDemo(vrc[kc][0])) _trk[uc]=(_trk[uc]||0)+1; } }
      if (lp>=2){ var vp=shPs.getRange(2,1,lp-1,HEAD_PTS.length).getValues();
        for (var ip=0;ip<vp.length;ip++){ var upe=String(vp[ip][0]||'').toLowerCase();
          stats.push({ usuario:vp[ip][0], puntos:vp[ip][1], nivel:vp[ip][2], observaciones:vp[ip][3], recorridos:(_trk[upe]||0), actualizado:vp[ip][4] }); } }
      return _json({ ok:true, stats:stats }, cb);
    }
    // Reiniciar una cuenta (solo admin): borra filas del usuario y su carpeta Drive.
    if (action === 'resetuser') {
      var rs = verifyIdToken(e.parameter.id_token);
      if (!rs) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(rs.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var rtgt = String(e.parameter.email||'').toLowerCase().trim();
      if (!rtgt || rtgt.indexOf('@')<0) return _json({ ok:false, error:'no_email' }, cb);
      var rdelAcct = String(e.parameter.account||'')==='1';
      var rrep = _resetUsuarioServer(ss, rtgt, rdelAcct);
      return _json({ ok:true, email:rtgt, account_deleted:rdelAcct, deleted:rrep }, cb);
    }
    // Datos de un usuario (solo admin): observaciones + recorridos
    if (action === 'userdata') {
      var ud = verifyIdToken(e.parameter.id_token);
      if (!ud) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(ud.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var tgt = String(e.parameter.email||'').toLowerCase();
      if (!tgt) return _json({ ok:false, error:'no_email' }, cb);
      var obsU = _rowsFor(_sheet(ss,'Observaciones',HEAD_OBS), 1, tgt, HEAD_OBS);
      var recU = _rowsFor(_sheet(ss,'Recorridos',HEAD_REC), 1, tgt, HEAD_REC);
      var pmU = _certificarPuntos(_sheet(ss,'Observaciones',HEAD_OBS), ss, tgt);
      return _json({ ok:true, email:tgt, puntos:pmU.puntos, nivel:pmU.nivel, observaciones:obsU, recorridos:recU }, cb);
    }
    // Lista de asistencia (solo admin)
    if (action === 'asistlist') {
      var aa = verifyIdToken(e.parameter.id_token);
      if (!aa) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(aa.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var shAs = _sheet(ss,'Asistencia',HEAD_ASIST); var la = shAs.getLastRow(); var asist=[];
      if (la>=2){ var va=shAs.getRange(2,1,la-1,HEAD_ASIST.length).getValues();
        for (var ia=0;ia<va.length;ia++) asist.push({ fecha:_ymd(va[ia][0]), email:va[ia][1], nombre:va[ia][2], recibido:va[ia][3], registrado_por:va[ia][4] }); }
      return _json({ ok:true, asistencia:asist }, cb);
    }
    // Padron de usuarios (solo admin)
    if (action === 'users') {
      var au = verifyIdToken(e.parameter.id_token);
      if (!au) return _json({ ok:false, error:'auth' }, cb);
      if (!_esAdmin(au.email)) return _json({ ok:false, error:'not_admin' }, cb);
      if (!ss) return _json({ ok:false, error:'no_sheet' }, cb);
      var shUs = _usuarios(ss); var lu = shUs.getLastRow(); var users=[];
      if (lu>=2){ var vu=shUs.getRange(2,1,lu-1,HEAD_USR.length).getValues();
        for (var iu=0;iu<vu.length;iu++) users.push({ email:vu[iu][0], name:vu[iu][1], picture:vu[iu][2], status:vu[iu][3], created_at:vu[iu][6], last_login:vu[iu][7], submissions:vu[iu][8], avatar_url:vu[iu][10] }); }
      return _json({ ok:true, users:users }, cb);
    }
    // Proxy de archivo privado (imagen/track): requiere login; propietario o admin
    if (action === 'file') {
      var uf = verifyIdToken(e.parameter.id_token);
      if (!uf) return _json({ ok:false, error:'auth' }, cb);
      var fid = String(e.parameter.id||'').trim();
      if (!/^[-\w]{20,}$/.test(fid)) return _json({ ok:false, error:'bad_id' }, cb);
      try {
        var file = DriveApp.getFileById(fid);
        var permit = _esAdmin(uf.email);
        if (!permit) { var pit = file.getParents(); while (pit.hasNext()) { if (String(pit.next().getName()).toLowerCase() === uf.email) { permit = true; break; } } }
        if (!permit) return _json({ ok:false, error:'forbidden' }, cb);
        var blob = file.getBlob();
        return _json({ ok:true, mime:blob.getContentType(), name:file.getName(), dataB64:Utilities.base64Encode(blob.getBytes()) }, cb);
      } catch (ef) { return _json({ ok:false, error:'not_found' }, cb); }
    }
    // v3.8: la consulta publica ?usuario= se elimino (exponia puntos sin login
    // y escribia en la hoja). Las apps usan action=me con id_token.
    if (e && e.parameter && e.parameter.usuario) return _json({ ok:false, error:'removed' }, cb);
    return _json({ ok:true, servicio:'ClimateGuard receptor', version:'3.8', sheet: !!sheetId }, cb);
  } catch (err) {
    return _json({ ok:false, error:String(err) }, cb);
  }
}

// ── helpers ─────────────────────────────────────────────────────────────────
function _valorActual(sh, id, col) {
  var last = sh.getLastRow(); if (last < 2 || !id) return '';
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for (var i = 0; i < ids.length; i++) { if (String(ids[i][0]) === id) return sh.getRange(i+2, col, 1, 1).getValue(); }
  return '';
}
function _sheet(ss, nombre, header) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); if (header) sh.appendRow(header); }
  else if (header && sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}
/* Reinicia una cuenta: borra sus filas en las pestanas y envia su carpeta a la
   papelera. delAcct=true tambien borra la fila de Usuarios (registro). */
function _resetUsuarioServer(ss, email, delAcct){
  email = String(email).toLowerCase();
  var plan = [ {name:'Observaciones',col0:1}, {name:'Recorridos',col0:1}, {name:'Puntos',col0:0}, {name:'Asistencia',col0:1} ];
  if (delAcct) plan.push({ name:'Usuarios', col0:0 });
  var rep = {};
  for (var pi=0; pi<plan.length; pi++){
    var p = plan[pi];
    var sh = ss.getSheetByName(p.name); if(!sh){ rep[p.name]=0; continue; }
    var last = sh.getLastRow(); if(last<2){ rep[p.name]=0; continue; }
    var vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    var rows = [];
    for (var i=0;i<vals.length;i++){ if(String(vals[i][p.col0]||'').toLowerCase()===email) rows.push(i+2); }
    rep[p.name] = rows.length;
    for (var k=rows.length-1;k>=0;k--){ sh.deleteRow(rows[k]); }
  }
  var folders=0, files=0;
  try{
    var raiz = DriveApp.getFolderById(CARPETA_ID);
    var it = raiz.getFoldersByName(email);
    while(it.hasNext()){ var f=it.next(); folders++; var fi=f.getFiles(); while(fi.hasNext()){ fi.next(); files++; } f.setTrashed(true); }
  }catch(eD){}
  rep.DriveFolder = folders; rep.DriveFiles = files;
  return rep;
}
function _upsert(sh, id, fila) {
  var last = sh.getLastRow();
  if (last >= 2 && id) {
    var ids = sh.getRange(2,1,last-1,1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) { sh.getRange(i+2,1,1,fila.length).setValues([fila]); return; }
    }
  }
  sh.appendRow(fila);
}
function _certificarPuntos(shO, ss, usuario) {
  var last = shO.getLastRow();
  var pts = 0, cnt = 0; usuario = String(usuario).toLowerCase();
  if (last >= 2) {
    var w = shO.getLastColumn();
    var vals = shO.getRange(2,1,last-1,w).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][1]).toLowerCase() !== usuario) continue;   // col B = usuario
      var estado = String(vals[i][15]||'').toLowerCase();           // col P = estado
      if (estado.indexOf('invalid')===0 || estado.indexOf('falso')===0 || estado==='demo') continue;
      if (_esDemo(vals[i][0])) continue;                            // v3.8: demo no suma
      var lat = vals[i][7], especie = vals[i][2];
      var base = 10 + (lat!=='' && lat!=null ? 5:0) + (especie ? 5:0);
      pts += base * BONO; cnt += 1;
    }
  }
  // Recorridos (tracks): 10 + 5 si distancia >= 200 m (igual que la app)
  try {
    var shR = _sheet(ss, 'Recorridos', HEAD_REC);
    var lastR = shR.getLastRow();
    if (lastR >= 2) {
      var vr = shR.getRange(2,1,lastR-1,shR.getLastColumn()).getValues();
      for (var r = 0; r < vr.length; r++) {
        if (String(vr[r][1]).toLowerCase() !== usuario) continue;   // col B = usuario
        if (_esDemo(vr[r][0])) continue;
        var dm = Number(vr[r][4] || 0);                             // col E = distance_m
        pts += (10 + (dm >= 200 ? 5 : 0)) * BONO;
      }
    }
  } catch (eR) {}
  var nivel = 1;
  for (var j = 0; j < NIVELES_MIN.length; j++) if (pts >= NIVELES_MIN[j]) nivel = j+1;
  var shP = _sheet(ss, 'Puntos', HEAD_PTS);
  _upsert(shP, usuario, [usuario, pts, nivel, cnt, new Date().toISOString()]);
  return { puntos: pts, nivel: nivel };
}
function _upsertFile(folder, name, blob, prefijo){
  var it = folder.getFilesByName(name);
  while (it.hasNext()) { try { it.next().setTrashed(true); } catch(_){} }
  if (prefijo) {   // v3.8: borra tambien la version con otra extension del mismo registro
    var it2 = folder.searchFiles('title contains "' + String(prefijo).replace(/"/g,'') + '" and trashed = false');
    while (it2.hasNext()) { var f2 = it2.next(); if (f2.getName().indexOf(prefijo) === 0 && f2.getName() !== name) { try { f2.setTrashed(true); } catch(_){} } }
  }
  return folder.createFile(blob.setName(name));
}
/* v3.8: una sola lectura de la hoja -> indice por id (fila, usuario, valores). */
function _indice(sh){
  var last = sh.getLastRow(), w = Math.max(sh.getLastColumn(), 1);
  var idx = { byId:{}, rows:[], last:last };
  if (last < 2) { idx.last = Math.max(last, 1); return idx; }
  var vals = sh.getRange(2, 1, last-1, w).getValues();
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || ''); idx.rows.push(vals[i]);
    if (id && !idx.byId[id]) idx.byId[id] = { r:i+2, usuario:String(vals[i][1]||'').toLowerCase(), row:vals[i] };
  }
  return idx;
}
function _upsertIdx(sh, idx, id, usuario, fila){
  var prev = idx.byId[id];
  if (prev) { sh.getRange(prev.r, 1, 1, fila.length).setValues([fila]); prev.row = fila; prev.usuario = usuario; return; }
  sh.appendRow(fila);
  idx.last += 1; idx.byId[id] = { r:idx.last, usuario:usuario, row:fila }; idx.rows.push(fila);
}
function _rateOkIdx(ix, email){
  var now = Date.now(), h = 0, d = 0;
  for (var i = 0; i < ix.rows.length; i++) {
    if (String(ix.rows[i][1]).toLowerCase() !== email) continue;
    var v = ix.rows[i][14]; var t = v ? (v instanceof Date ? v.getTime() : Date.parse(v)) : 0; if (!t) continue;
    if (now - t < 3600e3) h++;
    if (now - t < 86400e3) d++;
  }
  return h < RATE_PER_HOUR && d < RATE_PER_DAY;
}
/* v3.8: ejecutar UNA vez (lo llama actualizarEsquema). Marca las filas de
   demostracion (id demo-*) con estado 'demo' y recalcula los puntos de los
   usuarios afectados: dejan de sumar y de verse en la web. */
function marcarDemo(){
  var id = getSheetId(); if (!id) return;
  var ss = SpreadsheetApp.openById(id);
  var shO = _sheet(ss, 'Observaciones', HEAD_OBS);
  var last = shO.getLastRow(); var afect = {}; var n = 0;
  if (last >= 2) {
    var v = shO.getRange(2, 1, last-1, COL_OBS_ESTADO).getValues();
    for (var i = 0; i < v.length; i++) if (_esDemo(v[i][0])) { shO.getRange(i+2, COL_OBS_ESTADO).setValue('demo'); afect[String(v[i][1]).toLowerCase()] = 1; n++; }
  }
  var shR = _sheet(ss, 'Recorridos', HEAD_REC); var lr = shR.getLastRow();
  if (lr >= 2) { var vr = shR.getRange(2, 1, lr-1, 2).getValues(); for (var k = 0; k < vr.length; k++) if (_esDemo(vr[k][0])) afect[String(vr[k][1]).toLowerCase()] = 1; }
  Object.keys(afect).forEach(function(u){ if (u) _certificarPuntos(shO, ss, u); });
  Logger.log('marcarDemo: filas demo=' + n + ', usuarios recalculados=' + Object.keys(afect).length);
}
function _carpetaUsuario(ss, email){
  email = String(email).toLowerCase();
  var raiz = DriveApp.getFolderById(CARPETA_ID);
  if (!ss) { var it0 = raiz.getFoldersByName(email); return it0.hasNext()? it0.next() : raiz.createFolder(email); }
  var sh = _usuarios(ss); var COL_FID = 10; var row = _userRow(sh, email);
  if (row > 0) { var fid = String(sh.getRange(row, COL_FID).getValue()||''); if (fid) { try { var f = DriveApp.getFolderById(fid); if (f && !f.isTrashed()) return f; } catch(_){} } }
  var lock = LockService.getScriptLock(); try { lock.waitLock(8000); } catch(_){}
  try {
    row = _userRow(sh, email);
    if (row > 0) { var fid2 = String(sh.getRange(row, COL_FID).getValue()||''); if (fid2) { try { var f2 = DriveApp.getFolderById(fid2); if (f2 && !f2.isTrashed()) return f2; } catch(_){} } }
    var it = raiz.getFoldersByName(email); var folder = it.hasNext()? it.next() : raiz.createFolder(email);
    if (row > 0) sh.getRange(row, COL_FID).setValue(folder.getId());
    return folder;
  } finally { try { lock.releaseLock(); } catch(_){} }
}
function _setUsuarioAvatar(ss, email, url){ var sh=_usuarios(ss); var row=_userRow(sh,String(email).toLowerCase()); if(row>0) sh.getRange(row,11).setValue(url); }
function _usuarioAvatar(ss, email){ var sh=_usuarios(ss); var row=_userRow(sh,String(email).toLowerCase()); return row>0? String(sh.getRange(row,11).getValue()||'') : ''; }
function _rowsFor(sh, usuarioCol0, email, header){
  var last=sh.getLastRow(); if(last<2) return [];
  var w=sh.getLastColumn(); var vals=sh.getRange(2,1,last-1,w).getValues();
  email=String(email).toLowerCase(); var out=[];
  for(var i=0;i<vals.length;i++){ if(String(vals[i][usuarioCol0]).toLowerCase()!==email) continue;
    if(_esDemo(vals[i][0])) continue;   // v3.8: sin datos de demostracion
    var o={}; for(var j=0;j<header.length;j++){ o[header[j]]=vals[i][j]; } out.push(o); }
  return out;
}
/* Ejecutar UNA vez para dejar la base lista para el sitio web:
   actualiza encabezados, BORRA todos los datos y VACIA la carpeta de Drive
   (mantiene la propia hoja de calculo). */
function resetYActualizar(){
  var id=getSheetId(); if(!id){ Logger.log('Ejecuta primero crearBaseDeDatos()'); return; }
  var ss=SpreadsheetApp.openById(id);
  _setHead(ss,'Observaciones',HEAD_OBS); _setHead(ss,'Recorridos',HEAD_REC); _setHead(ss,'Puntos',HEAD_PTS);
  _setHead(ss,'Usuarios',HEAD_USR); _setHead(ss,'Asistencia',HEAD_ASIST);
  ['Observaciones','Recorridos','Puntos','Usuarios','Asistencia'].forEach(function(n){
    var sh=ss.getSheetByName(n); if(!sh) return; var last=sh.getLastRow(); if(last>1) sh.deleteRows(2,last-1);
  });
  var raiz=DriveApp.getFolderById(CARPETA_ID); var nf=0, nfi=0;
  var fol=raiz.getFolders(); while(fol.hasNext()){ try{ fol.next().setTrashed(true); nf++; }catch(_){} }
  var fil=raiz.getFiles(); while(fil.hasNext()){ var fx=fil.next(); if(fx.getId()===id) continue; try{ fx.setTrashed(true); nfi++; }catch(_){} }
  Logger.log('reset ok · datos borrados · carpetas '+nf+' · archivos '+nfi);
}
function _setHead(ss, name, head){ var sh=ss.getSheetByName(name)||ss.insertSheet(name); sh.getRange(1,1,1,head.length).setValues([head]); }
function _json(obj, callback) {
  var txt = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + txt + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}