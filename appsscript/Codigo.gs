/**
 * ClimateGuard - Receptor de datos + Autenticacion Google (Apps Script Web App)
 * ===========================================================================
 * v2 (login real): en lugar de un TOKEN publico compartido, cada envio debe
 * traer un ID token de Google (OpenID Connect). El servidor lo verifica con
 * el endpoint tokeninfo de Google, obtiene el email VERIFICADO del usuario y
 * solo entonces guarda los datos. Ademas mantiene un padron "Usuarios" con
 * estado por usuario (active / blocked / pending) y limita la tasa de envios.
 *
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

// ── TRANSICION (despliegue v1.5.0) ─────────────────────────────────────────
// Mientras queden usuarios con la app vieja (v1.4.7) en cache, se admite el
// token compartido anterior para que sus envios NO fallen durante el cambio.
// Esos envios se guardan pero NO pasan por el padron ni el limite de tasa.
// >>> Cuando TODOS usen v1.5.0, pon esto en false para seguridad total. <<<
const ACEPTAR_TOKEN_LEGADO = false;
const TOKEN_LEGADO = 'cg-kLVvRhNatYBkOTphQmCOx585YIiM';

const BONO        = 1;                    // los puntos se "reconocen" al enviar
const NIVELES_MIN = [0, 100, 250, 500, 1000];

const HEAD_OBS = ['id','usuario','especie','clase','iucn','cant','notas','lat','lng','accuracy','fecha','updated_at','foto','submissionId','recibido','estado','altitude','alt_accuracy'];
const HEAD_REC = ['id','usuario','name','points','distance_m','start','end','submissionId','recibido'];
const HEAD_PTS = ['usuario','puntos','nivel','observaciones','actualizado'];
const HEAD_USR = ['email','name','picture','status','role','sub','created_at','last_login','submissions'];
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
  _sheet(ss,'Usuarios',HEAD_USR);
  _sheet(ss,'Asistencia',HEAD_ASIST);
  Logger.log('Esquema actualizado (Usuarios + Asistencia + encabezados).');
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
    sh.appendRow([u.email, u.name, u.picture, st, '', u.sub, now, now, 0]);
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
  if (!/^[0-9a-f]{64}$/.test(hash)) return null;
  var sh = _usuarios(ss); var last = sh.getLastRow(); if (last < 2) return null;
  var vals = sh.getRange(2,1,last-1,2).getValues();   // email, name
  for (var i=0;i<vals.length;i++){
    var email = String(vals[i][0]||'').toLowerCase().trim(); if (!email) continue;
    if (_sha256hex(email) === hash) return { email: email, name: (String(vals[i][1]||'') || email) };
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

    // 1) Autenticacion: ID token de Google (preferente). Durante la transicion
    //    se acepta tambien el token compartido de la app vieja (v1.4.7).
    var user = verifyIdToken(body.idToken);
    var usuario;
    if (user && user.email) {
      usuario = user.email;
    } else if (ACEPTAR_TOKEN_LEGADO && body.token === TOKEN_LEGADO) {
      usuario = String(body.usuario || 'anonimo');
      user = null;   // cliente viejo: se salta padron y limite de tasa
    } else {
      return _json({ ok:false, error:'auth' });
    }
    var sid     = String(body.submissionId || '');

    var sheetId = getSheetId();
    var ss = sheetId ? SpreadsheetApp.openById(sheetId) : null;

    // 2) Padron: registrar/actualizar y comprobar estado (solo login Google)
    if (ss && user) {
      var st = String(_touchUser(ss, user)).toLowerCase();
      if (st === 'blocked') return _json({ ok:false, error:'blocked' });
      if (st === 'pending') return _json({ ok:false, error:'pending' });
      // 3) Limite de tasa
      var shOrate = _sheet(ss, 'Observaciones', HEAD_OBS);
      if (!_rateOk(shOrate, usuario)) return _json({ ok:false, error:'rate' });
    }

    // 4) Archivos en Drive (carpeta por envio)
    var raiz = DriveApp.getFolderById(CARPETA_ID);
    var sello = String(body.fecha || new Date().toISOString()).replace(/[:.]/g,'-');
    var sub = raiz.createFolder(sello + ' - ' + usuario.replace(/[^\w\-@. ]+/g,'_'));
    if (body.observaciones) sub.createFile('observaciones.geojson', JSON.stringify(body.observaciones), 'application/geo+json');
    if (body.recorridos)    sub.createFile('recorridos.geojson',    JSON.stringify(body.recorridos),    'application/geo+json');
    if (body.fotos && body.fotos.length) {
      for (var i = 0; i < body.fotos.length; i++) {
        var f = body.fotos[i];
        var blob = Utilities.newBlob(Utilities.base64Decode(f.dataB64), f.mime || 'image/jpeg', f.name || ('foto_'+(i+1)+'.jpg'));
        sub.createFile(blob);
      }
    }

    // 5) Base de datos en Sheets (upsert por id) — usuario = email verificado
    var puntosCert = null, nivelCert = null;
    if (ss) {
      var obsF = (body.observaciones && body.observaciones.features) || [];
      var recF = (body.recorridos && body.recorridos.features) || [];

      var shO = _sheet(ss, 'Observaciones', HEAD_OBS);
      obsF.forEach(function(ft){
        var p = ft.properties || {}, g = ft.geometry || {}, c = (g.coordinates||[null,null]);
        var estado = _valorActual(shO, String(p.id||''), 16) || 'valido';
        _upsert(shO, String(p.id||''), [String(p.id||''), usuario, p.especie||'', p.clase||'', p.iucn||'',
          p.cant||1, p.notas||'', c[1], c[0], (p.accuracy==null?'':p.accuracy), p.fecha||'', p.updated_at||'',
          p.foto||'', sid, new Date().toISOString(), estado,
          (p.altitude==null?'':p.altitude), (p.alt_accuracy==null?'':p.alt_accuracy)]);
      });

      var shR = _sheet(ss, 'Recorridos', HEAD_REC);
      recF.forEach(function(ft){
        var p = ft.properties || {};
        _upsert(shR, String(p.id||''), [String(p.id||''), usuario, p.name||'', p.points||'', p.distance_m||'',
          p.start||'', p.end||'', sid, new Date().toISOString()]);
      });

      var r = _certificarPuntos(shO, ss, usuario);
      puntosCert = r.puntos; nivelCert = r.nivel;
      _bumpSubmissions(ss, usuario);
    }

    return _json({ ok:true, usuario:usuario, carpeta:sub.getName(), url:sub.getUrl(), puntos:puntosCert, nivel:nivelCert, sheet: !!sheetId });
  } catch (err) {
    return _json({ ok:false, error:String(err) });
  }
}

// ── GET: login/estado (JSONP), consulta de puntos y asistencia ──────────────
function doGet(e) {
  var cb = e && e.parameter && e.parameter.callback;
  try {
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
      return _json({ ok:true, email:user.email, name:user.name, picture:user.picture, status:st, puntos:puntos, nivel:nivel }, cb);
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
      var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); // dia del servidor
      var dup = _marcarAsistencia(ss, fecha, miembro.email, miembro.name, admin.email);
      return _json({ ok:true, name:miembro.name, email:miembro.email, dup:dup, date:fecha }, cb);
    }

    // Consulta de puntos por email (?usuario=email[&callback=fn])
    var u = e && e.parameter && e.parameter.usuario;
    if (u && ss) {
      var r = _certificarPuntos(_sheet(ss,'Observaciones',HEAD_OBS), ss, String(u).toLowerCase());
      return _json({ ok:true, usuario:u, puntos:r.puntos, nivel:r.nivel }, cb);
    }
    return _json({ ok:true, servicio:'ClimateGuard receptor', sheet: !!sheetId }, cb);
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
      if (estado.indexOf('invalid')===0 || estado.indexOf('falso')===0) continue;
      var lat = vals[i][7], especie = vals[i][2];
      var base = 10 + (lat!=='' && lat!=null ? 5:0) + (especie ? 5:0);
      pts += base * BONO; cnt += 1;
    }
  }
  var nivel = 1;
  for (var j = 0; j < NIVELES_MIN.length; j++) if (pts >= NIVELES_MIN[j]) nivel = j+1;
  var shP = _sheet(ss, 'Puntos', HEAD_PTS);
  _upsert(shP, usuario, [usuario, pts, nivel, cnt, new Date().toISOString()]);
  return { puntos: pts, nivel: nivel };
}
function _json(obj, callback) {
  var txt = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + txt + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
