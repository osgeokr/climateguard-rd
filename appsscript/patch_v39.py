"""Aplica los 3 ganchos de la sesion propia (v3.9) a un Codigo.gs existente.
Uso: python3 patch_v39.py Codigo.gs   (modifica el archivo en su lugar)"""
import sys
P = sys.argv[1]
s = open(P, encoding='utf-8').read()
if 'sesionGet(' in s:
    sys.exit('ya aplicado')
def rep(old, new):
    global s
    n = s.count(old)
    if n != 1: sys.exit(f'No encontrado de forma unica ({n}): {old!r}')
    s = s.replace(old, new)
rep("  if (!idToken) return null;\n",
    "  if (!idToken) return null;\n"
    "  if (String(idToken).indexOf('cgs1.') === 0) return verifySession(idToken);   // v3.9: sesion propia (Sesion.gs)\n")
rep("    var body = JSON.parse(e.postData.contents);\n",
    "    var body = JSON.parse(e.postData.contents);\n"
    "    var _sesP = sesionPost(body); if (_sesP) return _sesP;   // v3.9: sesion propia (Sesion.gs)\n")
rep("  var cb = e && e.parameter && e.parameter.callback;\n  try {\n",
    "  var cb = e && e.parameter && e.parameter.callback;\n  try {\n"
    "    var _ses = sesionGet(e, cb); if (_ses) return _ses;   // v3.9: login/avatar/logout + ?session= (Sesion.gs)\n")
s = s.replace(" * v3.7 (admin):", " * v3.9 (sesion): clave de sesion propia sin caducidad, action=login/avatar/\n"
              " *   logoutuser/logoutall. Codigo en Sesion.gs (archivo aparte del proyecto).\n * v3.7 (admin):", 1)
open(P, 'w', encoding='utf-8').write(s)
print('OK')
