# -*- coding: utf-8 -*-
"""Usuarios, sesiones y permisos del CRM.

Los usuarios viven en ``DATA_DIR/usuarios.json`` (el volumen persistente de
Railway). Nunca se guarda la contraseña: sólo un hash PBKDF2-SHA256 con sal
propia por usuario.

El primer arranque siembra los usuarios que vengan en la variable de entorno
``SEMILLA_USUARIOS`` (JSON con el hash ya calculado), para no dejar ninguna
credencial escrita en el repositorio.

Variables de entorno:
  DATA_DIR          carpeta persistente (default: ./data)
  SEMILLA_USUARIOS  JSON [{"usuario","rol","hash","nombre"}] para el alta inicial
  SESION_SECRETO    secreto para firmar la cookie de sesión (si falta se genera
                    y se guarda en DATA_DIR/.sesion_secreto)
  SESION_HORAS      duración de la sesión en horas (default: 12)
"""
import base64
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
from datetime import datetime

DATA_DIR = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data'))
ARCHIVO = os.path.join(DATA_DIR, 'usuarios.json')
ARCHIVO_SECRETO = os.path.join(DATA_DIR, '.sesion_secreto')
# Sin "recordarme" la sesión dura lo que dure el navegador abierto (y como
# mucho SESION_HORAS_CORTA); con "recordarme" aguanta la jornada entera.
SESION_HORAS_CORTA = int(os.environ.get('SESION_HORAS_CORTA', '2'))
SESION_HORAS = int(os.environ.get('SESION_HORAS', '12'))

_ITERACIONES = 240_000
_lock = threading.RLock()

# Intentos fallidos por usuario: frena el probar contraseñas a lo bruto.
_MAX_INTENTOS = 10
_ESPERA_BLOQUEO = 300  # segundos
_intentos = {}


# ============================================================
# Roles y permisos
# ============================================================
# Secciones de la interfaz (coinciden con data-pagina en index.html).
SECCIONES = ['panel', 'reportes', 'agendados', 'venta', 'caja', 'inventario',
             'planilla', 'hoy', 'calendario', 'pacientes', 'remarketing',
             'analitica', 'metaads', 'usuarios']

# Qué recurso de la API cubre cada prefijo de ruta. El permiso se decide por
# recurso, no por sección: varias pantallas comparten los mismos endpoints
# (Calendario, por ejemplo, lee de AGENDADOS).
RECURSOS = {
    # Lo que necesita cualquiera que haya entrado, sea cual sea su rol.
    'comun': ['/api/estado', '/api/yo', '/api/logout', '/api/asistente'],
    'agendados': ['/api/crm/agendados'],
    'venta': ['/api/crm/venta'],
    'hoy': ['/api/crm/hoy', '/api/crm/tareas'],
    'caja': ['/api/crm/caja', '/api/crm/cuotas'],
    'inventario': ['/api/crm/inventario'],
    'planilla': ['/api/crm/planilla'],
    'pipeline': ['/api/crm/pipeline'],
    # Las notas se ven dentro de la ficha del paciente, no sólo en el pipeline.
    'notas': ['/api/crm/notas'],
    'pacientes': ['/api/crm/pacientes', '/api/pacientes'],
    'remarketing': ['/api/crm/remarketing'],
    # '/api/crm/pacientes/historias' (subir/listar/descargar/borrar el PDF de
    # la ficha clínica) vive bajo el prefijo de 'pacientes', pero es el
    # recurso 'historias': se declara aparte, más específico, para que
    # recurso_de() (que elige el prefijo más largo) no lo confunda con el
    # directorio de pacientes — si no, un rol sin 'historias' pero con
    # 'pacientes' (p. ej. CRM1) podía leer y descargar historias clínicas.
    'historias': ['/api/crm/historias', '/api/crm/pacientes/historias'],
    'dashboard': ['/api/crm/dashboard'],
    'analitica': ['/api/analitica'],
    'metaads': ['/api/meta', '/api/meta-mensual'],
    'reportes': ['/api/reporte', '/api/reportes', '/api/exportar', '/api/sync',
                 '/api/maestro', '/api/backup', '/api/backups'],
    'whatsapp': ['/api/whatsapp/campanas'],
    'usuarios': ['/api/usuarios'],
}

# Rutas que no piden sesión. El webhook de WhatsApp lo llama Meta, no el
# navegador, y ya valida la firma HMAC por su cuenta.
#
# OJO: /api/debug/* NO va aquí. Son endpoints de diagnóstico que devuelven
# volcados crudos de VENTA/AGENDADOS/maestro (nombres, teléfonos, montos) —
# estuvieron públicos por error (cualquiera en internet podía leerlos sin
# loguearse) hasta que se encontró en una auditoría. Al no estar en RECURSOS
# ni en PUBLICAS, "puede()" los niega por defecto a todos salvo ADMIN.
PUBLICAS = ['/login', '/api/login', '/api/whatsapp/webhook', '/static',
            '/favicon.ico', '/salud']

ROLES = {
    'ADMIN': {
        'nombre': 'Administrador',
        'descripcion': 'Acceso total: todas las secciones y todos los cambios.',
        'secciones': list(SECCIONES),
        'lectura': '*',
        'escritura': '*',
        'denegar': [],
    },
    'DOCTOR': {
        'nombre': 'Médico',
        'descripcion': ('Panel, Hoy, Calendario y Pacientes, con sus historias '
                        'clínicas. No ve ventas, caja, planilla ni reportes.'),
        'secciones': ['panel', 'hoy', 'calendario', 'pacientes'],
        # 'agendados' no es una sección suya, pero Hoy y Calendario leen y
        # marcan citas ahí: sin ese permiso no podría atender su propio día.
        'lectura': ['comun', 'hoy', 'agendados', 'pacientes', 'historias', 'notas'],
        'escritura': ['comun', 'hoy', 'agendados', 'pacientes', 'historias', 'notas'],
        # Igual que CRM: "compró" escribe en VENTA DIARIA y el médico no
        # registra ventas.
        'denegar': [('POST', '/api/crm/agendados/*/compro')],
    },
    'CRM': {
        'nombre': 'CRM',
        'descripcion': ('Panel, Agendados, Venta diaria (sólo lectura), Hoy, '
                        'Calendario y Re llamadas. No puede registrar ni editar '
                        'ventas.'),
        'secciones': ['panel', 'agendados', 'venta', 'hoy', 'calendario', 'remarketing'],
        # Historias entra porque el flujo de "asistió" en Hoy abre la historia
        # clínica; sin esto ese botón se rompería.
        'lectura': ['comun', 'agendados', 'venta', 'hoy', 'historias', 'remarketing'],
        # 'comun' incluye /api/yo/clave: todos pueden cambiar su contraseña.
        'escritura': ['comun', 'agendados', 'hoy', 'historias', 'remarketing'],
        # "Compró" crea filas en VENTA DIARIA aunque cuelgue de /agendados:
        # se niega aparte o sería una puerta trasera para registrar ventas.
        'denegar': [('POST', '/api/crm/agendados/*/compro')],
    },
    'CRM1': {
        'nombre': 'CRM (con Pacientes)',
        'descripcion': ('Panel, Agendados, Venta diaria (sólo lectura), Citas de '
                        'hoy, Calendario y Pacientes (sólo lectura, sin historias '
                        'clínicas ni ficha de detalle).'),
        'secciones': ['panel', 'agendados', 'venta', 'hoy', 'calendario', 'pacientes'],
        # Sin 'historias' ni 'notas': no debe ver ni editar la ficha clínica del
        # paciente, sólo el listado. Sin 'analitica': así el bloque de "Buscar
        # historial"/"reactivar" (que cuelga de esa sección) queda oculto.
        'lectura': ['comun', 'agendados', 'venta', 'hoy', 'pacientes'],
        'escritura': ['comun', 'agendados', 'hoy'],
        'denegar': [('POST', '/api/crm/agendados/*/compro')],
    },
}
ROL_POR_DEFECTO = 'CRM'


def _coincide(ruta, prefijo):
    """Compara por segmentos: /api/crm/venta no debe capturar /api/crm/ventax."""
    return ruta == prefijo or ruta.startswith(prefijo + '/')


def _coincide_patron(ruta, patron):
    """Compara una ruta contra un patrón con '*' como comodín de un segmento."""
    a, b = ruta.strip('/').split('/'), patron.strip('/').split('/')
    if len(a) != len(b):
        return False
    return all(y == '*' or x == y for x, y in zip(a, b))


def recurso_de(ruta):
    """Recurso al que pertenece una ruta, o None si no está mapeada."""
    mejor, largo = None, -1
    for recurso, prefijos in RECURSOS.items():
        for p in prefijos:
            if _coincide(ruta, p) and len(p) > largo:
                mejor, largo = recurso, len(p)
    return mejor


def es_publica(ruta):
    return any(_coincide(ruta, p) for p in PUBLICAS)


def puede(rol, metodo, ruta):
    """¿El rol puede ejecutar ``metodo`` sobre ``ruta``?

    Lo que no está mapeado se niega salvo para ADMIN: así, si mañana se agrega
    un endpoint y se olvida ponerlo en RECURSOS, queda cerrado y no abierto."""
    cfg = ROLES.get(rol)
    if not cfg:
        return False
    for m, patron in cfg['denegar']:
        if m == metodo and _coincide_patron(ruta, patron):
            return False
    recurso = recurso_de(ruta)
    escribe = metodo not in ('GET', 'HEAD', 'OPTIONS')
    permitidos = cfg['escritura'] if escribe else cfg['lectura']
    if permitidos == '*':
        return True
    if recurso is None:
        return False
    if not escribe and recurso in cfg['escritura']:
        return True
    return recurso in permitidos


def perfil_publico(u):
    """Datos del usuario que sí puede ver el navegador (nunca el hash)."""
    cfg = ROLES.get(u['rol'], ROLES[ROL_POR_DEFECTO])
    return {
        'usuario': u['usuario'],
        'nombre': u.get('nombre') or u['usuario'],
        'rol': u['rol'],
        'rol_nombre': cfg['nombre'],
        'secciones': cfg['secciones'],
        'lectura': cfg['lectura'],
        'escritura': cfg['escritura'],
        'activo': u.get('activo', True),
        'creado': u.get('creado'),
        'ultimo_ingreso': u.get('ultimo_ingreso'),
    }


# ============================================================
# Contraseñas
# ============================================================
def hash_clave(clave, iteraciones=_ITERACIONES):
    if not clave:
        raise ValueError('La contraseña no puede estar vacía')
    sal = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac('sha256', str(clave).encode(), sal, iteraciones)
    return 'pbkdf2_sha256${}${}${}'.format(
        iteraciones, base64.b64encode(sal).decode(), base64.b64encode(dk).decode())


_falso = None


def _hash_falso():
    """Hash de descarte, con el mismo coste que uno real, para comparar contra
    él cuando el usuario no existe y no delatar por tiempo qué usuarios hay."""
    global _falso
    if _falso is None:
        _falso = hash_clave(secrets.token_hex(16))
    return _falso


def verificar_clave(clave, guardado):
    try:
        algo, iteraciones, sal, esperado = str(guardado).split('$')
        if algo != 'pbkdf2_sha256':
            return False
        dk = hashlib.pbkdf2_hmac('sha256', str(clave or '').encode(),
                                 base64.b64decode(sal), int(iteraciones))
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(base64.b64encode(dk).decode(), esperado)


# ============================================================
# Almacén
# ============================================================
def _ahora():
    return datetime.now().strftime('%d/%m/%Y %H:%M')


def _leer():
    if not os.path.exists(ARCHIVO):
        return []
    try:
        with open(ARCHIVO, encoding='utf-8') as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []
    return datos if isinstance(datos, list) else []


def _escribir(usuarios):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = ARCHIVO + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ARCHIVO)
    try:
        os.chmod(ARCHIVO, 0o600)
    except OSError:
        pass


def normalizar_usuario(usuario):
    """Los nombres de usuario no distinguen mayúsculas ni espacios sobrantes."""
    return str(usuario or '').strip().upper()


_clave_usuario = normalizar_usuario


def obtener(usuario):
    nombre = _clave_usuario(usuario)
    for u in _leer():
        if _clave_usuario(u.get('usuario')) == nombre:
            return u
    return None


def listar():
    return [perfil_publico(u) for u in
            sorted(_leer(), key=lambda x: _clave_usuario(x.get('usuario')))]


def hay_usuarios():
    return bool(_leer())


def crear(usuario, clave, rol, nombre=''):
    usuario = _clave_usuario(usuario)
    if not usuario:
        raise ValueError('Indica el usuario')
    if rol not in ROLES:
        raise ValueError(f'Rol inválido: {rol}')
    if len(str(clave or '')) < 4:
        raise ValueError('La contraseña debe tener al menos 4 caracteres')
    with _lock:
        usuarios = _leer()
        if any(_clave_usuario(u.get('usuario')) == usuario for u in usuarios):
            raise ValueError(f'El usuario {usuario} ya existe')
        nuevo = {'usuario': usuario, 'nombre': str(nombre or '').strip(),
                 'rol': rol, 'hash': hash_clave(clave), 'activo': True,
                 'creado': _ahora(), 'ultimo_ingreso': None}
        usuarios.append(nuevo)
        _escribir(usuarios)
        return perfil_publico(nuevo)


def actualizar(usuario, cambios):
    usuario = _clave_usuario(usuario)
    with _lock:
        usuarios = _leer()
        for u in usuarios:
            if _clave_usuario(u.get('usuario')) != usuario:
                continue
            if cambios.get('rol'):
                if cambios['rol'] not in ROLES:
                    raise ValueError(f"Rol inválido: {cambios['rol']}")
                u['rol'] = cambios['rol']
            if 'nombre' in cambios:
                u['nombre'] = str(cambios.get('nombre') or '').strip()
            if 'activo' in cambios and cambios['activo'] is not None:
                u['activo'] = bool(cambios['activo'])
            if cambios.get('clave'):
                if len(str(cambios['clave'])) < 4:
                    raise ValueError('La contraseña debe tener al menos 4 caracteres')
                u['hash'] = hash_clave(cambios['clave'])
            _escribir(usuarios)
            _intentos.pop(usuario, None)
            return perfil_publico(u)
    return None


def borrar(usuario):
    usuario = _clave_usuario(usuario)
    with _lock:
        usuarios = _leer()
        quedan = [u for u in usuarios if _clave_usuario(u.get('usuario')) != usuario]
        if len(quedan) == len(usuarios):
            return False
        # Nunca dejar el sistema sin un administrador que pueda entrar.
        if not any(u['rol'] == 'ADMIN' and u.get('activo', True) for u in quedan):
            raise ValueError('No puedes borrar al último administrador activo')
        _escribir(quedan)
        return True


def sembrar():
    """Alta inicial desde SEMILLA_USUARIOS. Sólo crea los que faltan: nunca
    pisa una contraseña que el administrador ya haya cambiado."""
    crudo = os.environ.get('SEMILLA_USUARIOS', '').strip()
    if not crudo:
        return 0
    try:
        semilla = json.loads(crudo)
    except json.JSONDecodeError:
        print('[usuarios] SEMILLA_USUARIOS no es JSON válido, se ignora')
        return 0
    creados = 0
    with _lock:
        usuarios = _leer()
        existentes = {_clave_usuario(u.get('usuario')) for u in usuarios}
        for s in semilla if isinstance(semilla, list) else []:
            nombre = _clave_usuario(s.get('usuario'))
            if not nombre or nombre in existentes or not s.get('hash'):
                continue
            usuarios.append({'usuario': nombre, 'nombre': s.get('nombre') or nombre,
                             'rol': s.get('rol') if s.get('rol') in ROLES else ROL_POR_DEFECTO,
                             'hash': s['hash'], 'activo': True,
                             'creado': _ahora(), 'ultimo_ingreso': None})
            existentes.add(nombre)
            creados += 1
        if creados:
            _escribir(usuarios)
    return creados


def autenticar(usuario, clave):
    """Devuelve el perfil si las credenciales son correctas.

    Lanza ValueError con el motivo si no lo son (usuario inexistente, inactivo,
    bloqueado por intentos fallidos o contraseña equivocada)."""
    nombre = _clave_usuario(usuario)
    bloqueo = _intentos.get(nombre)
    if bloqueo and bloqueo['n'] >= _MAX_INTENTOS:
        restan = int(bloqueo['hasta'] - time.time())
        if restan > 0:
            raise ValueError(f'Demasiados intentos fallidos. Espera {restan // 60 + 1} min.')
        _intentos.pop(nombre, None)

    u = obtener(nombre)
    # Se compara igual contra un hash de descarte cuando el usuario no existe,
    # para que el tiempo de respuesta no delate qué usuarios son válidos.
    ok = verificar_clave(clave, u['hash'] if u else _hash_falso())

    def fallo(motivo):
        d = _intentos.setdefault(nombre, {'n': 0, 'hasta': 0})
        d['n'] += 1
        d['hasta'] = time.time() + _ESPERA_BLOQUEO
        raise ValueError(motivo)

    if not u or not ok:
        fallo('Usuario o contraseña incorrectos')
    if not u.get('activo', True):
        fallo('Ese usuario está desactivado')

    _intentos.pop(nombre, None)
    with _lock:
        usuarios = _leer()
        for x in usuarios:
            if _clave_usuario(x.get('usuario')) == nombre:
                x['ultimo_ingreso'] = _ahora()
                break
        _escribir(usuarios)
    return perfil_publico(u)


# ============================================================
# Sesiones (cookie firmada)
# ============================================================
def _secreto():
    env = os.environ.get('SESION_SECRETO', '').strip()
    if env:
        return env.encode()
    with _lock:
        if os.path.exists(ARCHIVO_SECRETO):
            with open(ARCHIVO_SECRETO, 'rb') as f:
                guardado = f.read().strip()
            if guardado:
                return guardado
        nuevo = secrets.token_hex(32).encode()
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(ARCHIVO_SECRETO, 'wb') as f:
            f.write(nuevo)
        try:
            os.chmod(ARCHIVO_SECRETO, 0o600)
        except OSError:
            pass
        return nuevo


def crear_token(usuario, horas=None):
    horas = SESION_HORAS if horas is None else horas
    cuerpo = json.dumps({'u': _clave_usuario(usuario),
                         'exp': int(time.time() + horas * 3600)},
                        separators=(',', ':')).encode()
    datos = base64.urlsafe_b64encode(cuerpo).decode().rstrip('=')
    firma = hmac.new(_secreto(), datos.encode(), hashlib.sha256).hexdigest()
    return f'{datos}.{firma}'


def leer_token(token):
    """Usuario del token si la firma es válida y no expiró, si no None."""
    try:
        datos, firma = str(token or '').split('.', 1)
    except ValueError:
        return None
    esperada = hmac.new(_secreto(), datos.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(firma, esperada):
        return None
    try:
        relleno = '=' * (-len(datos) % 4)
        payload = json.loads(base64.urlsafe_b64decode(datos + relleno))
    except (ValueError, TypeError):
        return None
    if payload.get('exp', 0) < time.time():
        return None
    u = obtener(payload.get('u'))
    if not u or not u.get('activo', True):
        return None
    return perfil_publico(u)
