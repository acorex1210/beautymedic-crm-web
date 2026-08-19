"""Historias clínicas en PDF por paciente.

Guarda los archivos en disco (``DATA_DIR/historias``), con un índice JSON por
paciente. Sigue el mismo patrón que ``meta_ads.py``: el directorio base lo
decide quien llama, así el módulo no depende de la configuración de la app.

Estructura:
    <dir>/<clave_segura>/index.json      metadatos de los documentos
    <dir>/<clave_segura>/<doc_id>.pdf    el archivo

La clave del paciente viene del directorio (``t<telefono>`` o ``n<nombre>``),
así que se sanea antes de usarla como nombre de carpeta.
"""
import json
import os
import re
from datetime import datetime

MAX_BYTES = 15 * 1024 * 1024   # 15 MB por archivo
CABECERA_PDF = b'%PDF-'


def _clave_segura(clave):
    """Nombre de carpeta seguro para una clave de paciente.

    Sólo deja [A-Za-z0-9_-]; evita que una clave con '..' o '/' escape del
    directorio base.
    """
    s = re.sub(r'[^A-Za-z0-9_-]', '_', str(clave or '').strip())
    if not s:
        raise ValueError('Clave de paciente vacía')
    return s[:120]


def _doc_id_seguro(doc_id):
    s = re.sub(r'[^A-Za-z0-9_-]', '', str(doc_id or ''))
    if not s:
        raise ValueError('Documento inválido')
    return s


def carpeta(clave):
    """Nombre de carpeta que le corresponde a una clave de paciente."""
    return _clave_segura(clave)


def _dir_paciente(dir, clave):
    return os.path.join(dir, _clave_segura(clave))


def _ruta_index(dir, clave):
    return os.path.join(_dir_paciente(dir, clave), 'index.json')


def _leer_index(dir, clave):
    ruta = _ruta_index(dir, clave)
    if not os.path.exists(ruta):
        return []
    try:
        with open(ruta, encoding='utf-8') as f:
            return json.load(f)
    except (ValueError, OSError):
        return []


def _escribir_index(dir, clave, index):
    os.makedirs(_dir_paciente(dir, clave), exist_ok=True)
    with open(_ruta_index(dir, clave), 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def _nombre_limpio(nombre):
    n = os.path.basename(str(nombre or '')).strip() or 'historia.pdf'
    if not n.lower().endswith('.pdf'):
        n += '.pdf'
    return n[:150]


def guardar(dir, clave, contenido, nombre, nota=''):
    """Guarda un PDF de historia clínica. Lanza ValueError si no es válido."""
    if not contenido:
        raise ValueError('El archivo está vacío')
    if len(contenido) > MAX_BYTES:
        raise ValueError(f'El archivo supera el máximo de {MAX_BYTES // (1024 * 1024)} MB')
    if not contenido.startswith(CABECERA_PDF):
        raise ValueError('El archivo no es un PDF válido')
    ahora = datetime.now()
    doc_id = ahora.strftime('%Y%m%d_%H%M%S_%f')
    carpeta = _dir_paciente(dir, clave)
    os.makedirs(carpeta, exist_ok=True)
    with open(os.path.join(carpeta, f'{doc_id}.pdf'), 'wb') as f:
        f.write(contenido)
    doc = {
        'id': doc_id,
        'archivo': _nombre_limpio(nombre),
        'fecha': ahora.isoformat(timespec='seconds'),
        'bytes': len(contenido),
        'nota': str(nota or '').strip()[:300],
    }
    index = _leer_index(dir, clave)
    index.insert(0, doc)
    _escribir_index(dir, clave, index)
    return doc


def listar(dir, clave):
    return _leer_index(dir, clave)


def ruta(dir, clave, doc_id):
    """Ruta del PDF en disco. Lanza ValueError si no existe."""
    doc_id = _doc_id_seguro(doc_id)
    p = os.path.join(_dir_paciente(dir, clave), f'{doc_id}.pdf')
    if not os.path.exists(p):
        raise ValueError('Documento no encontrado')
    return p


def documento(dir, clave, doc_id):
    doc_id = _doc_id_seguro(doc_id)
    for d in _leer_index(dir, clave):
        if d['id'] == doc_id:
            return d
    raise ValueError('Documento no encontrado')


def borrar(dir, clave, doc_id):
    doc_id = _doc_id_seguro(doc_id)
    index = _leer_index(dir, clave)
    resto = [d for d in index if d['id'] != doc_id]
    if len(resto) == len(index):
        raise ValueError('Documento no encontrado')
    p = os.path.join(_dir_paciente(dir, clave), f'{doc_id}.pdf')
    if os.path.exists(p):
        os.remove(p)
    _escribir_index(dir, clave, resto)
    return {'ok': True, 'borrado': doc_id}


def conteos(dir):
    """{clave_segura: n_documentos} para marcar en el listado quién tiene historia."""
    if not os.path.isdir(dir):
        return {}
    out = {}
    for nombre in os.listdir(dir):
        carpeta = os.path.join(dir, nombre)
        if not os.path.isdir(carpeta):
            continue
        ruta_idx = os.path.join(carpeta, 'index.json')
        if not os.path.exists(ruta_idx):
            continue
        try:
            with open(ruta_idx, encoding='utf-8') as f:
                out[nombre] = len(json.load(f))
        except (ValueError, OSError):
            continue
    return out
