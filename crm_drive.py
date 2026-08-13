# -*- coding: utf-8 -*-
"""crm_drive.py
============
Lectura y escritura de las hojas AGENDADOS y VENTA DIARIA (Google Drive)
para la página web tipo CRM de Derma Essenza.

- ``leer_*`` devuelve las filas actuales (con caché corta en TMP_DIR).
- ``agregar_*`` añade una fila nueva SIN crear archivos nuevos (el maestro
  sigue alimentándose de los .xlsx de siempre):
    * En .xlsx clásico usa cirugía XML (preserva autoFiltros, estilos,
      tablas dinámicas y fórmulas) y sube de vuelta conservando el mismo
      file id, con concurrencia best effort: un ``_lock`` serializa las
      escrituras de la app, un control preflight re-descarga y reintenta si
      la revisión del archivo cambió mientras se construía la fila (alguien
      editó directo en Drive), y un control postflight detecta con
      revisions.list si otra edición se coló y lo reporta en el resultado.
    * Si el archivo fuera una Google Sheet nativa, usa la Sheets API
      (values.append con INSERT_ROWS): append atómico del servidor.

Reutiliza la configuración de alimentar_maestro.py (CREDENCIALES,
AGENDADOS_FID, VENTA_FID, TMP_DIR) y sus normalizadores (num, esc_xml).
"""
import io
import os
import re
import shutil
import threading
import time
import zipfile
from datetime import datetime

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import alimentar_maestro as am  # noqa: E402

SCOPE = 'https://www.googleapis.com/auth/drive'
MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
MIME_SHEETS = 'application/vnd.google-apps.spreadsheet'
CACHE_TTL = 120          # segundos que se conserva la copia local antes de re-descargar
MAX_REINTENTOS = 3       # reintentos ante conflicto 412 (alguien editó el xlsx a la vez)
_SELECTO_LIMIT = 200

_lock = threading.Lock()

# Columnas de cada hoja (letra -> nombre mostrable en el CRM)
AGENDADOS_COLS = {
    'B': 'CRM', 'C': 'DIA', 'D': 'MES', 'E': 'AÑO', 'G': 'NOMBRE',
    'H': 'RED SOCIAL', 'I': 'TELEFONO', 'J': 'CORREO', 'K': 'AGENDADO POR',
    'L': 'DIA CITA', 'M': 'MES CITA', 'N': 'AÑO CITA', 'O': 'CAMPAÑA',
    'P': 'HORA', 'Q': 'CONFIRMADO', 'R': 'OBSERVACION', 'S': 'RECONFIRMADO',
    'T': 'OBSERVACION2',
}
VENTA_COLS = {
    'B': 'DIA', 'C': 'MES', 'D': 'AÑO', 'E': 'DNI', 'F': 'CEL. PACIENTE',
    'G': 'NOMBRE Y APELLIDO', 'H': 'NUEVO/RECURRENTE', 'I': 'DISTRITO',
    'J': 'EDAD', 'K': 'SEXO', 'L': 'TRATAMIENTO', 'M': 'DOCTOR',
    'N': 'STATUS', 'O': 'VENTA', 'P': 'PAGO', 'Q': 'COMISIONA',
    'R': 'OBSERVACION',
}

# Campo lógico -> letra canónica (layout de Beauty Medic)
AGENDADOS_CANON = {
    'crm': 'B', 'dia': 'C', 'mes': 'D', 'anio': 'E',
    'dni': 'F', 'nombre': 'G', 'red_social': 'H', 'telefono': 'I',
    'asistencia': 'A', 'correo': 'J',
    'agendado_por': 'K', 'dia_cita': 'L', 'mes_cita': 'M', 'anio_cita': 'N',
    'campana': 'O', 'hora': 'P', 'confirmado': 'Q', 'observacion': 'R',
    'reconfirmado': 'S', 'observacion2': 'T',
}
# Campo lógico -> letra canónica para VENTA (layout de Beauty Medic)
VENTA_CANON = {
    'dia': 'B', 'mes': 'C', 'anio': 'D', 'dni': 'E', 'cel': 'F',
    'nombre': 'G', 'nuevo': 'H', 'distrito': 'I', 'edad': 'J', 'sexo': 'K',
    'tratamiento': 'L', 'doctor': 'M', 'status': 'N', 'venta': 'O',
    'pago': 'P', 'comisiona': 'Q', 'campana': 'S', 'observacion': 'R',
}
# Nombres de encabezado normalizados (sin acentos ni Ñ) -> campo lógico
_HEADERS_AGENDADOS = {
    'CRM': 'crm', 'NOMBRE': 'nombre', 'RED SOCIAL': 'red_social',
    'TELEFONO': 'telefono', 'CORREO': 'correo', 'AGENDADO POR': 'agendado_por',
    'CAMPANA': 'campana', 'HORA': 'hora',
    'CONFIRMADO': 'confirmado', 'OBSERVACION': 'observacion',
    'RECONFIRMADO': 'reconfirmado', 'OBSERVACION2': 'observacion2',
    'DNI': 'dni', 'ASISTENCIA': 'asistencia',
}
_HEADERS_VENTA = {
    'DNI': 'dni', 'CEL': 'cel', 'CEL. PACIENTE': 'cel',
    'NOMBRE Y APELLIDO': 'nombre', 'NOMBRE': 'nombre',
    'NUEVO/RECURRENTE': 'nuevo', 'DISTRITO': 'distrito',
    'DISTRITO/DEPARTAMENTO': 'distrito', 'EDAD': 'edad', 'SEXO': 'sexo',
    'TRATAMIENTO': 'tratamiento', 'DOCTOR': 'doctor', 'STATUS': 'status',
    'VENTA': 'venta', 'PAGO': 'pago', 'COMISIONA': 'comisiona',
    'CAMPANA': 'campana', 'OBSERVACION': 'observacion',
}


def _normalizar_header(texto):
    s = str(texto or '').strip().upper()
    for a, b in (('Á', 'A'), ('É', 'E'), ('Í', 'I'), ('Ó', 'O'),
                 ('Ú', 'U'), ('Ñ', 'N')):
        s = s.replace(a, b)
    return re.sub(r'\s+', ' ', s)


def _campo_de_header(texto, mapas_header, contador):
    t = _normalizar_header(texto)
    if t in ('DIA', 'MES', 'ANO'):
        base = {'DIA': 'dia', 'MES': 'mes', 'ANO': 'anio'}[t]
        contador[base] = contador.get(base, 0) + 1
        return base if contador[base] == 1 else f'{base}_cita'
    if t in ('DIA 2', 'DIA CITA'):
        return 'dia_cita'
    if t in ('MES 2', 'MES CITA'):
        return 'mes_cita'
    if t in ('ANO 2', 'ANO CITA'):
        return 'anio_cita'
    return mapas_header.get(t)


def _detectar_columnas(ruta, hoja, mapas_header):
    """Detecta la fila de encabezados de la hoja.

    Devuelve (n_fila_encabezados, {letra_real: campo},
    {letra_real: encabezado_texto}).
    """
    wb = openpyxl.load_workbook(ruta, data_only=True)
    ws = wb[hoja]
    mejor = None
    for r in range(1, 13):
        contador = {}
        campos = {}
        textos = {}
        for c in range(1, 25):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            letra = openpyxl.utils.get_column_letter(c)
            campo = _campo_de_header(v, mapas_header, contador)
            if campo:
                campos[letra] = campo
                textos[letra] = str(v).strip()
        if len(campos) >= 3 and (mejor is None
                                 or len(campos) > len(mejor[1])):
            mejor = (r, campos, textos)
    wb.close()
    if mejor is None:
        return 4, {}, {}
    return mejor


def _mapa_inverso(campos, canon):
    """{letra_canónica: letra_real} a partir de los campos detectados."""
    out = {}
    for real, campo in campos.items():
        cl = canon.get(campo)
        if cl:
            out[cl] = real
    return out


def _remapear(valores, inverso):
    """Re-mapea {letra_canónica: valor} a {letra_real: valor} (sólo campos presentes)."""
    return {inverso[cl]: v for cl, v in valores.items() if cl in inverso}


def _adaptador_agendados(ruta, valores):
    _, campos, _ = _detectar_columnas(ruta, 'AGENDADOS', _HEADERS_AGENDADOS)
    return _remapear(valores, _mapa_inverso(campos, AGENDADOS_CANON))


def _adaptador_venta(ruta, hoja, valores):
    _, campos, _ = _detectar_columnas(ruta, hoja, _HEADERS_VENTA)
    return _remapear(valores, _mapa_inverso(campos, VENTA_CANON))

# Campos para las listas de valores únicos (selects del formulario)
AGENDADOS_SELECTS = {'crm': 'B', 'red_social': 'H', 'agendado_por': 'K',
                     'campana': 'O', 'confirmado': 'Q', 'reconfirmado': 'S'}
VENTA_SELECTS = {'nuevo': 'H', 'distrito': 'I', 'sexo': 'K', 'tratamiento': 'L',
                 'doctor': 'M', 'status': 'N', 'pago': 'P'}

# Estilo por defecto por hoja (sólo se usa si la última fila no define el estilo)
DEFAULT_STYLE = {'AGENDADOS': '115', 'VENTA 2026': '137', 'VENTA 2025': '99'}

# ============================================================
# DRIVE
# ============================================================
def _credenciales():
    am.garantizar_credenciales()
    if not os.path.exists(am.CREDENCIALES):
        raise FileNotFoundError(
            f'No hay credenciales en {am.CREDENCIALES}. Configura CREDENCIALES '
            'o GDRIVE_CREDENTIALS_JSON.')
    return service_account.Credentials.from_service_account_file(
        am.CREDENCIALES, scopes=[SCOPE])


def _drive():
    return build('drive', 'v3', credentials=_credenciales())


def _sheets():
    return build('sheets', 'v4', credentials=_credenciales())


def _es_sheets(fid):
    """True si el archivo de Drive es una Google Sheet nativa."""
    meta = _drive().files().get(fileId=fid, fields='mimeType').execute()
    return meta.get('mimeType') == MIME_SHEETS


def subir_drive(fid, ruta):
    """Sube el contenido de ruta a Drive conservando el mismo file id."""
    drv = _drive()
    with open(ruta, 'rb') as fh:
        data = fh.read()
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=MIME_XLSX,
                              resumable=False)
    drv.files().update(fileId=fid, media_body=media).execute()


def _revision_actual(fid):
    """Revisión actual (headRevisionId) del archivo en Drive."""
    meta = _drive().files().get(fileId=fid, fields='headRevisionId').execute()
    return meta.get('headRevisionId')


def _conflicto_post_subida(fid, rev_padre):
    """Post-subida (best effort): True si entre la revisión que leímos y
    nuestra subida se coló otra edición (o nuestra subida no quedó como
    cabeza). La API v3 no ofrece actualización condicional, así que esto
    solo detecta y avisa; no puede evitarlo.
    """
    try:
        res = _drive().files().revisions().list(
            fileId=fid, pageSize=10, fields='revisions(id,modifiedTime)',
            orderBy='modifiedTime desc').execute()
        for i, r in enumerate(res.get('revisions', [])):
            if r['id'] == rev_padre:
                return i != 1
        return False
    except Exception:
        return False


def _ruta_cache(nombre):
    return os.path.join(am.TMP_DIR, f'{nombre}.xlsx')


def descargar(nombre, fid, forzar=False):
    """Descarga de Drive a TMP_DIR con caché de CACHE_TTL segundos.

    Si el archivo es un Google Sheets (Docs Editors), lo exporta a .xlsx.
    """
    ruta = _ruta_cache(nombre)
    if (not forzar and os.path.exists(ruta)
            and time.time() - os.path.getmtime(ruta) < CACHE_TTL):
        return ruta
    drv = _drive()
    meta = drv.files().get(fileId=fid, fields='mimeType').execute()
    if meta.get('mimeType') == MIME_XLSX:
        req = drv.files().get_media(fileId=fid)
    else:
        req = drv.files().export(fileId=fid, mimeType=MIME_XLSX)
    with open(ruta, 'wb') as fh:
        dl = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
    return ruta


def invalidar(nombre):
    ruta = _ruta_cache(nombre)
    if os.path.exists(ruta):
        os.remove(ruta)


def fecha_descarga(nombre):
    ruta = _ruta_cache(nombre)
    if os.path.exists(ruta):
        return datetime.fromtimestamp(os.path.getmtime(ruta)).strftime('%d/%m/%Y %H:%M')
    return None


# ============================================================
# CIRUGÍA XML PARA AGREGAR FILAS
# ============================================================
def _mapa_hojas(ruta):
    """{nombre de hoja: ruta dentro del zip (xl/worksheets/sheetN.xml)}."""
    with zipfile.ZipFile(ruta) as z:
        wb = z.read('xl/workbook.xml').decode('utf-8')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="worksheets/([^"]+)"', rels))
    out = {}
    for m in re.finditer(r'<sheet\b[^>]*>', wb):
        tag = m.group(0)
        nm = re.search(r'name="([^"]+)"', tag)
        rid = re.search(r'r:id="(rId\d+)"', tag)
        if nm and rid and rid.group(1) in rid_target:
            out[nm.group(1)] = 'xl/worksheets/' + rid_target[rid.group(1)]
    return out


def _ultima_fila_datos(x):
    """Última fila (número) que contiene al menos una celda con valor."""
    ultima = 0
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*</row>', x, re.S):
        if re.search(r'<v>|<is>', m.group(0)):
            ultima = max(ultima, int(m.group(1)))
    return ultima


def _recortar_trailing(x, ultima):
    """Elimina las filas vacías de formato que quedan tras la última fila con datos."""
    def _keep(m):
        r = int(m.group(1))
        if r <= ultima or re.search(r'<v>|<is>', m.group(0)):
            return m.group(0)
        return ''
    return re.sub(r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*</row>', _keep, x)


def _estilos_fila(x, fila):
    """Estilos (s) de cada columna de una fila."""
    m = re.search(r'<row r="%d"[^>]*>(?:(?!</row>).)*</row>' % fila, x, re.S)
    if not m:
        return {}
    estilos = {}
    for c in re.finditer(r'<c r="([A-Z]+)%d"([^>]*)>' % fila, m.group(0)):
        s = re.search(r's="(\d+)"', c.group(2))
        if s:
            estilos[c.group(1)] = s.group(1)
    return estilos


def _celda_xml(col, fila, valor, estilo):
    ref = f'{col}{fila}'
    s = f' s="{estilo}"' if estilo else ''
    if isinstance(valor, bool):
        return ''
    if isinstance(valor, (int, float)):
        v = int(valor) if float(valor) == int(valor) else valor
        return f'<c r="{ref}"{s}><v>{v}</v></c>'
    t = am.esc_xml(valor)
    return f'<c r="{ref}"{s} t="inlineStr"><is><t>{t}</t></is></c>'


def agregar_fila_xlsx(origen, destino, hoja, fila_num, valores, estilo_default):
    """Agrega una fila nueva (valores {col: valor}) a la hoja indicada."""
    build_dir = os.path.join(am.TMP_DIR, 'build_crm')
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    os.makedirs(build_dir)
    with zipfile.ZipFile(origen) as z:
        z.extractall(build_dir)

    hoja_file = _mapa_hojas(origen).get(hoja)
    if not hoja_file:
        raise ValueError(f'Hoja "{hoja}" no encontrada en el archivo')
    sheet_path = os.path.join(build_dir, hoja_file)
    with open(sheet_path, encoding='utf-8') as f:
        x = f.read()

    ultima = _ultima_fila_datos(x)
    fila_num = max(fila_num, ultima + 1)
    # quitar filas de formato vacías finales para que la fila nueva quede contigua
    x = _recortar_trailing(x, ultima)

    estilos = _estilos_fila(x, ultima)
    celdas = []
    for col in sorted(valores, key=lambda c: openpyxl.utils.column_index_from_string(c)):
        v = valores[col]
        if v is None or v == '':
            continue
        celdas.append(_celda_xml(col, fila_num, v, estilos.get(col, estilo_default)))
    fila_xml = f'<row r="{fila_num}" ht="14.25" customHeight="1">' + ''.join(celdas) + '</row>'
    x = x.replace('</sheetData>', fila_xml + '</sheetData>', 1)

    # extender los autoFilter si la fila nueva cae fuera de su rango
    def _extender(m):
        if int(m.group(4)) >= fila_num:
            return m.group(0)
        return f'{m.group(1)}{m.group(2)}{m.group(3)}{fila_num}{m.group(5)}'

    x = re.sub(r'(<autoFilter ref="\$[A-Z]+\$\d+:\$)([A-Z]+)(\$)(\d+)(")',
               _extender, x)

    with open(sheet_path, 'w', encoding='utf-8') as f:
        f.write(x)

    if os.path.exists(destino):
        os.remove(destino)
    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(build_dir):
            for fn in files:
                full = os.path.join(root, fn)
                z.write(full, os.path.relpath(full, build_dir))
    shutil.rmtree(build_dir)


def _siguiente_fila(ruta, hoja):
    hoja_file = _mapa_hojas(ruta).get(hoja)
    if not hoja_file:
        raise ValueError(f'Hoja "{hoja}" no encontrada')
    with zipfile.ZipFile(ruta) as z:
        x = z.read(hoja_file).decode('utf-8')
    return _ultima_fila_datos(x) + 1


# ============================================================
# NORMALIZACIÓN
# ============================================================
def _mes(m):
    s = str(m or '').strip().upper()
    if not s:
        return ''
    return 'SET' if s == 'SEP' else s


def _tel(t):
    if t is None:
        return None
    d = re.sub(r'\D', '', str(t))
    return int(d) if d else None


def _int(v):
    n = am.num(v)
    return int(n) if n is not None else None


def _txt(v):
    s = str(v or '').strip()
    return s if s else None


def normalizar_agendado(d):
    out = {}
    if _txt(d.get('crm')): out['B'] = _txt(d['crm'])
    if _int(d.get('dia')): out['C'] = _int(d['dia'])
    if _mes(d.get('mes')): out['D'] = _mes(d['mes'])
    if _int(d.get('anio')): out['E'] = _int(d['anio'])
    if _txt(d.get('nombre')): out['G'] = _txt(d['nombre'])
    if _txt(d.get('red_social')): out['H'] = _txt(d['red_social'])
    tel = _tel(d.get('telefono'))
    if tel is not None: out['I'] = tel
    if _txt(d.get('correo')): out['J'] = _txt(d['correo'])
    if _txt(d.get('agendado_por')): out['K'] = _txt(d['agendado_por'])
    if _int(d.get('dia_cita')): out['L'] = _int(d['dia_cita'])
    if _mes(d.get('mes_cita')): out['M'] = _mes(d['mes_cita'])
    if _int(d.get('anio_cita')): out['N'] = _int(d['anio_cita'])
    if _txt(d.get('campana')): out['O'] = _txt(d['campana'])
    if _txt(d.get('hora')): out['P'] = _txt(d['hora'])
    if _txt(d.get('confirmado')): out['Q'] = _txt(d['confirmado'])
    if _txt(d.get('observacion')): out['R'] = _txt(d['observacion'])
    if _txt(d.get('reconfirmado')): out['S'] = _txt(d['reconfirmado'])
    if _txt(d.get('observacion2')): out['T'] = _txt(d['observacion2'])
    return out


def normalizar_venta(d, hoja='VENTA 2026'):
    out = {}
    if _int(d.get('dia')): out['B'] = _int(d['dia'])
    if _mes(d.get('mes')): out['C'] = _mes(d['mes'])
    if _int(d.get('anio')): out['D'] = _int(d['anio'])
    if _int(d.get('dni')): out['E'] = _int(d['dni'])
    tel = _tel(d.get('cel'))
    if tel is not None: out['F'] = tel
    if _txt(d.get('nombre')): out['G'] = _txt(d['nombre'])
    if _txt(d.get('nuevo')): out['H'] = _txt(d['nuevo'])
    if _txt(d.get('distrito')): out['I'] = _txt(d['distrito'])
    if _int(d.get('edad')): out['J'] = _int(d['edad'])
    if _txt(d.get('sexo')): out['K'] = _txt(d['sexo'])
    if _txt(d.get('tratamiento')): out['L'] = _txt(d['tratamiento'])
    if _txt(d.get('doctor')): out['M'] = _txt(d['doctor'])
    if _txt(d.get('status')): out['N'] = _txt(d['status'])
    if am.num(d.get('venta')) is not None: out['O'] = am.num(d['venta'])
    if _txt(d.get('pago')): out['P'] = _txt(d['pago'])
    if hoja == 'VENTA 2025':
        # la hoja 2025 sólo llega hasta Q (OBSERVACION en Q)
        if _txt(d.get('observacion')): out['Q'] = _txt(d['observacion'])
    else:
        if _int(d.get('comisiona')): out['Q'] = _int(d['comisiona'])
        if _txt(d.get('observacion')): out['R'] = _txt(d['observacion'])
    return out


# ============================================================
# AGREGAR FILAS (Google Sheet nativa: append atómico | xlsx: cirugía XML)
# ============================================================
def _append_sheets(fid, hoja, valores, ruta):
    """Escribe una fila nueva al final de los datos de una Google Sheet nativa.

    ``ruta`` es la copia exportada a .xlsx; de ella se calcula la fila destino
    (última fila con datos + 1) y se escribe con ``values.update``.

    No se usa ``values.append`` porque con hojas cuya columna A está vacía
    (como el AGENDADOS de Derma Essenza) Google no detecta la tabla e inserta
    la fila en la fila 1, corriendo todo el contenido hacia abajo.

    Devuelve el número de fila escrito.
    """
    srv = _sheets()
    fila = _siguiente_fila(ruta, hoja)
    letras = sorted(valores, key=openpyxl.utils.column_index_from_string)
    ncols = openpyxl.utils.column_index_from_string(letras[-1]) if letras else 1
    fila_row = [valores.get(openpyxl.utils.get_column_letter(c))
                for c in range(1, ncols + 1)]
    fila_row = ['' if v is None else v for v in fila_row]
    srv.spreadsheets().values().update(
        spreadsheetId=fid,
        range=f"'{hoja}'!A{fila}",
        valueInputOption='RAW',
        body={'values': [fila_row]}).execute()
    return fila


def _agregar_xlsx_seguro(nombre, fid, hoja, estilo, valores, adaptador=None):
    """Append sobre .xlsx de Drive con control de concurrencia best effort.

    La API v3 de Drive no tiene actualización condicional (CAS), así que:
      1. ``_lock`` serializa las escrituras de la propia app en el proceso.
      2. Preflight: si la revisión del archivo cambió mientras descargábamos
         y construíamos la fila (alguien editó directo en Drive), se
         re-descarga y se reintenta con la versión nueva.
      3. Postflight: tras subir, se comprueba con revisions.list si otra
         edición se coló en el ínterin y se reporta en el resultado.

    ``adaptador`` (opcional) recibe (ruta_descargada, valores) y devuelve
    los valores ya re-mapeados a las letras reales de la hoja.

    Devuelve (fila, conflicto).
    """
    destino = os.path.join(am.TMP_DIR, f'{nombre}_editado.xlsx')
    fila = None
    conflicto = False
    for intento in range(MAX_REINTENTOS):
        rev_antes = _revision_actual(fid)
        ruta = descargar(nombre, fid, forzar=True)
        valores_hoja = adaptador(ruta, valores) if adaptador else valores
        fila = _siguiente_fila(ruta, hoja)
        agregar_fila_xlsx(ruta, destino, hoja, fila, valores_hoja, estilo)
        if _revision_actual(fid) != rev_antes:
            if intento < MAX_REINTENTOS - 1:
                continue
            raise RuntimeError(
                'El archivo de Drive cambió durante la escritura; '
                'reintenta la operación')
        subir_drive(fid, destino)
        conflicto = _conflicto_post_subida(fid, rev_antes)
        break
    invalidar(nombre)
    return fila, conflicto


def agregar_agendado(datos):
    valores = normalizar_agendado(datos)
    if not valores.get('G') and 'I' not in valores:
        raise ValueError('Indica al menos el nombre o el teléfono del paciente')
    with _lock:
        fid = am.AGENDADOS_FID
        if _es_sheets(fid):
            ruta = descargar('AGENDADOS', fid, forzar=True)
            valores_hoja = _adaptador_agendados(ruta, valores)
            fila = _append_sheets(fid, 'AGENDADOS', valores_hoja, ruta)
            invalidar('AGENDADOS')
            return {'ok': True, 'fila': fila, 'hoja': 'AGENDADOS',
                    'valores': valores}
        fila, conflicto = _agregar_xlsx_seguro('AGENDADOS', fid, 'AGENDADOS',
                                               DEFAULT_STYLE['AGENDADOS'],
                                               valores,
                                               adaptador=_adaptador_agendados)
    return {'ok': True, 'fila': fila, 'hoja': 'AGENDADOS',
            'valores': valores, 'conflicto': conflicto}


def agregar_venta(datos):
    anio = am.num(datos.get('anio')) or 2026
    hoja = 'VENTA 2026' if anio >= 2026 else 'VENTA 2025'
    valores = normalizar_venta(datos, hoja)
    if not valores.get('G'):
        raise ValueError('Indica el nombre del paciente')

    def adaptador(ruta, v):
        return _adaptador_venta(ruta, hoja, v)

    with _lock:
        fid = am.VENTA_FID
        if _es_sheets(fid):
            ruta = descargar('VENTA_DIARIA', fid, forzar=True)
            valores_hoja = adaptador(ruta, valores)
            fila = _append_sheets(fid, hoja, valores_hoja, ruta)
            invalidar('VENTA_DIARIA')
            return {'ok': True, 'fila': fila, 'hoja': hoja, 'valores': valores}
        fila, conflicto = _agregar_xlsx_seguro('VENTA_DIARIA', fid, hoja,
                                               DEFAULT_STYLE[hoja], valores,
                                               adaptador=adaptador)
    return {'ok': True, 'fila': fila, 'hoja': hoja, 'valores': valores,
            'conflicto': conflicto}


# ============================================================
# LECTURA PARA MOSTRAR
# ============================================================
def leer_agendados():
    """Devuelve {'filas': [{letra_canónica: valor}...], 'total': n,
    'columnas': {letra_canónica: encabezado}}."""
    ruta = descargar('AGENDADOS', am.AGENDADOS_FID)
    n_fila, campos, textos = _detectar_columnas(ruta, 'AGENDADOS',
                                                _HEADERS_AGENDADOS)
    inverso = _mapa_inverso(campos, AGENDADOS_CANON)
    if not inverso:
        n_fila = 4
        inverso = {c: c for c in AGENDADOS_COLS}
        textos = AGENDADOS_COLS
    columnas = {}
    for cl, real in sorted(inverso.items(),
                           key=lambda kv: openpyxl.utils.column_index_from_string(kv[1])):
        columnas[cl] = textos[real]
    wb = openpyxl.load_workbook(ruta, data_only=True)
    ws = wb['AGENDADOS']
    filas = []
    for r in range(n_fila + 1, ws.max_row + 1):
        fila = {}
        for cl, real in inverso.items():
            v = ws.cell(row=r,
                        column=openpyxl.utils.column_index_from_string(real)).value
            if v is None:
                continue
            if isinstance(v, float) and v == int(v):
                v = int(v)
            fila[cl] = v
        if fila:
            filas.append(fila)
    wb.close()
    return {'filas': filas, 'total': len(filas),
            'descargado': fecha_descarga('AGENDADOS'), 'columnas': columnas}


def leer_venta():
    """Devuelve {'hojas': {hoja: {'filas': [...], 'columnas': {...}}}}."""
    ruta = descargar('VENTA_DIARIA', am.VENTA_FID)
    wb = openpyxl.load_workbook(ruta, data_only=True)
    out = {}
    for hoja in ('VENTA 2026', 'VENTA 2025'):
        if hoja not in wb.sheetnames:
            continue
        n_fila, campos, textos = _detectar_columnas(ruta, hoja, _HEADERS_VENTA)
        inverso = _mapa_inverso(campos, VENTA_CANON)
        if not inverso:
            n_fila = 5 if hoja == 'VENTA 2026' else 8
            inverso = {c: c for c in VENTA_COLS}
            textos = VENTA_COLS
        columnas = {}
        for cl, real in sorted(inverso.items(),
                               key=lambda kv: openpyxl.utils.column_index_from_string(kv[1])):
            columnas[cl] = textos[real]
        ws = wb[hoja]
        filas = []
        for r in range(n_fila + 1, ws.max_row + 1):
            fila = {}
            for cl, real in inverso.items():
                v = ws.cell(row=r,
                            column=openpyxl.utils.column_index_from_string(real)).value
                if v is None:
                    continue
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                fila[cl] = v
            if fila:
                filas.append(fila)
        out[hoja] = {'filas': filas, 'columnas': columnas}
    wb.close()
    return {'hojas': out, 'descargado': fecha_descarga('VENTA_DIARIA')}


def _unicos(filas, col):
    s = set()
    for f in filas:
        v = f.get(col)
        if v is not None and str(v).strip():
            s.add(str(v).strip())
    return sorted(s)[:_SELECTO_LIMIT]


def valores_unicos_agendados(filas):
    return {k: _unicos(filas, col) for k, col in AGENDADOS_SELECTS.items()}


def valores_unicos_venta(filas):
    return {k: _unicos(filas, col) for k, col in VENTA_SELECTS.items()}
