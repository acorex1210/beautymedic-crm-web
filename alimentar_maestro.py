# -*- coding: utf-8 -*-
"""
alimentar_maestro.py
====================
Sincroniza el archivo maestro 'BD DATA.xlsx' desde dos fuentes de Google Drive:

  1) AGENDADOS  -> copia columnas B..O (nuevas citas agendadas)
  2) VENTA DIARIA -> consulta consultiva: completa ASISTENCIA, DISTRITO, EDAD,
     SEXO, DNI y los TRAT/PAGO 1..4 + PAGO TOTAL de quien asistió.

La edición del maestro se hace por cirugía XML (se preservan tablas dinámicas,
slicers, fórmulas y formato). No se usa openpyxl para guardar.

Uso:
  python3 alimentar_maestro.py            # modo revisión (no toca el maestro)
  python3 alimentar_maestro.py --aplicar  # aplica cambios al maestro (con backup)
  python3 alimentar_maestro.py --sin-descarga   # reutiliza archivos ya descargados

Configuración vía variables de entorno (útiles para despliegue web):
  CREDENCIALES           ruta al JSON de la cuenta de servicio (default ~/credenciales-bm.json)
  GDRIVE_CREDENTIALS_JSON contenido del JSON (JSON crudo o base64; se escribe en CREDENCIALES)
  GOOGLE_APPLICATION_CREDENTIALS ruta a un archivo JSON, contenido crudo o base64 (alternativa)
  AGENDADOS_FID          id de archivo AGENDADOS en Drive
  VENTA_FID              id de archivo VENTA DIARIA en Drive
  MAESTRO_PATH           ruta al maestro BD DATA.xlsx
  TMP_DIR                directorio temporal para descargas y simulaciones
"""
import argparse
import base64
import io
import json
import os
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from datetime import datetime

import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ============================================================
# CONFIGURACIÓN
# ============================================================
_TMP_DEFAULT = os.path.expanduser(
    '/var/folders/fc/2jp7n0610cbfckv3jpbzt73h0000gn/T/opencode/maestro_auto')
CREDENCIALES = os.environ.get('CREDENCIALES',
                              os.path.expanduser('~/credenciales-bm.json'))
GDRIVE_CREDENTIALS_JSON = os.environ.get('GDRIVE_CREDENTIALS_JSON', '')
AGENDADOS_FID = os.environ.get('AGENDADOS_FID', '12fWJpIBpr3GH7Yj57iyyndm_m37rr7V2')
VENTA_FID = os.environ.get('VENTA_FID', '1LHtZk0vAGgnyOsODwU6f4LvtUoQxWNis')
MAESTRO = os.environ.get('MAESTRO_PATH',
                         os.path.expanduser('~/Downloads/BD DATA.xlsx'))
TMP_DIR = os.environ.get('TMP_DIR', _TMP_DEFAULT)
os.makedirs(TMP_DIR, exist_ok=True)


def _normalizar_json_credenciales(v):
    """Devuelve el JSON crudo de la cuenta de servicio, aceptando JSON o base64."""
    if v is None:
        return ''
    v = str(v).strip()
    if not v:
        return ''
    try:
        json.loads(v)
        return v
    except (ValueError, TypeError):
        pass
    try:
        dec = base64.b64decode(v, validate=True).decode('utf-8')
        json.loads(dec)
        return dec
    except (ValueError, TypeError, Exception):  # noqa: BLE001
        return v


def _contenido_credenciales():
    """Busca el JSON de credenciales en el orden: GDRIVE_CREDENTIALS_JSON,
    GOOGLE_APPLICATION_CREDENTIALS (ruta o contenido), CREDENCIALES (archivo)."""
    for var in ('GDRIVE_CREDENTIALS_JSON', 'GOOGLE_APPLICATION_CREDENTIALS'):
        v = os.environ.get(var, '')
        if not v:
            continue
        if os.path.isfile(v):
            try:
                with open(v, encoding='utf-8') as f:
                    return f.read()
            except OSError:
                continue
        return _normalizar_json_credenciales(v)
    if os.path.exists(CREDENCIALES):
        try:
            with open(CREDENCIALES, encoding='utf-8') as f:
                return f.read()
        except OSError:
            return ''
    return ''


def garantizar_credenciales():
    """Escribe las credenciales en CREDENCIALES si no existen y hay fuente."""
    if not os.path.exists(CREDENCIALES):
        contenido = _contenido_credenciales()
        if contenido:
            os.makedirs(os.path.dirname(CREDENCIALES) or '.', exist_ok=True)
            with open(CREDENCIALES, 'w', encoding='utf-8') as f:
                f.write(contenido)


def credenciales_disponibles():
    return os.path.exists(CREDENCIALES) or bool(_contenido_credenciales())

MESES = {'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN',
         'JUL', 'AGO', 'SET', 'SEP', 'OCT', 'NOV', 'DIC'}

# STATUS de VENTA DIARIA que indican que el tratamiento SÍ se realizó
ASISTE_POR_TEXTO = ('SE REALIZO', 'COMPRO', 'COMPLETA', 'SESION')
REVISAR_STATUS = ('DEJO PAGADO',)

# ============================================================
# UTILIDADES DE NORMALIZACIÓN
# ============================================================
_ACENTOS = {'Á': 'A', 'À': 'A', 'Ä': 'A', 'É': 'E', 'È': 'E', 'Ë': 'E',
            'Í': 'I', 'Ì': 'I', 'Ï': 'I', 'Ó': 'O', 'Ò': 'O', 'Ö': 'O',
            'Ú': 'U', 'Ù': 'U', 'Ü': 'U', 'Ñ': 'N'}


def norm_name(n):
    if n is None:
        return ''
    s = str(n).upper().strip()
    s = ''.join(_ACENTOS.get(ch, ch) for ch in s)
    return re.sub(r'\s+', ' ', s)


def norm_phone(p):
    if p is None or p == '':
        return ''
    if isinstance(p, float) and p.is_integer():
        p = int(p)
    d = re.sub(r'\D', '', str(p))
    return d[-9:] if len(d) >= 9 else d


def norm_fecha(dia, mes, anio):
    try:
        d = int(dia)
    except (TypeError, ValueError):
        d = None
    m = str(mes).strip().upper() if mes is not None else None
    if m in ('SEP',):
        m = 'SET'
    try:
        a = int(anio)
    except (TypeError, ValueError):
        a = None
    return (d, m, a)


def num(x):
    if x is None:
        return None
    try:
        v = float(x)
        return int(v) if v == int(v) else v
    except (TypeError, ValueError):
        return None


def txt(x):
    if x is None:
        return None
    s = str(x).strip()
    return s if s else None


# ============================================================
# DESCARGA DESDE DRIVE
# ============================================================
def descargar(fid, nombre, forzar=False):
    """Descarga de Drive a TMP_DIR (o reutiliza la copia local si existe)."""
    ruta = os.path.join(TMP_DIR, f'{nombre}.xlsx')
    if os.path.exists(ruta) and not forzar:
        return ruta
    garantizar_credenciales()
    if not os.path.exists(CREDENCIALES):
        raise FileNotFoundError(
            f'No hay credenciales en {CREDENCIALES}. Configura CREDENCIALES o '
            'GDRIVE_CREDENTIALS_JSON.')
    creds = service_account.Credentials.from_service_account_file(
        CREDENCIALES, scopes=['https://www.googleapis.com/auth/drive.readonly'])
    drive = build('drive', 'v3', credentials=creds)
    req = drive.files().get_media(fileId=fid)
    with open(ruta, 'wb') as fh:
        dl = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
    print(f'  Descargado {nombre} -> {os.path.basename(ruta)}')
    return ruta


# ============================================================
# LECTURA DE FUENTES (solo lectura)
# ============================================================
def leer_agendados(path):
    """Devuelve lista de dicts {col: valor} filas B..O de la hoja AGENDADOS."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['AGENDADOS']
    filas = []
    for r in range(5, ws.max_row + 1):
        if ws.cell(row=r, column=3).value is None and ws.cell(row=r, column=7).value is None:
            continue
        fila = {}
        for c in range(2, 16):  # B..O
            v = ws.cell(row=r, column=c).value
            if v is not None:
                fila[openpyxl.utils.get_column_letter(c)] = v
        filas.append((r, fila))
    return filas


def leer_venta(path):
    """Devuelve lista de dicts de las hojas VENTA 2026 y VENTA 2025."""
    wb = openpyxl.load_workbook(path, data_only=True)
    filas = []
    for hoja, hr in [('VENTA 2026', 5), ('VENTA 2025', 8)]:
        if hoja not in wb.sheetnames:
            continue
        ws = wb[hoja]
        for r in range(hr + 1, ws.max_row + 1):
            if ws.cell(row=r, column=3).value is None and ws.cell(row=r, column=7).value is None:
                continue
            fila = {
                'hoja': hoja,
                'dia': ws.cell(row=r, column=2).value,
                'mes': ws.cell(row=r, column=3).value,
                'anio': ws.cell(row=r, column=4).value,
                'dni': ws.cell(row=r, column=5).value,
                'cel': ws.cell(row=r, column=6).value,
                'nombre': ws.cell(row=r, column=7).value,
                'distrito': ws.cell(row=r, column=9).value,
                'edad': ws.cell(row=r, column=10).value,
                'sexo': ws.cell(row=r, column=11).value,
                'tratamiento': ws.cell(row=r, column=12).value,
                'status': ws.cell(row=r, column=14).value,
                'venta': ws.cell(row=r, column=15).value,
                'fila': r,
            }
            filas.append(fila)
    return filas


def leer_maestro(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['BD DATA']
    return ws


def _rango_pivot(path, nombre_pivot):
    """Devuelve (hoja_origen, fila_inicio, fila_fin) del origen de una tabla
    dinámica (pivot) leyendo el XML del workbook, o None si no se encuentra."""
    try:
        with zipfile.ZipFile(path) as z:
            nombres = z.namelist()
            for pt in nombres:
                if not (pt.startswith('xl/pivotTables/') and pt.endswith('.xml')):
                    continue
                xml = z.read(pt).decode('utf-8', 'replace')
                m = re.search(r'<pivotTableDefinition[^>]*name="([^"]+)"', xml)
                if not m or m.group(1) != nombre_pivot:
                    continue
                target = None
                rels = pt.replace('.xml', '.xml.rels')
                if rels in nombres:
                    r = z.read(rels).decode('utf-8', 'replace')
                    m2 = re.search(r'Target="([^"]*pivotCache[^"]*)"', r)
                    if m2:
                        target = m2.group(1).replace('../', 'xl/')
                if target and target in nombres:
                    cdef = z.read(target).decode('utf-8', 'replace')
                else:
                    m2 = re.search(r'cacheId="(\d+)"', xml)
                    cid = int(m2.group(1)) if m2 else 0
                    cdef = z.read(
                        f'xl/pivotCache/pivotCacheDefinition{cid + 1}.xml'
                    ).decode('utf-8', 'replace')
                m3 = re.search(
                    r'<worksheetSource[^>]*ref="([A-Z]+)(\d+):[A-Z]+(\d+)"'
                    r'[^>]*sheet="([^"]+)"', cdef)
                if m3:
                    return m3.group(4), int(m3.group(2)), int(m3.group(3))
    except (KeyError, zipfile.BadZipFile):
        return None
    return None


def verificar_venta_vs_td(path, anio, mes, desde, hasta):
    """Verifica que la hoja VENTA {anio} coincida con su tabla dinámica TD {anio}
    para el periodo (mes, anio, desde..hasta).

    Devuelve dict con la suma de VENTA (col O) de la hoja completa, la suma del
    rango cubierto por la TD (pivot) y si coinciden. Si hay filas de venta fuera
    del rango de la tabla dinámica, algún dato está mal.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:  # noqa: BLE001
        return {'ok': False, 'error': 'No se pudo abrir VENTA_DIARIA.xlsx'}
    hoja = f'VENTA {anio}'
    if hoja not in wb.sheetnames:
        return {'ok': False, 'error': f'No hay hoja {hoja} en VENTA_DIARIA.xlsx'}
    ws = wb[hoja]
    rango = _rango_pivot(path, f'TD {anio}')
    fila_fin = ws.max_row
    if rango:
        fila_fin = min(ws.max_row, rango[2])
    venta_tot = 0.0
    td_tot = 0.0
    filas_fuera = 0
    for r in range(5, ws.max_row + 1):
        d = ws.cell(row=r, column=2).value
        m = ws.cell(row=r, column=3).value
        a = ws.cell(row=r, column=4).value
        v = ws.cell(row=r, column=15).value
        if (a == anio and m == mes and isinstance(d, (int, float))
                and desde <= int(d) <= hasta and isinstance(v, (int, float))):
            venta_tot += v
            if r <= fila_fin:
                td_tot += v
            else:
                filas_fuera += 1
    coincide = abs(venta_tot - td_tot) < 0.01
    mensaje = (f'VENTA {anio} y TD {anio} coinciden en el periodo: S/ {venta_tot:,.2f}'
               if coincide else
               f'VENTA {anio} (S/ {venta_tot:,.2f}) NO coincide con TD {anio} '
               f'(S/ {td_tot:,.2f}): {filas_fuera} fila(s) de venta quedan fuera del '
               'rango de la tabla dinámica.')
    return {'ok': True, 'venta': round(venta_tot, 2), 'td': round(td_tot, 2),
            'coincide': coincide, 'filas_fuera': filas_fuera, 'mensaje': mensaje}


# ============================================================
# CÁLCULO DE CAMBIOS
# ============================================================
class Calculo:
    def __init__(self, maestro_ws, agendados, venta):
        self.maestro = maestro_ws
        self.new_rows = []       # lista de (provisional, {col: valor})
        self.updates = {}        # {fila_maestro: {col: valor}} (existentes y provisionales)
        self.matches = []        # (venta_fila, maestro_fila, modo, hoja)
        self.pendientes = []     # ventas sin match
        self.revisar = []        # casos a revisar
        self.incompletas = []    # filas AGENDADOS incompletas/sin nombre
        self._nueva_info = {}    # provisional -> (ph, nm, fecha)
        self._indexar()
        self._calcular_nuevos(agendados)
        self._indexar_nuevos()
        self._calcular_ventas(venta)

    # ----- índice del maestro -----
    def _indexar(self):
        ws = self.maestro
        self.m_rows = []
        self.by_phone = defaultdict(list)
        self.by_name = defaultdict(list)
        self.keys_full = set()
        self.keys_loose = set()
        self.last_row = 4
        for r in range(5, ws.max_row + 1):
            if ws.cell(row=r, column=3).value is None and ws.cell(row=r, column=7).value is None:
                continue
            self.last_row = r
            self.m_rows.append(r)
            cel = {openpyxl.utils.get_column_letter(c): ws.cell(row=r, column=c).value
                   for c in range(2, 16) if ws.cell(row=r, column=c).value is not None}
            ph = norm_phone(cel.get('I'))
            nm = norm_name(cel.get('G'))
            fc = norm_fecha(cel.get('L'), cel.get('M'), cel.get('N'))
            if ph:
                self.by_phone[ph].append(r)
            if nm:
                self.by_name[nm].append(r)
            camp = norm_name(cel.get('O'))
            self.keys_full.add((ph, nm, fc, camp))
            if nm:
                self.keys_loose.add((nm, fc, camp))

    # ----- 1) AGENDADOS -> filas nuevas (con fila provisional) -----
    def _calcular_nuevos(self, agendados):
        vistos = set()
        for r, fila in agendados:
            ph = norm_phone(fila.get('I'))
            nm = norm_name(fila.get('G'))
            fc = norm_fecha(fila.get('L'), fila.get('M'), fila.get('N'))
            camp = norm_name(fila.get('O'))
            if not nm and not ph:
                self.incompletas.append((r, fila, 'sin nombre y sin teléfono'))
                continue
            if (ph, nm, fc, camp) in self.keys_full:
                continue
            if nm and (nm, fc, camp) in self.keys_loose:
                continue
            if ph and any(self._fecha(x) == fc for x in self.by_phone.get(ph, [])):
                self.incompletas.append((r, fila, 'teléfono ya existe con misma fecha de cita'))
                continue
            if (ph, nm, fc, camp) in vistos:
                continue
            vistos.add((ph, nm, fc, camp))
            if not nm:
                self.incompletas.append((r, fila, 'sin nombre en AGENDADOS'))
            prov = self.last_row + len(self.new_rows) + 1
            self.new_rows.append((prov, fila))

    def _indexar_nuevos(self):
        for prov, fila in self.new_rows:
            ph = norm_phone(fila.get('I'))
            nm = norm_name(fila.get('G'))
            fc = norm_fecha(fila.get('L'), fila.get('M'), fila.get('N'))
            self._nueva_info[prov] = (ph, nm, fc)
            if ph:
                self.by_phone[ph].append(prov)
            if nm:
                self.by_name[nm].append(prov)

    # ----- helpers que funcionan para filas existentes y provisionales -----
    def _valor(self, fila, letra):
        if fila in self._nueva_info:
            return None
        col = openpyxl.utils.column_index_from_string(letra)
        return self.maestro.cell(row=fila, column=col).value

    def _fecha(self, fila):
        if fila in self._nueva_info:
            return self._nueva_info[fila][2]
        return norm_fecha(self._valor(fila, 'L'), self._valor(fila, 'M'), self._valor(fila, 'N'))

    def _asistio(self, fila):
        return txt(self._valor(fila, 'P')) == 'ASISTIO'

    # ----- 2) VENTA DIARIA -> completar P..AB -----
    def _calcular_ventas(self, venta):
        ws = self.maestro
        para_llenar = defaultdict(list)  # maestro_row -> [venta_row]
        for v in venta:
            status = txt(v['status']) or ''
            st = status.upper()
            ph = norm_phone(v['cel'])
            nm = norm_name(v['nombre'])
            fv = norm_fecha(v['dia'], v['mes'], v['anio'])

            if st in REVISAR_STATUS:
                self.revisar.append((v, 'STATUS por revisar'))
                continue
            if not (st.startswith(ASISTE_POR_TEXTO) or any(x in st for x in ASISTE_POR_TEXTO)):
                self.pendientes.append((v, f'STATUS "{status}" (no asistió o desconocido)'))
                continue

            def empty_p_rows(lista):
                return [r for r in lista if not self._asistio(r)]

            modo = None
            candidatos = []
            # Etapa 1: mismo teléfono + misma fecha de cita
            if ph:
                cand = [r for r in self.by_phone[ph] if self._fecha(r) == fv]
                if any(self._asistio(r) for r in cand):
                    self.pendientes.append((v, 'ya asistió con mismo teléfono+fecha'))
                    continue
                cand = empty_p_rows(cand)
                if len(cand) == 1:
                    candidatos, modo = cand, 'telefono+fecha'
                elif len(cand) > 1:
                    self.revisar.append((v, f'{len(cand)} filas con mismo teléfono+fecha'))
            # Etapa 2: mismo nombre + misma fecha
            if not candidatos and nm:
                cand = [r for r in self.by_name[nm] if self._fecha(r) == fv]
                if any(self._asistio(r) for r in cand):
                    self.pendientes.append((v, 'ya asistió con mismo nombre+fecha'))
                    continue
                cand = empty_p_rows(cand)
                if len(cand) == 1:
                    candidatos, modo = cand, 'nombre+fecha'
                elif len(cand) > 1:
                    self.revisar.append((v, f'{len(cand)} filas con mismo nombre+fecha'))
            if not candidatos:
                self.pendientes.append((v, 'sin coincidencia en maestro'))
                continue

            fila_m = candidatos[0]
            para_llenar[fila_m].append(v)
            self.matches.append((v['fila'], fila_m, modo, v['hoja']))

        # ----- armar las celdas a escribir -----
        for fila_m, ventas in para_llenar.items():
            ventas.sort(key=lambda x: (x['anio'] or 0, str(x['mes'] or ''), x['dia'] or 0, x['fila']))
            primer = ventas[0]
            u = self.updates.setdefault(fila_m, {})
            if not self._asistio(fila_m):
                u['P'] = 'ASISTIO'
            for col, v in (('F', primer['dni']), ('Q', primer['distrito']),
                           ('R', primer['edad']), ('S', primer['sexo'])):
                if v is not None and txt(self._valor(fila_m, col)) is None:
                    u[col] = v
            par = [('T', 'U'), ('V', 'W'), ('X', 'Y'), ('Z', 'AA')]
            for i, v in enumerate(ventas[:4]):
                trat_col, pago_col = par[i]
                if txt(self._valor(fila_m, trat_col)) is None:
                    u[trat_col] = v['tratamiento']
                if num(self._valor(fila_m, pago_col)) is None and v['venta'] is not None:
                    u[pago_col] = num(v['venta'])

    def resumen(self):
        n_match_exacto = sum(1 for _, _, m, _ in self.matches if m in ('telefono+fecha', 'nombre+fecha'))
        n_sin_fecha = sum(1 for _, _, m, _ in self.matches if m == 'sin fecha' or 'sin fecha' in m)
        return {
            'filas_nuevas': len(self.new_rows),
            'ventas_totales': len(self.pendientes) + len(self.matches) + len(self.revisar),
            'matches': len(self.matches),
            'matches_exactos': n_match_exacto,
            'matches_sin_fecha': n_sin_fecha,
            'pendientes': len(self.pendientes),
            'revisar': len(self.revisar),
            'incompletas': len(self.incompletas),
            'filas_a_actualizar': len(self.updates),
            'celdas_a_escribir': sum(len(v) for v in self.updates.values()),
        }


# ============================================================
# APLICACIÓN VÍA XML (preserva tablas dinámicas, slicers, fórmulas)
# ============================================================
AB_FORMULA = ('+Tabla1[[#This Row],[PAGO 1]]+Tabla1[[#This Row],[PAGO 2]]'
              '+Tabla1[[#This Row],[PAGO 3]]+Tabla1[[#This Row],[PAGO 4]]')


def esc_xml(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def celda_xml(col, fila, valor, formula=False):
    ref = f'{col}{fila}'
    if formula:
        return f'<c r="{ref}" s="1"><f>{esc_xml(AB_FORMULA)}</f><v>0</v></c>'
    if valor is None:
        return f'<c r="{ref}" s="1"/>'
    if isinstance(valor, (int, float)):
        return f'<c r="{ref}" s="1"><v>{num(valor)}</v></c>'
    t = esc_xml(valor)
    return f'<c r="{ref}" s="1" t="inlineStr"><is><t>{t}</t></is></c>'


_CELL_RE = re.compile(
    r'<c r="([A-Z]{1,3})\d+"[^>]*?/>(?=<c|</row>|$)'
    r'|<c r="([A-Z]{1,3})\d+"[^>]*?>(?:(?!</c>).)*?</c>', re.S)
_CELL_OPEN_RE = re.compile(r'<c r="([A-Z]{1,3})\d+"[^>]*>')


def parse_row(row_xml):
    """Extrae celdas de una fila -> dict {col: xml de celda}."""
    celdas = {}
    pos = 0
    for m in _CELL_RE.finditer(row_xml):
        xml = m.group(0)
        col = _CELL_OPEN_RE.search(xml).group(1)
        celdas[col] = xml
    return celdas


_COL_ORDER = {openpyxl.utils.get_column_letter(i): i for i in range(1, 64)}


def build_row(fila, celdas):
    items = sorted(celdas.items(), key=lambda kv: _COL_ORDER[kv[0]])
    return f'<row r="{fila}" spans="2:28" x14ac:dyDescent="0.3">' + ''.join(x for _, x in items) + '</row>'


def aplicar_xml(origen, destino, new_rows, updates):
    build_dir = os.path.join(TMP_DIR, 'build')
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    os.makedirs(build_dir)
    with zipfile.ZipFile(origen) as z:
        z.extractall(build_dir)

    sheet_path = os.path.join(build_dir, 'xl', 'worksheets', 'sheet1.xml')
    with open(sheet_path, encoding='utf-8') as f:
        x = f.read()

    # ---- actualizar celdas de filas existentes ----
    row_re = re.compile(r'<row r="(\d+)"[^>]*>.*?</row>', re.S)

    def reemplazar_row(m):
        fila = int(m.group(1))
        if fila not in updates:
            return m.group(0)
        celdas = parse_row(m.group(0))
        for col, val in updates[fila].items():
            celdas[col] = celda_xml(col, fila, val)
        return build_row(fila, celdas)

    x = row_re.sub(reemplazar_row, x)

    # ---- agregar filas nuevas ----
    ultima = max(int(m.group(1)) for m in row_re.finditer(x)) or 4
    for i, (fila_num, fila) in enumerate(new_rows):
        celdas = {}
        for col, val in fila.items():
            celdas[col] = celda_xml(col, fila_num, val)
        for col, val in updates.get(fila_num, {}).items():
            celdas[col] = celda_xml(col, fila_num, val)
        celdas['AB'] = celda_xml('AB', fila_num, 0, formula=True)
        x = x.replace('</sheetData>', build_row(fila_num, celdas) + '</sheetData>', 1)
        if fila_num > ultima:
            ultima = fila_num

    # ---- actualizar tabla y dimension ----
    x = re.sub(r'<table ref="[^"]+"', f'<table ref="B4:AB{ultima}"', x)
    x = re.sub(r'<autoFilter ref="[^"]+"', f'<autoFilter ref="B4:AB{ultima}"', x)
    x = re.sub(r'<dimension ref="[^"]+"', f'<dimension ref="A1:AB{ultima}"', x)

    with open(sheet_path, 'w', encoding='utf-8') as f:
        f.write(x)

    # ---- actualizar tabla1.xml ----
    tab_path = os.path.join(build_dir, 'xl', 'tables', 'table1.xml')
    with open(tab_path, encoding='utf-8') as f:
        t = f.read()
    t = re.sub(r'ref="B4:AB\d+"', f'ref="B4:AB{ultima}"', t)
    with open(tab_path, 'w', encoding='utf-8') as f:
        f.write(t)

    # ---- re-empacar ----
    if os.path.exists(destino):
        os.remove(destino)
    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(build_dir):
            for fn in files:
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, build_dir)
                z.write(full, rel)


# ============================================================
# REPORTE
# ============================================================
def escribir_reporte(calc, ruta):
    s = []
    s.append('=' * 60)
    s.append(f'REPORTE DE ACTUALIZACION DEL MAESTRO - {datetime.now():%d/%m/%Y %H:%M}')
    s.append('=' * 60)
    r = calc.resumen()
    s.append(f'Filas nuevas (AGENDADOS):        {r["filas_nuevas"]}')
    s.append(f'Ventas procesadas:               {r["ventas_totales"]}')
    s.append(f'  con match en maestro:          {r["matches"]} (exactas {r["matches_exactos"]}, sin fecha {r["matches_sin_fecha"]})')
    s.append(f'  pendientes (sin match):        {r["pendientes"]}')
    s.append(f'  a revisar:                     {r["revisar"]}')
    s.append(f'Filas del maestro a actualizar:  {r["filas_a_actualizar"]}')
    s.append(f'Celdas a escribir:               {r["celdas_a_escribir"]}')
    s.append(f'Incompletos en AGENDADOS (no cop.): {r["incompletas"]}')
    s.append('')

    if calc.incompletas:
        s.append('--- AGENDADOS INCOMPLETOS (no copiados, revisar fuente) ---')
        for r_ag, fila, motivo in calc.incompletas[:60]:
            s.append(f'  [{motivo}] fila {r_ag}: {fila.get("C")}/{fila.get("D")}/{fila.get("E")} -> '
                     f'cita {fila.get("L")}/{fila.get("M")}/{fila.get("N")} | '
                     f'{fila.get("G") or "(sin nombre)"} | {fila.get("I")} | {fila.get("O")}')
        s.append('')

    if calc.new_rows:
        s.append('--- AGENDADOS NUEVOS ---')
        for r_ag, fila in calc.new_rows[:200]:
            s.append(f'  {fila.get("C")}/{fila.get("D")}/{fila.get("E")} -> '
                     f'cita {fila.get("L")}/{fila.get("M")}/{fila.get("N")} | '
                     f'{fila.get("G")} | {fila.get("I")} | {fila.get("O")}')

    s.append('')
    s.append('--- VENTAS PENDIENTES (sin coincidencia) ---')
    for v, motivo in calc.pendientes:
        s.append(f'  [{motivo}] {v["fila"]} {v["nombre"]} | {v["cel"]} | '
                 f'{v["dia"]}/{v["mes"]}/{v["anio"]} | {v["tratamiento"]} | S/{v["venta"]}')
    s.append('')
    s.append('--- MATCHES SIN FECHA EXACTA (revisar) ---')
    ws = calc.maestro
    for vf, mf, modo, hoja in calc.matches:
        if 'sin fecha' in modo:
            nombre_m = ws.cell(row=mf, column=7).value if mf not in calc._nueva_info else '(fila nueva)'
            fecha_m = [ws.cell(row=mf, column=c).value for c in (12, 13, 14)] if mf not in calc._nueva_info else '(nueva)'
            s.append(f'  venta {vf} -> maestro fila {mf} | {nombre_m} | cita {fecha_m}')
    s.append('')
    s.append('--- CASOS A REVISAR ---')
    for v, motivo in calc.revisar:
        s.append(f'  [{motivo}] {v["fila"]} {v["nombre"]} | {v["cel"]} | '
                 f'{v["dia"]}/{v["mes"]}/{v["anio"]} | {v["tratamiento"]} | {v["status"]}')
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write('\n'.join(s))
    return '\n'.join(s)


# ============================================================
# FLUJO PRINCIPAL (reutilizable desde CLI o web)
# ============================================================
def ejecutar_sync(aplicar=True, sin_descarga=False):
    """Ejecuta el flujo completo de sincronización.

    Devuelve un dict con el resumen, el texto del reporte y los archivos usados.
    Si aplicar=True, respalda y escribe los cambios en el maestro.
    """
    resultado = {'ok': True, 'aplicar': aplicar, 'errores': [],
                 'backup': None, 'maestro': MAESTRO,
                 'agendados': None, 'venta_diaria': None}
    try:
        if sin_descarga:
            ag = os.path.join(TMP_DIR, 'AGENDADOS.xlsx')
            ve = os.path.join(TMP_DIR, 'VENTA_DIARIA.xlsx')
        else:
            ag = descargar(AGENDADOS_FID, 'AGENDADOS', forzar=True)
            ve = descargar(VENTA_FID, 'VENTA_DIARIA', forzar=True)

        maestro_ws = leer_maestro(MAESTRO)
        agendados = leer_agendados(ag)
        venta = leer_venta(ve)

        calc = Calculo(maestro_ws, agendados, venta)
        texto = escribir_reporte(calc, os.path.join(TMP_DIR, 'reporte_actualizacion.txt'))
        resultado['reporte'] = texto
        resultado['resumen'] = calc.resumen()
        resultado['agendados'] = ag
        resultado['venta_diaria'] = ve

        if aplicar:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup = MAESTRO.replace('.xlsx', f'_backup_{ts}.xlsx')
            shutil.copy2(MAESTRO, backup)
            resultado['backup'] = backup
            aplicar_xml(MAESTRO, MAESTRO, calc.new_rows, calc.updates)
    except Exception as e:  # noqa: BLE001
        resultado['ok'] = False
        resultado['errores'].append(str(e))
    return resultado


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--aplicar', action='store_true', help='Guardar los cambios en el maestro')
    ap.add_argument('--sin-descarga', action='store_true', help='Reutilizar archivos ya descargados')
    args = ap.parse_args()

    print('1) Descargando fuentes...')
    print('2) Leyendo maestro y fuentes...')
    print('3) Calculando cambios...')

    resultado = ejecutar_sync(aplicar=args.aplicar, sin_descarga=args.sin_descarga)

    if not resultado['ok']:
        print('\nERROR:', resultado['errores'])
        sys.exit(1)

    print(resultado['reporte'])

    if not args.aplicar:
        print('\n[DRY RUN] No se modificó el maestro. Usa --aplicar para guardar.')
    else:
        print('4) Aplicando cambios (con backup)...')
        print(f'  Backup: {resultado["backup"]}')
        print(f'  Maestro actualizado: {MAESTRO}')


if __name__ == '__main__':
    main()
