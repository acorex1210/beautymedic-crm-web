# -*- coding: utf-8 -*-
"""crm_plus.py
============
Funciones estilo CRM (tipo KOMMO) sobre Google Drive: pipeline kanban,
tareas/recordatorios y notas de seguimiento por paciente.

Guarda todo en un workbook propio ``CRM.xlsx`` (hojas TARJETAS, TAREAS,
NOTAS, CUOTAS, CAJA, PRODUCTOS, MOVIMIENTOS_STOCK, TRABAJADORES y PLANILLA)
que se busca por nombre en Drive. Además agrega vistas derivadas: directorio
de pacientes, dashboard de métricas, panel de actividades de hoy (combinando
AGENDADOS y VENTA DIARIA), el catálogo/kardex de inventario y la planilla
quincenal de sueldos.
"""
import calendar
import io
import json
import os
import re
import threading
import time
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

import openpyxl

import alimentar_maestro as am  # noqa: E402
import crm_drive as cd          # noqa: E402

TZ = ZoneInfo('America/Lima')
NOMBRE_ARCHIVO = os.environ.get('CRM_ARCHIVO', 'CRM.xlsx')

ETAPAS = ['NUEVO', 'AGENDADO', 'CONFIRMADO', 'ATENDIDO', 'GANADO', 'PERDIDO']
TIPO_TAREA = ['LLAMADA', 'SEGUIMIENTO', 'RECORDATORIO', 'OTRO']
ESTADOS_TAREA = ['PENDIENTE', 'COMPLETADA', 'CANCELADA']
PRIORIDADES = ['ALTA', 'MEDIA', 'BAJA']
TIPO_NOTA = ['LLAMADA', 'WHATSAPP', 'INBOX', 'OTRO']

_MM = ['', 'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO',
       'SET', 'OCT', 'NOV', 'DIC']
_MM_IDX = {m: i for i, m in enumerate(_MM) if m}

HOJAS = {
    'TARJETAS': ['ID', 'NOMBRE', 'TELEFONO', 'ETAPA', 'CRM', 'CAMPANA', 'VALOR',
                 'PRIORIDAD', 'FECHA_ALTA', 'FECHA_ULT', 'CITA_DIA', 'CITA_MES',
                 'CITA_HORA', 'NOTA'],
    'TAREAS': ['ID', 'TITULO', 'TIPO', 'FECHA', 'HORA', 'ESTADO', 'PRIORIDAD',
               'CONTACTO', 'NOTA'],
    'NOTAS': ['ID', 'TELEFONO', 'CONTACTO', 'FECHA', 'TIPO', 'TEXTO'],
    'CUOTAS': ['ID', 'PACIENTE', 'TELEFONO', 'TRATAMIENTO', 'MONTO_TOTAL',
               'N_CUOTAS', 'PAGADAS', 'MONTO_CUOTA', 'PROX_FECHA', 'ESTADO',
               'NOTA'],
    'CAJA': ['ID', 'FECHA', 'HORA', 'TIPO', 'MONTO', 'CONCEPTO', 'RESPONSABLE',
             'DESGLOSE', 'MONTO_USD', 'DESGLOSE_USD'],
    'PRODUCTOS': ['ID', 'CODIGO', 'PRODUCTO', 'COSTO_BRUTO', 'STOCK',
                  'STOCK_MINIMO', 'UNIDAD', 'FECHA_ALTA', 'ACTUALIZADO'],
    'MOVIMIENTOS_STOCK': ['ID', 'FECHA', 'CODIGO', 'PRODUCTO', 'TIPO', 'CANTIDAD',
                          'STOCK_RESULTANTE', 'REFERENCIA', 'NOTA'],
    'TRABAJADORES': ['ID', 'NOMBRE', 'CARGO', 'SUELDO_QUINCENA', 'PORCENTAJE_COMISION',
                     'METODO_PAGO', 'CUENTA', 'ESTADO', 'FECHA_ALTA'],
    'PLANILLA': ['ID', 'TRABAJADOR_ID', 'NOMBRE', 'ANIO', 'MES', 'QUINCENA',
                'SUELDO_BASE', 'COMISION', 'MONTO_TOTAL', 'ESTADO', 'FECHA_PAGO',
                'METODO_PAGO', 'NOTA'],
}

TARJETA_COLS = {'id': 'A', 'nombre': 'B', 'telefono': 'C', 'etapa': 'D',
                'crm': 'E', 'campana': 'F', 'valor': 'G', 'prioridad': 'H',
                'fecha_alta': 'I', 'fecha_ult': 'J', 'cita_dia': 'K',
                'cita_mes': 'L', 'cita_hora': 'M', 'nota': 'N'}
TAREA_COLS = {'id': 'A', 'titulo': 'B', 'tipo': 'C', 'fecha': 'D', 'hora': 'E',
              'estado': 'F', 'prioridad': 'G', 'contacto': 'H', 'nota': 'I'}
NOTA_COLS = {'id': 'A', 'telefono': 'B', 'contacto': 'C', 'fecha': 'D',
             'tipo': 'E', 'texto': 'F'}
CUOTA_COLS = {'id': 'A', 'paciente': 'B', 'telefono': 'C', 'tratamiento': 'D',
              'monto_total': 'E', 'n_cuotas': 'F', 'pagadas': 'G',
              'monto_cuota': 'H', 'prox_fecha': 'I', 'estado': 'J', 'nota': 'K'}
CAJA_COLS = {'id': 'A', 'fecha': 'B', 'hora': 'C', 'tipo': 'D', 'monto': 'E',
             'concepto': 'F', 'responsable': 'G', 'desglose': 'H',
             'monto_usd': 'I', 'desglose_usd': 'J'}
TIPO_CAJA = ['APERTURA', 'INGRESO', 'EGRESO', 'CIERRE']
PRODUCTO_COLS = {'id': 'A', 'codigo': 'B', 'producto': 'C', 'costo_bruto': 'D',
                 'stock': 'E', 'stock_minimo': 'F', 'unidad': 'G',
                 'fecha_alta': 'H', 'actualizado': 'I'}
MOVIMIENTO_COLS = {'id': 'A', 'fecha': 'B', 'codigo': 'C', 'producto': 'D',
                   'tipo': 'E', 'cantidad': 'F', 'stock_resultante': 'G',
                   'referencia': 'H', 'nota': 'I'}
TIPO_MOVIMIENTO = ['ENTRADA', 'SALIDA', 'AJUSTE']
TRABAJADOR_COLS = {'id': 'A', 'nombre': 'B', 'cargo': 'C', 'sueldo_quincena': 'D',
                   'porcentaje_comision': 'E', 'metodo_pago': 'F', 'cuenta': 'G',
                   'estado': 'H', 'fecha_alta': 'I'}
ESTADO_TRABAJADOR = ['ACTIVO', 'INACTIVO']
PLANILLA_COLS = {'id': 'A', 'trabajador_id': 'B', 'nombre': 'C', 'anio': 'D',
                 'mes': 'E', 'quincena': 'F', 'sueldo_base': 'G', 'comision': 'H',
                 'monto_total': 'I', 'estado': 'J', 'fecha_pago': 'K',
                 'metodo_pago': 'L', 'nota': 'M'}
ESTADO_PAGO = ['PENDIENTE', 'PAGADO']
# Denominaciones del arqueo (formato de recepción/caja)
BILLETES_SOLES = [200, 100, 50, 20, 10]
MONEDAS_SOLES = [5, 2, 1, 0.5, 0.2, 0.1]
DENOMINACIONES_CAJA = BILLETES_SOLES + MONEDAS_SOLES
DENOMINACIONES_USD = [100, 50, 20, 10]

_lock = threading.RLock()


def _ahora():
    return datetime.now(TZ).strftime('%d/%m/%Y %H:%M')


def _hoy():
    hoy = datetime.now(TZ)
    return hoy.day, _MM[hoy.month], hoy.year


def _fecha_hoy_str():
    dia, mes, anio = _hoy()
    return f'{dia}/{mes}/{anio}'


def _fecha_orden(fecha):
    try:
        d, m, a = str(fecha).split('/')
        return (int(a), _MM_IDX.get(m, 0), int(d))
    except (ValueError, AttributeError):
        return (0, 0, 0)


# ============================================================
# WORKBOOK CRM.xlsx (Drive)
# ============================================================
def _crear_wb():
    wb = openpyxl.Workbook()
    for i, nombre in enumerate(HOJAS):
        ws = wb.active if i == 0 else wb.create_sheet(nombre)
        ws.title = nombre
        for c, col in enumerate(HOJAS[nombre], 1):
            celda = ws.cell(row=1, column=c, value=col)
            celda.font = openpyxl.styles.Font(bold=True, color='FFFFFF')
            celda.fill = openpyxl.styles.PatternFill('solid', fgColor='1F3864')
        ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(HOJAS[nombre]))}1'
        ws.freeze_panes = 'A2'
    return wb


def _fid():
    drv = cd._drive()
    res = drv.files().list(
        q=f"name='{NOMBRE_ARCHIVO}' and trashed=false",
        fields='files(id)', pageSize=10).execute()
    archivos = res.get('files', [])
    if archivos:
        return archivos[0]['id']
    # La cuenta de servicio no puede crear archivos (no tiene cuota de
    # almacenamiento): pedimos un archivo vacío compartido con el bot.
    try:
        correo = cd._credenciales().service_account_email
    except Exception:  # noqa: BLE001
        correo = 'la cuenta de servicio'
    raise RuntimeError(
        f'No se encontró el archivo "{NOMBRE_ARCHIVO}" en Google Drive. '
        f'Crea un archivo Excel vacío con ese nombre y compártelo con '
        f'{correo} (edición) para habilitar el CRM.')


def _bajar():
    """Descarga CRM.xlsx de Drive. Si es un Google Sheets, lo exporta a xlsx."""
    ruta = os.path.join(am.TMP_DIR, NOMBRE_ARCHIVO)
    if os.path.exists(ruta) and time.time() - os.path.getmtime(ruta) < cd.CACHE_TTL:
        return ruta
    fid = _fid()
    drv = cd._drive()
    meta = drv.files().get(fileId=fid, fields='mimeType').execute()
    if meta.get('mimeType') == cd.MIME_XLSX:
        req = drv.files().get_media(fileId=fid)
    else:
        req = drv.files().export(fileId=fid, mimeType=cd.MIME_XLSX)
    with open(ruta, 'wb') as fh:
        dl = cd.MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
    return ruta


def _leer_hoja(nombre):
    ruta = _bajar()
    wb = openpyxl.load_workbook(ruta, data_only=True)
    if nombre not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[nombre]
    cols = [c for c in range(1, ws.max_column + 1)
            if ws.cell(row=1, column=c).value is not None]
    filas = []
    for r in range(2, ws.max_row + 1):
        f = {}
        for c in cols:
            v = ws.cell(row=r, column=c).value
            if v is None or v == '':
                continue
            f[openpyxl.utils.get_column_letter(c)] = v
        if f:
            filas.append(f)
    wb.close()
    return filas


def _guardar(filas_por_hoja):
    with _lock:
        wb = _crear_wb()
        for nombre, filas in filas_por_hoja.items():
            ws = wb[nombre]
            for i, f in enumerate(filas, 2):
                for k, v in f.items():
                    if v is not None and v != '':
                        ws.cell(row=i, column=openpyxl.utils.column_index_from_string(k),
                                value=v)
        buf = io.BytesIO()
        wb.save(buf)
        wb.close()
        buf.seek(0)
        destino = os.path.join(am.TMP_DIR,
                               os.path.splitext(NOMBRE_ARCHIVO)[0] + '_editado.xlsx')
        with open(destino, 'wb') as fh:
            fh.write(buf.read())
        cd.subir_drive(_fid(), destino)
        cd.invalidar(os.path.splitext(NOMBRE_ARCHIVO)[0])


def _reescribir(hoja, filas):
    with _lock:
        todo = {n: _leer_hoja(n) for n in HOJAS}
        todo[hoja] = filas
        _guardar(todo)


def _siguiente_id(filas):
    return max([am.num(f.get('A')) or 0 for f in filas] + [0]) + 1


# ============================================================
# TARJETAS (pipeline)
# ============================================================
def _aplicar_tarjeta(f, d):
    def pon(campo, letra):
        if campo in d and d[campo] not in (None, ''):
            f[letra] = str(d[campo]).strip()
    if 'telefono' in d:
        tel = re.sub(r'\D', '', str(d.get('telefono') or ''))
        if tel:
            f['C'] = tel
    if 'etapa' in d:
        etapa = str(d.get('etapa') or '').upper().strip()
        if etapa in ETAPAS:
            f['D'] = etapa
    if 'prioridad' in d:
        p = str(d.get('prioridad') or '').upper().strip()
        if p in PRIORIDADES:
            f['H'] = p
    if 'valor' in d and am.num(d.get('valor')) is not None:
        f['G'] = am.num(d['valor'])
    if 'cita_dia' in d and am.num(d.get('cita_dia')) is not None:
        f['K'] = am.num(d['cita_dia'])
    if 'cita_mes' in d:
        m = cd._mes(d.get('cita_mes'))
        if m:
            f['L'] = m
    for campo, letra in (('nombre', 'B'), ('crm', 'E'), ('campana', 'F'),
                         ('cita_hora', 'M'), ('nota', 'N')):
        pon(campo, letra)


def _tarjeta_named(f):
    return {
        'id': am.num(f.get('A')), 'nombre': f.get('B'), 'telefono': f.get('C'),
        'etapa': f.get('D') or 'NUEVO', 'crm': f.get('E'), 'campana': f.get('F'),
        'valor': am.num(f.get('G')), 'prioridad': f.get('H') or 'MEDIA',
        'fecha_alta': f.get('I'), 'fecha_ult': f.get('J'),
        'cita_dia': am.num(f.get('K')), 'cita_mes': f.get('L'),
        'cita_hora': f.get('M'), 'nota': f.get('N'),
    }


def leer_tarjetas():
    return [_tarjeta_named(f) for f in _leer_hoja('TARJETAS')]


def crear_tarjeta(datos):
    with _lock:
        filas = _leer_hoja('TARJETAS')
        f = {'A': _siguiente_id(filas), 'I': _ahora(), 'J': _ahora()}
        _aplicar_tarjeta(f, datos)
        f['D'] = f.get('D') or 'NUEVO'
        f['H'] = f.get('H') or 'MEDIA'
        filas.append(f)
        _reescribir('TARJETAS', filas)
        return _tarjeta_named(f)


def actualizar_tarjeta(tid, cambios):
    with _lock:
        filas = _leer_hoja('TARJETAS')
        for f in filas:
            if am.num(f.get('A')) == tid:
                _aplicar_tarjeta(f, cambios)
                f['J'] = _ahora()
                _reescribir('TARJETAS', filas)
                return _tarjeta_named(f)
    return None


def borrar_tarjeta(tid):
    with _lock:
        filas = _leer_hoja('TARJETAS')
        nuevas = [f for f in filas if am.num(f.get('A')) != tid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('TARJETAS', nuevas)
        return True


# ============================================================
# TAREAS
# ============================================================
def _aplicar_tarea(f, d):
    def pon(campo, letra):
        if campo in d and d[campo] not in (None, ''):
            f[letra] = str(d[campo]).strip()
    if 'tipo' in d:
        t = str(d.get('tipo') or '').upper().strip()
        if t in TIPO_TAREA:
            f['C'] = t
    if 'estado' in d:
        e = str(d.get('estado') or '').upper().strip()
        if e in ESTADOS_TAREA:
            f['F'] = e
    if 'prioridad' in d:
        p = str(d.get('prioridad') or '').upper().strip()
        if p in PRIORIDADES:
            f['G'] = p
    for campo, letra in (('titulo', 'B'), ('fecha', 'D'), ('hora', 'E'),
                         ('contacto', 'H'), ('nota', 'I')):
        pon(campo, letra)


def _tarea_named(f):
    return {
        'id': am.num(f.get('A')), 'titulo': f.get('B'), 'tipo': f.get('C'),
        'fecha': f.get('D'), 'hora': f.get('E'), 'estado': f.get('F') or 'PENDIENTE',
        'prioridad': f.get('G') or 'MEDIA', 'contacto': f.get('H'),
        'nota': f.get('I'),
    }


def leer_tareas(estado=None):
    filas = _leer_hoja('TAREAS')
    out = [_tarea_named(f) for f in filas]
    if estado:
        out = [t for t in out if t['estado'] == estado]
    out.sort(key=lambda t: (t['fecha'] or '9999', t['hora'] or '99'))
    return out


def crear_tarea(datos):
    with _lock:
        filas = _leer_hoja('TAREAS')
        f = {'A': _siguiente_id(filas)}
        _aplicar_tarea(f, datos)
        f['C'] = f.get('C') or 'OTRO'
        f['F'] = f.get('F') or 'PENDIENTE'
        f['G'] = f.get('G') or 'MEDIA'
        filas.append(f)
        _reescribir('TAREAS', filas)
        return _tarea_named(f)


def actualizar_tarea(tid, cambios):
    with _lock:
        filas = _leer_hoja('TAREAS')
        for f in filas:
            if am.num(f.get('A')) == tid:
                _aplicar_tarea(f, cambios)
                _reescribir('TAREAS', filas)
                return _tarea_named(f)
    return None


def borrar_tarea(tid):
    with _lock:
        filas = _leer_hoja('TAREAS')
        nuevas = [f for f in filas if am.num(f.get('A')) != tid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('TAREAS', nuevas)
        return True


# ============================================================
# NOTAS de seguimiento
# ============================================================
def _aplicar_nota(f, d):
    if d.get('telefono'):
        tel = re.sub(r'\D', '', str(d['telefono']))
        if tel:
            f['B'] = tel
    if d.get('contacto') not in (None, ''):
        f['C'] = str(d['contacto']).strip()
    if 'tipo' in d:
        t = str(d.get('tipo') or '').upper().strip()
        if t in TIPO_NOTA:
            f['E'] = t
    if d.get('texto') not in (None, ''):
        f['F'] = str(d['texto']).strip()


def _nota_named(f):
    return {
        'id': am.num(f.get('A')), 'telefono': f.get('B'), 'contacto': f.get('C'),
        'fecha': f.get('D'), 'tipo': f.get('E'), 'texto': f.get('F'),
    }


def leer_notas(telefono=None, contacto=None):
    filas = _leer_hoja('NOTAS')
    out = [_nota_named(f) for f in filas]
    if telefono:
        tel = re.sub(r'\D', '', str(telefono))
        out = [n for n in out if n['telefono'] and tel in str(n['telefono'])]
    elif contacto:
        c = str(contacto).strip().lower()
        out = [n for n in out if c and c in str(n.get('contacto') or '').lower()]
    out.sort(key=lambda n: n['fecha'] or '', reverse=True)
    return out


def agregar_nota(datos):
    with _lock:
        filas = _leer_hoja('NOTAS')
        f = {'A': _siguiente_id(filas), 'D': _ahora()}
        _aplicar_nota(f, datos)
        f['E'] = f.get('E') or 'OTRO'
        filas.append(f)
        _reescribir('NOTAS', filas)
        return _nota_named(f)


# ============================================================
# DIRECTORIO DE PACIENTES (agregado)
# ============================================================
def _clave(tel, nombre):
    if tel:
        return 't' + tel
    return 'n' + str(nombre or '').strip().lower()


def leer_pacientes():
    ag = cd.leer_agendados()['filas']
    ve = [f for v in cd.leer_venta()['hojas'].values() for f in v['filas']]
    tar = leer_tarjetas()
    notas = _leer_hoja('NOTAS')

    pacientes = {}

    def base(nombre, tel):
        k = _clave(tel, nombre)
        p = pacientes.get(k)
        if not p:
            p = {'clave': k, 'nombre': nombre, 'telefono': tel, 'correo': '',
                 'crm': '', 'campana': '', 'etapa': None, 'citas': 0,
                 'compras': 0, 'total': 0.0, 'proxima_cita': None,
                 'ultima_actividad': None, 'notas': 0, 'doctores': ''}
            pacientes[k] = p
        return p

    for f in ag:
        tel = f.get('I')
        tel = re.sub(r'\D', '', str(tel)) if tel is not None else ''
        p = base(f.get('G'), tel)
        if not p['crm'] and f.get('B'):
            p['crm'] = str(f['B']).strip()
        if not p['campana'] and f.get('O'):
            p['campana'] = str(f['O']).strip()
        if not p['correo'] and f.get('J'):
            p['correo'] = str(f['J']).strip()
        p['citas'] += 1
        if f.get('L') and f.get('M') and f.get('N'):
            fecha = f'{f["L"]}/{f["M"]}/{f["N"]}'
            hora = f.get('P') or ''
            if p['proxima_cita'] is None or fecha > p['proxima_cita'].split(' ')[0]:
                p['proxima_cita'] = f'{fecha} {hora}'.strip()
            if p['ultima_actividad'] is None or fecha > p['ultima_actividad'].split(' ')[0]:
                p['ultima_actividad'] = f'{fecha} {hora}'.strip()

    for f in ve:
        tel = f.get('F')
        tel = re.sub(r'\D', '', str(tel)) if tel is not None else ''
        p = base(f.get('G'), tel)
        if not p['crm'] and f.get('H'):
            p['crm'] = str(f['H']).strip()
        if not p['campana'] and f.get('L'):
            p['campana'] = str(f['L']).strip()
        if f.get('O'):
            p['compras'] += 1
            p['total'] = (p['total'] or 0) + (am.num(f['O']) or 0)
        if f.get('M') and str(f['M']).strip() not in p['doctores']:
            p['doctores'] = (p['doctores'] + ' · ' if p['doctores'] else '') + str(f['M']).strip()
        if f.get('B') and f.get('C') and f.get('D'):
            fecha = f'{f["B"]}/{f["C"]}/{f["D"]}'
            if p['ultima_actividad'] is None or fecha > p['ultima_actividad'].split(' ')[0]:
                p['ultima_actividad'] = fecha

    for t in tar:
        tel = re.sub(r'\D', '', str(t.get('telefono') or ''))
        p = base(t.get('nombre'), tel)
        if not p['crm'] and t.get('crm'):
            p['crm'] = str(t['crm']).strip()
        if not p['campana'] and t.get('campana'):
            p['campana'] = str(t['campana']).strip()
        if t.get('etapa'):
            p['etapa'] = t['etapa']
        if t.get('fecha_ult') and (p['ultima_actividad'] is None
                                   or t['fecha_ult'] > p['ultima_actividad']):
            p['ultima_actividad'] = t['fecha_ult']

    for f in notas:
        tel = f.get('B')
        tel = re.sub(r'\D', '', str(tel)) if tel is not None else ''
        p = base(f.get('C'), tel)
        p['notas'] += 1
        if f.get('D') and (p['ultima_actividad'] is None
                           or f['D'] > p['ultima_actividad']):
            p['ultima_actividad'] = f['D']

    out = list(pacientes.values())
    out.sort(key=lambda p: (p['ultima_actividad'] or ''), reverse=True)
    return out


# ============================================================
# DASHBOARD
# ============================================================
def _status_normalizado(v):
    s = unicodedata.normalize('NFD', str(v or ''))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.strip().upper()


def _es_no_realizado(v):
    s = _status_normalizado(v)
    return 'NO SE REALIZO' in s or 'NO ASISTIO' in s


def _es_venta_registrada(v):
    """Replica esVentaRegistrada() del frontend: cuenta como venta un status
    REALIZO/COMPRO/COMPLETA/SESION/DEJO PAGADO, salvo que sea NO SE REALIZO
    o NO ASISTIO."""
    s = _status_normalizado(v)
    if _es_no_realizado(v):
        return False
    return any(x in s for x in ('REALIZO', 'COMPRO', 'COMPLETA', 'SESION', 'DEJO PAGADO'))


def leer_dashboard():
    tar = leer_tarjetas()
    ag = cd.leer_agendados()['filas']
    ve = cd.leer_venta()['hojas']

    funnel = {e: 0 for e in ETAPAS}
    for t in tar:
        funnel[t['etapa']] = funnel.get(t['etapa'], 0) + 1

    ventas_doctor = {}
    ventas_mes = {}
    total_ventas = 0
    ventas_registradas = 0
    for v in ve.values():
        for f in v['filas']:
            if not _es_venta_registrada(f.get('N')):
                continue
            monto = am.num(f.get('O')) or 0
            ventas_registradas += 1
            if monto:
                total_ventas += monto
                doc = str(f.get('M') or '').strip() or 'Sin doctor'
                ventas_doctor[doc] = ventas_doctor.get(doc, 0) + monto
                mes = str(f.get('C') or '').strip().upper()
                anio = am.num(f.get('D')) or datetime.now(TZ).year
                if mes:
                    k = f'{mes} {anio}'
                    ventas_mes[k] = ventas_mes.get(k, 0) + monto

    ag_por_crm = {}
    ag_por_campana = {}
    for f in ag:
        crm = str(f.get('B') or '').strip() or 'Sin CRM'
        camp = str(f.get('O') or '').strip() or 'Sin campaña'
        ag_por_crm[crm] = ag_por_crm.get(crm, 0) + 1
        ag_por_campana[camp] = ag_por_campana.get(camp, 0) + 1

    leads = sum(funnel.get(e, 0) for e in ('NUEVO', 'AGENDADO', 'CONFIRMADO',
                                            'ATENDIDO'))
    ganados = funnel.get('GANADO', 0)
    conversion = round(100 * ganados / leads, 1) if leads else 0

    return {
        'funnel': funnel,
        'conversion': conversion,
        'ganados': ganados,
        'perdidos': funnel.get('PERDIDO', 0),
        'total_ventas': round(total_ventas, 2),
        'ventas_registradas': ventas_registradas,
        'ventas_doctor': dict(sorted(ventas_doctor.items(),
                                     key=lambda kv: kv[1], reverse=True)),
        'ventas_mes': dict(sorted(ventas_mes.items(),
                                  key=lambda kv: kv[1], reverse=True)),
        'ag_por_crm': dict(sorted(ag_por_crm.items(),
                                  key=lambda kv: kv[1], reverse=True)),
        'ag_por_campana': dict(sorted(ag_por_campana.items(),
                                      key=lambda kv: kv[1], reverse=True)),
    }


# ============================================================
# ACTIVIDADES DE HOY
# ============================================================
def leer_hoy():
    dia, mes, anio = _hoy()
    clave_hoy = f'{dia}/{mes}/{anio}'
    hoy_iso = datetime.now(TZ).strftime('%Y-%m-%d')

    citas = []
    for f in cd.leer_agendados()['filas']:
        if not (f.get('L') and f.get('M') and f.get('N')):
            continue
        if f'{f["L"]}/{f["M"]}/{f["N"]}' != clave_hoy:
            continue
        st = str(f.get('Q') or '').upper()
        if 'NO ASISTIO' in st or 'CANCEL' in st or 'NO CONFIRM' in st:
            continue
        citas.append({'nombre': f.get('G'), 'telefono': f.get('I'),
                      'hora': f.get('P'), 'crm': f.get('B'),
                      'campana': f.get('O'), 'estado': f.get('Q'),
                      'dia': f.get('L'), 'mes': f.get('M'),
                      'fila': f.get('_fila')})
    citas.sort(key=lambda c: c['hora'] or '99')

    ventas = []
    for v in cd.leer_venta()['hojas'].values():
        for f in v['filas']:
            if not (f.get('B') and f.get('C') and f.get('D')):
                continue
            if f'{f["B"]}/{f["C"]}/{f["D"]}' != clave_hoy:
                continue
            ventas.append({'nombre': f.get('G'), 'telefono': f.get('F'),
                           'doctor': f.get('M'), 'status': f.get('N'),
                           'monto': am.num(f.get('O'))})
    ventas.sort(key=lambda v: v['monto'] or 0, reverse=True)

    tareas_hoy = []
    tareas_vencidas = []
    for t in leer_tareas():
        if t['estado'] != 'PENDIENTE':
            continue
        if t['fecha'] == hoy_iso:
            tareas_hoy.append(t)
        elif t['fecha'] and t['fecha'] < hoy_iso:
            tareas_vencidas.append(t)

    return {
        'fecha': clave_hoy,
        'citas': citas,
        'ventas': ventas,
        'tareas_hoy': tareas_hoy,
        'tareas_vencidas': tareas_vencidas,
    }


# ============================================================
# CRM Plus: Cuotas / pagos a plazos
# ============================================================
def _cuota_named(f):
    return {
        'id': am.num(f.get('A')),
        'paciente': f.get('B'),
        'telefono': f.get('C'),
        'tratamiento': f.get('D'),
        'monto_total': am.num(f.get('E')),
        'n_cuotas': am.num(f.get('F')),
        'pagadas': am.num(f.get('G')),
        'monto_cuota': am.num(f.get('H')),
        'prox_fecha': f.get('I'),
        'estado': f.get('J') or 'PENDIENTE',
        'nota': f.get('K'),
    }


def _aplicar_cuota(f, d):
    for campo, letra in (('paciente', 'B'), ('telefono', 'C'),
                         ('tratamiento', 'D'), ('prox_fecha', 'I'),
                         ('nota', 'K')):
        if campo in d and d[campo] not in (None, ''):
            f[letra] = str(d[campo]).strip()
    for campo, letra in (('monto_total', 'E'), ('n_cuotas', 'F'),
                         ('pagadas', 'G'), ('monto_cuota', 'H')):
        if campo in d and d[campo] not in (None, ''):
            try:
                f[letra] = float(d[campo])
            except (TypeError, ValueError):
                pass
    if 'estado' in d:
        e = str(d.get('estado') or '').upper().strip()
        if e in ('PENDIENTE', 'PAGADO', 'ATRASADO'):
            f['J'] = e


def _estado_cuota(cuota):
    """Recalcula estado según cuotas pagadas y fecha de pago."""
    n = cuota['n_cuotas'] or 0
    pag = cuota['pagadas'] or 0
    if n and pag >= n:
        return 'PAGADO'
    try:
        prox = datetime.strptime(cuota['prox_fecha'], '%d/%m/%Y').date()
    except (TypeError, ValueError):
        prox = None
    if prox is not None and prox < datetime.now(TZ).date():
        return 'ATRASADO'
    return 'PENDIENTE'


def leer_cuotas(estado=None, telefono=None):
    filas = _leer_hoja('CUOTAS')
    out = []
    for f in filas:
        c = _cuota_named(f)
        if not c['pagadas']:
            c['pagadas'] = 0
        if not c['n_cuotas']:
            c['n_cuotas'] = 0
        c['saldo'] = (c['monto_total'] or 0) - (c['monto_cuota'] or 0) * c['pagadas']
        c['estado'] = _estado_cuota(c)
        out.append(c)
    if estado:
        e = str(estado).upper().strip()
        out = [c for c in out if c['estado'] == e]
    if telefono:
        tel = re.sub(r'\D', '', str(telefono))
        out = [c for c in out if c['telefono'] and tel in str(c['telefono'])]
    out.sort(key=lambda c: c['estado'] != 'ATRASADO')  # atrasados primero
    return out


def crear_cuota(datos):
    with _lock:
        filas = _leer_hoja('CUOTAS')
        f = {'A': _siguiente_id(filas)}
        _aplicar_cuota(f, datos)
        f['F'] = f.get('F') or 1
        f['G'] = f.get('G') or 0
        if not f.get('H') and f.get('E') and f['F']:
            f['H'] = float(f['E']) / float(f['F'])
        f['J'] = f.get('J') or 'PENDIENTE'
        filas.append(f)
        _reescribir('CUOTAS', filas)
        c = _cuota_named(f)
        c['saldo'] = (c['monto_total'] or 0) - (c['monto_cuota'] or 0) * (c['pagadas'] or 0)
        c['estado'] = _estado_cuota(c)
        return c


def registrar_pago_cuota(cid):
    """Marca una cuota adicional como pagada."""
    with _lock:
        filas = _leer_hoja('CUOTAS')
        for f in filas:
            if am.num(f.get('A')) == cid:
                pagadas = am.num(f.get('G')) or 0
                n = am.num(f.get('F')) or 0
                if pagadas >= n:
                    return {'error': 'La cuota ya está pagada'}
                f['G'] = pagadas + 1
                c = _cuota_named(f)
                c['saldo'] = (c['monto_total'] or 0) - (c['monto_cuota'] or 0) * (pagadas + 1)
                c['estado'] = _estado_cuota(c)
                _reescribir('CUOTAS', filas)
                return c
    return None


def actualizar_cuota(cid, cambios):
    with _lock:
        filas = _leer_hoja('CUOTAS')
        for f in filas:
            if am.num(f.get('A')) == cid:
                _aplicar_cuota(f, cambios)
                _reescribir('CUOTAS', filas)
                return _cuota_named(f)
    return None


def borrar_cuota(cid):
    with _lock:
        filas = _leer_hoja('CUOTAS')
        nuevas = [f for f in filas if am.num(f.get('A')) != cid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('CUOTAS', nuevas)
        return True


# ============================================================
# CRM Plus: Caja (apertura / movimientos / cierre del día)
# ============================================================
def _json_col(f, letra):
    crudo = f.get(letra)
    if not crudo:
        return None
    try:
        return json.loads(crudo)
    except (TypeError, ValueError):
        return None


def _caja_named(f):
    return {
        'id': am.num(f.get('A')),
        'fecha': f.get('B'),
        'hora': f.get('C'),
        'tipo': f.get('D'),
        'monto': am.num(f.get('E')) or 0,
        'concepto': f.get('F'),
        'responsable': f.get('G'),
        'desglose': _json_col(f, 'H'),
        'monto_usd': am.num(f.get('I')) or 0,
        'desglose_usd': _json_col(f, 'J'),
    }


def leer_caja(fecha=None):
    filas = [_caja_named(f) for f in _leer_hoja('CAJA')]
    if fecha:
        filas = [f for f in filas if f['fecha'] == fecha]
    return sorted(filas, key=lambda f: (f['fecha'] or '', f['hora'] or ''))


def _total_desglose(desglose, denominaciones):
    """Suma cantidad x valor para cada denominación."""
    total = 0.0
    for valor in denominaciones:
        cant = am.num((desglose or {}).get(str(valor))) or 0
        total += cant * valor
    return round(total, 2)


def _monto_desde_desglose(desglose):
    return _total_desglose(desglose, DENOMINACIONES_CAJA)


def _monto_usd_desde_desglose(desglose):
    return _total_desglose(desglose, DENOMINACIONES_USD)


def _crear_movimiento_caja(tipo, datos):
    with _lock:
        filas = _leer_hoja('CAJA')
        ahora = datetime.now(TZ)
        desglose = datos.get('desglose') or None
        desglose_usd = datos.get('desglose_usd') or None
        monto = (_monto_desde_desglose(desglose) if desglose
                 else (am.num(datos.get('monto')) or 0))
        monto_usd = (_monto_usd_desde_desglose(desglose_usd) if desglose_usd
                     else (am.num(datos.get('monto_usd')) or 0))
        f = {
            'A': _siguiente_id(filas),
            'B': datos.get('fecha') or _fecha_hoy_str(),
            'C': ahora.strftime('%H:%M'),
            'D': tipo,
            'E': monto,
            'F': str(datos.get('concepto') or '').strip(),
            'G': str(datos.get('responsable') or '').strip(),
            'H': json.dumps(desglose, ensure_ascii=False) if desglose else '',
            'I': monto_usd,
            'J': json.dumps(desglose_usd, ensure_ascii=False) if desglose_usd else '',
        }
        filas.append(f)
        _reescribir('CAJA', filas)
        return _caja_named(f)


def abrir_caja(datos):
    return _crear_movimiento_caja('APERTURA', datos)


def cerrar_caja(datos):
    return _crear_movimiento_caja('CIERRE', datos)


def registrar_movimiento_caja(datos):
    tipo = str(datos.get('tipo') or '').upper().strip()
    if tipo not in ('INGRESO', 'EGRESO'):
        raise ValueError("El tipo de movimiento debe ser 'INGRESO' o 'EGRESO'")
    return _crear_movimiento_caja(tipo, datos)


def borrar_movimiento_caja(cid):
    with _lock:
        filas = _leer_hoja('CAJA')
        nuevas = [f for f in filas if am.num(f.get('A')) != cid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('CAJA', nuevas)
        return True


def _metodo_pago(valor):
    """Clasifica el texto de la columna PAGO en las categorías del arqueo."""
    p = str(valor or '').strip().upper()
    if not p:
        return 'sin_dato'
    if 'EFECTIVO' in p:
        return 'efectivo'
    if 'YAPE' in p or 'PLIN' in p:
        return 'yape_plin'
    if 'TARJETA' in p or p.startswith('T.') or 'DEBITO' in p or 'CREDITO' in p or 'POS' in p:
        return 'tarjeta'
    if 'TRANSF' in p or 'DEPOSITO' in p or 'DEPÓSITO' in p or 'LINK' in p:
        return 'deposito'
    return 'otros'


def ventas_del_dia(fecha):
    """Ventas de VENTA DIARIA para la fecha 'D/MES/AAAA', separadas por forma de pago."""
    res = {'efectivo': 0.0, 'tarjeta': 0.0, 'yape_plin': 0.0,
           'deposito': 0.0, 'otros': 0.0, 'sin_dato': 0.0}
    for v in cd.leer_venta()['hojas'].values():
        for f in v['filas']:
            if not (f.get('B') and f.get('C') and f.get('D')):
                continue
            if f'{f["B"]}/{f["C"]}/{f["D"]}' != fecha:
                continue
            monto = am.num(f.get('O')) or 0
            if not monto:
                continue
            res[_metodo_pago(f.get('P'))] += monto
    res = {k: round(v, 2) for k, v in res.items()}
    res['total'] = round(sum(res.values()), 2)
    return res


def estado_caja(fecha=None):
    """Arqueo de caja del día, con las secciones del formato de recepción.

    1. Inicio de día (soles + dólares)
    2. Efectivo contado al cierre (desglose por denominación)
    3. Ventas totales por forma de pago (desde VENTA DIARIA)
    4. Gastos de caja chica / ingresos extra
    5. Ajuste: sobrante o faltante
    """
    fecha = fecha or _fecha_hoy_str()
    filas = leer_caja(fecha)
    apertura = next((f for f in filas if f['tipo'] == 'APERTURA'), None)
    cierre = next((f for f in filas if f['tipo'] == 'CIERRE'), None)
    movimientos = [f for f in filas if f['tipo'] in ('INGRESO', 'EGRESO')]
    ingresos = round(sum(f['monto'] for f in movimientos if f['tipo'] == 'INGRESO'), 2)
    egresos = round(sum(f['monto'] for f in movimientos if f['tipo'] == 'EGRESO'), 2)
    egresos_usd = round(sum(f['monto_usd'] for f in movimientos if f['tipo'] == 'EGRESO'), 2)
    ingresos_usd = round(sum(f['monto_usd'] for f in movimientos if f['tipo'] == 'INGRESO'), 2)

    ventas = ventas_del_dia(fecha)
    inicio = apertura['monto'] if apertura else 0
    inicio_usd = apertura['monto_usd'] if apertura else 0

    esperado = round(inicio + ventas['efectivo'] + ingresos - egresos, 2)
    esperado_usd = round(inicio_usd + ingresos_usd - egresos_usd, 2)
    contado = cierre['monto'] if cierre else None
    contado_usd = cierre['monto_usd'] if cierre else None
    diferencia = round(contado - esperado, 2) if cierre else None
    diferencia_usd = round(contado_usd - esperado_usd, 2) if cierre else None

    return {
        'fecha': fecha,
        'abierta': apertura is not None,
        'cerrada': cierre is not None,
        'apertura': apertura,
        'cierre': cierre,
        'movimientos': movimientos,
        'ventas': ventas,
        'ingresos': ingresos,
        'egresos': egresos,
        'ingresos_usd': ingresos_usd,
        'egresos_usd': egresos_usd,
        'inicio': inicio,
        'inicio_usd': inicio_usd,
        'esperado': esperado,
        'esperado_usd': esperado_usd,
        'contado': contado,
        'contado_usd': contado_usd,
        'diferencia': diferencia,
        'diferencia_usd': diferencia_usd,
        # alias de compatibilidad
        'monto_inicial': inicio,
        'ventas_efectivo': ventas['efectivo'],
    }


def historial_caja(limite=30):
    filas = leer_caja()
    fechas = sorted({f['fecha'] for f in filas if f['tipo'] == 'APERTURA'},
                    key=_fecha_orden, reverse=True)[:limite]
    return [estado_caja(f) for f in fechas]


# ============================================================
# CRM Plus: Inventario (catálogo de productos + kardex de stock)
# ============================================================
def _producto_named(f):
    p = {
        'id': am.num(f.get('A')),
        'codigo': f.get('B'),
        'producto': f.get('C'),
        'costo_bruto': am.num(f.get('D')) or 0,
        'stock': am.num(f.get('E')) or 0,
        'stock_minimo': am.num(f.get('F')) or 0,
        'unidad': f.get('G') or 'UND',
        'fecha_alta': f.get('H'),
        'actualizado': f.get('I'),
    }
    p['stock_bajo'] = p['stock'] <= p['stock_minimo']
    return p


def leer_productos():
    out = [_producto_named(f) for f in _leer_hoja('PRODUCTOS')]
    out.sort(key=lambda p: str(p['producto'] or '').upper())
    return out


def producto_por_codigo(codigo):
    codigo = str(codigo or '').strip().upper()
    if not codigo:
        return None
    for f in _leer_hoja('PRODUCTOS'):
        if str(f.get('B') or '').strip().upper() == codigo:
            return _producto_named(f)
    return None


def crear_producto(datos):
    with _lock:
        filas = _leer_hoja('PRODUCTOS')
        codigo = str(datos.get('codigo') or '').strip().upper()
        if not codigo:
            raise ValueError('Indica el código del producto')
        if any(str(f.get('B') or '').strip().upper() == codigo for f in filas):
            raise ValueError(f'Ya existe un producto con el código "{codigo}"')
        producto = str(datos.get('producto') or '').strip()
        if not producto:
            raise ValueError('Indica el nombre del producto')
        f = {
            'A': _siguiente_id(filas), 'B': codigo, 'C': producto,
            'D': am.num(datos.get('costo_bruto')) or 0,
            'E': am.num(datos.get('stock')) or 0,
            'F': am.num(datos.get('stock_minimo')) or 0,
            'G': str(datos.get('unidad') or 'UND').strip().upper(),
            'H': _ahora(), 'I': _ahora(),
        }
        filas.append(f)
        _reescribir('PRODUCTOS', filas)
        return _producto_named(f)


def actualizar_producto(pid, cambios):
    with _lock:
        filas = _leer_hoja('PRODUCTOS')
        for f in filas:
            if am.num(f.get('A')) == pid:
                if cambios.get('codigo'):
                    nuevo = str(cambios['codigo']).strip().upper()
                    if any(am.num(g.get('A')) != pid
                           and str(g.get('B') or '').strip().upper() == nuevo
                           for g in filas):
                        raise ValueError(f'Ya existe un producto con el código "{nuevo}"')
                    f['B'] = nuevo
                for campo, letra in (('producto', 'C'), ('unidad', 'G')):
                    if cambios.get(campo):
                        f[letra] = str(cambios[campo]).strip()
                for campo, letra in (('costo_bruto', 'D'), ('stock_minimo', 'F')):
                    if campo in cambios and cambios[campo] not in (None, ''):
                        f[letra] = am.num(cambios[campo]) or 0
                f['I'] = _ahora()
                _reescribir('PRODUCTOS', filas)
                return _producto_named(f)
    return None


def borrar_producto(pid):
    with _lock:
        filas = _leer_hoja('PRODUCTOS')
        nuevas = [f for f in filas if am.num(f.get('A')) != pid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('PRODUCTOS', nuevas)
        return True


def _movimiento_named(f):
    return {
        'id': am.num(f.get('A')), 'fecha': f.get('B'), 'codigo': f.get('C'),
        'producto': f.get('D'), 'tipo': f.get('E'), 'cantidad': am.num(f.get('F')) or 0,
        'stock_resultante': am.num(f.get('G')) or 0, 'referencia': f.get('H'),
        'nota': f.get('I'),
    }


def leer_movimientos_stock(codigo=None, limite=200):
    filas = [_movimiento_named(f) for f in _leer_hoja('MOVIMIENTOS_STOCK')]
    if codigo:
        c = str(codigo).strip().upper()
        filas = [f for f in filas if str(f['codigo'] or '').strip().upper() == c]
    filas.sort(key=lambda f: am.num(f['id']) or 0, reverse=True)
    return filas[:limite]


def registrar_movimiento_stock(codigo, tipo, cantidad, referencia='', nota=''):
    """Registra un movimiento de stock y actualiza el producto.

    ENTRADA suma, SALIDA resta, AJUSTE fija el stock al valor de
    ``cantidad`` directamente (para corregir conteos).
    """
    tipo = str(tipo or '').upper().strip()
    if tipo not in TIPO_MOVIMIENTO:
        raise ValueError(f"El tipo de movimiento debe ser {', '.join(TIPO_MOVIMIENTO)}")
    cantidad = am.num(cantidad)
    if cantidad is None or cantidad < 0:
        raise ValueError('Indica una cantidad válida')
    with _lock:
        todo = {n: _leer_hoja(n) for n in HOJAS}
        codigo_u = str(codigo or '').strip().upper()
        prod_f = next((f for f in todo['PRODUCTOS']
                       if str(f.get('B') or '').strip().upper() == codigo_u), None)
        if prod_f is None:
            raise ValueError(f'No existe un producto con el código "{codigo}"')
        stock_actual = am.num(prod_f.get('E')) or 0
        if tipo == 'ENTRADA':
            nuevo_stock = stock_actual + cantidad
        elif tipo == 'SALIDA':
            nuevo_stock = stock_actual - cantidad
        else:
            nuevo_stock = cantidad
        prod_f['E'] = nuevo_stock
        prod_f['I'] = _ahora()

        m = {'A': _siguiente_id(todo['MOVIMIENTOS_STOCK']), 'B': _ahora(),
             'C': prod_f.get('B'), 'D': prod_f.get('C'), 'E': tipo, 'F': cantidad,
             'G': nuevo_stock, 'H': str(referencia or '').strip(),
             'I': str(nota or '').strip()}
        todo['MOVIMIENTOS_STOCK'].append(m)
        _guardar(todo)
        return {'producto': _producto_named(prod_f), 'movimiento': _movimiento_named(m)}


def registrar_entrada_stock(codigo, cantidad, referencia='', nota=''):
    return registrar_movimiento_stock(codigo, 'ENTRADA', cantidad, referencia, nota)


def registrar_salida_stock(codigo, cantidad, referencia='', nota=''):
    return registrar_movimiento_stock(codigo, 'SALIDA', cantidad, referencia, nota)


def ajustar_stock(codigo, stock_nuevo, nota=''):
    return registrar_movimiento_stock(codigo, 'AJUSTE', stock_nuevo, 'Ajuste manual', nota)


# ============================================================
# CRM Plus: Planilla (trabajadores + pagos quincenales)
# ============================================================
def _trabajador_named(f):
    return {
        'id': am.num(f.get('A')),
        'nombre': f.get('B'),
        'cargo': f.get('C'),
        'sueldo_quincena': am.num(f.get('D')) or 0,
        'porcentaje_comision': am.num(f.get('E')) or 0,
        'metodo_pago': f.get('F'),
        'cuenta': f.get('G'),
        'estado': f.get('H') or 'ACTIVO',
        'fecha_alta': f.get('I'),
    }


def leer_trabajadores(solo_activos=False):
    out = [_trabajador_named(f) for f in _leer_hoja('TRABAJADORES')]
    if solo_activos:
        out = [t for t in out if t['estado'] == 'ACTIVO']
    out.sort(key=lambda t: str(t['nombre'] or '').upper())
    return out


def crear_trabajador(datos):
    with _lock:
        filas = _leer_hoja('TRABAJADORES')
        nombre = str(datos.get('nombre') or '').strip()
        if not nombre:
            raise ValueError('Indica el nombre del trabajador')
        f = {
            'A': _siguiente_id(filas), 'B': nombre,
            'C': str(datos.get('cargo') or '').strip(),
            'D': am.num(datos.get('sueldo_quincena')) or 0,
            'E': am.num(datos.get('porcentaje_comision')) or 0,
            'F': str(datos.get('metodo_pago') or '').strip(),
            'G': str(datos.get('cuenta') or '').strip(),
            'H': 'ACTIVO', 'I': _ahora(),
        }
        filas.append(f)
        _reescribir('TRABAJADORES', filas)
        return _trabajador_named(f)


def actualizar_trabajador(tid, cambios):
    with _lock:
        filas = _leer_hoja('TRABAJADORES')
        for f in filas:
            if am.num(f.get('A')) == tid:
                if cambios.get('nombre'):
                    f['B'] = str(cambios['nombre']).strip()
                if 'cargo' in cambios:
                    f['C'] = str(cambios.get('cargo') or '').strip()
                if 'sueldo_quincena' in cambios and cambios['sueldo_quincena'] not in (None, ''):
                    f['D'] = am.num(cambios['sueldo_quincena']) or 0
                if 'porcentaje_comision' in cambios and cambios['porcentaje_comision'] not in (None, ''):
                    f['E'] = am.num(cambios['porcentaje_comision']) or 0
                if 'metodo_pago' in cambios:
                    f['F'] = str(cambios.get('metodo_pago') or '').strip()
                if 'cuenta' in cambios:
                    f['G'] = str(cambios.get('cuenta') or '').strip()
                if cambios.get('estado') in ESTADO_TRABAJADOR:
                    f['H'] = cambios['estado']
                _reescribir('TRABAJADORES', filas)
                return _trabajador_named(f)
    return None


def borrar_trabajador(tid):
    with _lock:
        filas = _leer_hoja('TRABAJADORES')
        nuevas = [f for f in filas if am.num(f.get('A')) != tid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('TRABAJADORES', nuevas)
        return True


def _rango_quincena(anio, mes_num, quincena):
    ultimo = calendar.monthrange(anio, mes_num)[1]
    return (1, 15) if quincena == 1 else (16, ultimo)


def calcular_comision(nombre, anio, mes, quincena, porcentaje):
    """Comisión = % del monto de VENTA DIARIA (columna VENTA) de las ventas
    realizadas (excluye NO SE REALIZO) atribuidas a ``nombre`` por la columna
    DOCTOR, dentro de la quincena dada. Requiere que el nombre del trabajador
    coincida (sin distinguir mayúsculas) con lo escrito en esa columna."""
    porcentaje = am.num(porcentaje) or 0
    if not porcentaje:
        return 0.0
    mes = str(mes or '').strip().upper()
    mes_num = _MM_IDX.get(mes)
    nombre_u = str(nombre or '').strip().upper()
    if not mes_num or not nombre_u:
        return 0.0
    d1, d2 = _rango_quincena(int(anio), mes_num, int(quincena))
    total = 0.0
    for v in cd.leer_venta()['hojas'].values():
        for f in v['filas']:
            if str(f.get('M') or '').strip().upper() != nombre_u:
                continue
            if str(f.get('C') or '').strip().upper() != mes:
                continue
            if am.num(f.get('D')) != int(anio):
                continue
            dia = am.num(f.get('B'))
            if dia is None or not (d1 <= dia <= d2):
                continue
            if not _es_venta_registrada(f.get('N')):
                continue
            total += am.num(f.get('O')) or 0
    return round(total * porcentaje / 100.0, 2)


def _pago_named(f):
    return {
        'id': am.num(f.get('A')),
        'trabajador_id': am.num(f.get('B')),
        'nombre': f.get('C'),
        'anio': am.num(f.get('D')),
        'mes': f.get('E'),
        'quincena': am.num(f.get('F')),
        'sueldo_base': am.num(f.get('G')) or 0,
        'comision': am.num(f.get('H')) or 0,
        'monto_total': am.num(f.get('I')) or 0,
        'estado': f.get('J') or 'PENDIENTE',
        'fecha_pago': f.get('K'),
        'metodo_pago': f.get('L'),
        'nota': f.get('M'),
    }


def leer_planilla(anio=None, mes=None, quincena=None):
    out = [_pago_named(f) for f in _leer_hoja('PLANILLA')]
    if anio:
        out = [p for p in out if p['anio'] == int(anio)]
    if mes:
        out = [p for p in out if str(p['mes'] or '').strip().upper() == str(mes).strip().upper()]
    if quincena:
        out = [p for p in out if p['quincena'] == int(quincena)]
    out.sort(key=lambda p: str(p['nombre'] or '').upper())
    return out


def generar_planilla_quincena(anio, mes, quincena):
    """Crea (o refresca) el pago de la quincena para cada trabajador ACTIVO.
    No toca los pagos que ya están marcados como PAGADO."""
    anio = int(anio)
    mes = str(mes).strip().upper()
    quincena = int(quincena)
    if quincena not in (1, 2):
        raise ValueError('La quincena debe ser 1 o 2')
    if mes not in _MM_IDX:
        raise ValueError(f'Mes inválido: {mes}')
    with _lock:
        todo = {n: _leer_hoja(n) for n in HOJAS}
        trabajadores = [_trabajador_named(f) for f in todo['TRABAJADORES']
                        if (f.get('H') or 'ACTIVO') == 'ACTIVO']
        if not trabajadores:
            raise ValueError('No hay trabajadores activos en el catálogo')
        existentes = {}
        for f in todo['PLANILLA']:
            clave = (am.num(f.get('B')), am.num(f.get('D')),
                     str(f.get('E') or '').strip().upper(), am.num(f.get('F')))
            existentes[clave] = f
        creados = actualizados = sin_tocar = 0
        for t in trabajadores:
            comision = calcular_comision(t['nombre'], anio, mes, quincena, t['porcentaje_comision'])
            sueldo = t['sueldo_quincena']
            total = round(sueldo + comision, 2)
            clave = (t['id'], anio, mes, quincena)
            f = existentes.get(clave)
            if f is not None:
                if (f.get('J') or 'PENDIENTE') == 'PAGADO':
                    sin_tocar += 1
                    continue
                f['G'] = sueldo; f['H'] = comision; f['I'] = total
                actualizados += 1
            else:
                f = {'A': _siguiente_id(todo['PLANILLA']), 'B': t['id'], 'C': t['nombre'],
                     'D': anio, 'E': mes, 'F': quincena, 'G': sueldo, 'H': comision,
                     'I': total, 'J': 'PENDIENTE'}
                todo['PLANILLA'].append(f)
                creados += 1
        _guardar(todo)
        return {'creados': creados, 'actualizados': actualizados, 'sin_tocar': sin_tocar}


def actualizar_pago_planilla(pid, cambios):
    with _lock:
        filas = _leer_hoja('PLANILLA')
        for f in filas:
            if am.num(f.get('A')) == pid:
                if 'sueldo_base' in cambios and cambios['sueldo_base'] not in (None, ''):
                    f['G'] = am.num(cambios['sueldo_base']) or 0
                if 'comision' in cambios and cambios['comision'] not in (None, ''):
                    f['H'] = am.num(cambios['comision']) or 0
                if 'sueldo_base' in cambios or 'comision' in cambios:
                    f['I'] = round((am.num(f.get('G')) or 0) + (am.num(f.get('H')) or 0), 2)
                if cambios.get('estado') in ESTADO_PAGO:
                    f['J'] = cambios['estado']
                    if cambios['estado'] == 'PAGADO':
                        f['K'] = f.get('K') or _fecha_hoy_str()
                    else:
                        f['K'] = None
                if 'metodo_pago' in cambios:
                    f['L'] = str(cambios.get('metodo_pago') or '').strip()
                if 'nota' in cambios:
                    f['M'] = str(cambios.get('nota') or '').strip()
                _reescribir('PLANILLA', filas)
                return _pago_named(f)
    return None


def borrar_pago_planilla(pid):
    with _lock:
        filas = _leer_hoja('PLANILLA')
        nuevas = [f for f in filas if am.num(f.get('A')) != pid]
        if len(nuevas) == len(filas):
            return False
        _reescribir('PLANILLA', nuevas)
        return True
