# -*- coding: utf-8 -*-
"""crm_plus.py
============
Funciones estilo CRM (tipo KOMMO) sobre Google Drive: pipeline kanban,
tareas/recordatorios y notas de seguimiento por paciente.

Guarda todo en un workbook propio ``CRM.xlsx`` (hojas TARJETAS, TAREAS y
NOTAS) que se busca por nombre en Drive y se crea automáticamente la primera
vez. Además agrega vistas derivadas: directorio de pacientes, dashboard de
métricas y panel de actividades de hoy (combinando AGENDADOS y VENTA DIARIA).
"""
import io
import os
import re
import threading
import time
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

HOJAS = {
    'TARJETAS': ['ID', 'NOMBRE', 'TELEFONO', 'ETAPA', 'CRM', 'CAMPANA', 'VALOR',
                 'PRIORIDAD', 'FECHA_ALTA', 'FECHA_ULT', 'CITA_DIA', 'CITA_MES',
                 'CITA_HORA', 'NOTA'],
    'TAREAS': ['ID', 'TITULO', 'TIPO', 'FECHA', 'HORA', 'ESTADO', 'PRIORIDAD',
               'CONTACTO', 'NOTA'],
    'NOTAS': ['ID', 'TELEFONO', 'CONTACTO', 'FECHA', 'TIPO', 'TEXTO'],
}

TARJETA_COLS = {'id': 'A', 'nombre': 'B', 'telefono': 'C', 'etapa': 'D',
                'crm': 'E', 'campana': 'F', 'valor': 'G', 'prioridad': 'H',
                'fecha_alta': 'I', 'fecha_ult': 'J', 'cita_dia': 'K',
                'cita_mes': 'L', 'cita_hora': 'M', 'nota': 'N'}
TAREA_COLS = {'id': 'A', 'titulo': 'B', 'tipo': 'C', 'fecha': 'D', 'hora': 'E',
              'estado': 'F', 'prioridad': 'G', 'contacto': 'H', 'nota': 'I'}
NOTA_COLS = {'id': 'A', 'telefono': 'B', 'contacto': 'C', 'fecha': 'D',
             'tipo': 'E', 'texto': 'F'}

_lock = threading.RLock()


def _ahora():
    return datetime.now(TZ).strftime('%d/%m/%Y %H:%M')


def _hoy():
    hoy = datetime.now(TZ)
    return hoy.day, _MM[hoy.month], hoy.year


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
                      'dia': f.get('L'), 'mes': f.get('M')})
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
