# -*- coding: utf-8 -*-
"""Web app de reportes Derma Essenza.

Genera los reportes PDF de ventas por campaña/CRM y sincroniza el maestro
BD DATA.xlsx desde Google Drive (AGENDADOS + VENTA DIARIA), igual que los
scripts de terminal reporte_ventas_pdf.py y alimentar_maestro.py.

Variables de entorno (además de las de alimentar_maestro.py):
  DATA_DIR   carpeta persistente para reportes y backups (default: ./data)
  REPORTES_DIR  carpeta para los PDF (default: DATA_DIR/reportes)
"""
import io
import os
import shutil
import sys
import threading
from datetime import datetime
from typing import Optional

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import alimentar_maestro as am          # noqa: E402
import crm_drive as crm                 # noqa: E402
import crm_plus as cp                   # noqa: E402
import reporte_ventas_pdf as rv         # noqa: E402

DATA_DIR = os.environ.get('DATA_DIR', os.path.join(BASE_DIR, 'data'))
REPORTES_DIR = os.environ.get('REPORTES_DIR', os.path.join(DATA_DIR, 'reportes'))
os.makedirs(REPORTES_DIR, exist_ok=True)

_bloqueo = threading.Lock()

app = FastAPI(title='Reportes Beauty Medic', version='1.0.0')
app.add_middleware(CORSMiddleware, allow_origins=['*'], allow_methods=['*'],
                   allow_headers=['*'])

BRAND = {
    'BRAND_NOMBRE': os.environ.get('BRAND_NOMBRE', 'Beauty Medic'),
    'BRAND_BADGE': os.environ.get('BRAND_BADGE', 'BM'),
}


def _html_index():
    with open(os.path.join(BASE_DIR, 'templates', 'index.html'),
              encoding='utf-8') as f:
        html = f.read()
    for k, v in BRAND.items():
        html = html.replace('{{' + k + '}}', v)
    return html


def _nombre_reporte(mes, anio, desde, hasta):
    nombre = f'Reporte_Ventas_{desde}-{hasta}_{rv.NOMBRES_MES.get(mes, mes)}_{anio}'
    for f in os.listdir(REPORTES_DIR):
        if f.lower().endswith('.pdf'):
            try:
                os.remove(os.path.join(REPORTES_DIR, f))
            except OSError:
                pass
    return os.path.join(REPORTES_DIR, f'{nombre}.pdf')


def _listar_reportes():
    out = []
    for f in sorted(os.listdir(REPORTES_DIR), reverse=True):
        if f.lower().endswith('.pdf'):
            p = os.path.join(REPORTES_DIR, f)
            out.append({'archivo': f, 'modificado': datetime.fromtimestamp(
                os.path.getmtime(p)).strftime('%d/%m/%Y %H:%M'),
                'tamano_kb': round(os.path.getsize(p) / 1024, 1)})
    return out


# ============================================================
# Schemas
# ============================================================
class ReporteReq(BaseModel):
    mes: str = 'AGO'
    anio: int = 2026
    desde: int = 1
    hasta: int = 10
    fuente: str = 'maestro'


class SyncReq(BaseModel):
    aplicar: bool = False


class AgendadoReq(BaseModel):
    crm: str = ''
    dia: Optional[int] = None
    mes: str = ''
    anio: Optional[int] = None
    nombre: str = ''
    red_social: str = ''
    telefono: str = ''
    correo: str = ''
    agendado_por: str = ''
    dia_cita: Optional[int] = None
    mes_cita: str = ''
    anio_cita: Optional[int] = None
    campana: str = ''
    hora: str = ''
    confirmado: str = ''
    observacion: str = ''
    reconfirmado: str = ''
    observacion2: str = ''


class VentaReq(BaseModel):
    dia: Optional[int] = None
    mes: str = ''
    anio: Optional[int] = None
    dni: str = ''
    cel: str = ''
    nombre: str = ''
    nuevo: str = ''
    distrito: str = ''
    edad: Optional[int] = None
    sexo: str = ''
    tratamiento: str = ''
    doctor: str = ''
    status: str = ''
    venta: Optional[float] = None
    pago: str = ''
    comisiona: Optional[float] = None
    observacion: str = ''


class TarjetaReq(BaseModel):
    nombre: str = ''
    telefono: str = ''
    etapa: str = ''
    crm: str = ''
    campana: str = ''
    valor: Optional[float] = None
    prioridad: str = ''
    cita_dia: Optional[int] = None
    cita_mes: str = ''
    cita_hora: str = ''
    nota: str = ''


class TareaReq(BaseModel):
    titulo: str = ''
    tipo: str = ''
    fecha: str = ''
    hora: str = ''
    estado: str = ''
    prioridad: str = ''
    contacto: str = ''
    nota: str = ''


class NotaReq(BaseModel):
    telefono: str = ''
    contacto: str = ''
    tipo: str = ''
    texto: str = ''


MESES_VALIDOS = {'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO',
                 'SET', 'SEP', 'OCT', 'NOV', 'DIC'}


def _validar_fecha(dia, mes, anio, requerido_mes=False):
    if dia is not None and not (1 <= dia <= 31):
        raise HTTPException(400, f'Día inválido: {dia}')
    if mes:
        m = mes.upper().replace('SEP', 'SET')
        if m not in MESES_VALIDOS:
            raise HTTPException(400, f'Mes inválido: {mes}')
    if anio is not None and not (2020 <= anio <= 2100):
        raise HTTPException(400, f'Año inválido: {anio}')


# ============================================================
# Página y estado
# ============================================================
@app.get('/', response_class=HTMLResponse)
async def index():
    return HTMLResponse(_html_index())


@app.get('/api/estado')
async def estado():
    m_ok = os.path.exists(am.MAESTRO)
    return {
        'credenciales_ok': am.credenciales_disponibles(),
        'credenciales': am.CREDENCIALES,
        'maestro_ok': m_ok,
        'maestro': am.MAESTRO,
        'maestro_modificado': (datetime.fromtimestamp(os.path.getmtime(am.MAESTRO))
                               .strftime('%d/%m/%Y %H:%M') if m_ok else None),
        'agendados_ok': os.path.exists(os.path.join(am.TMP_DIR, 'AGENDADOS.xlsx')),
        'venta_ok': os.path.exists(os.path.join(am.TMP_DIR, 'VENTA_DIARIA.xlsx')),
        'meses': list(rv.NOMBRES_MES.keys()),
        'reportes': _listar_reportes(),
    }


# ============================================================
# Generar reporte PDF
# ============================================================
@app.post('/api/reporte')
async def generar_reporte(data: ReporteReq):
    mes = (data.mes or 'AGO').upper().replace('SEP', 'SET')
    if mes not in rv.NOMBRES_MES:
        raise HTTPException(400, f'Mes inválido: {mes}')
    if not (1 <= data.desde <= 31 and 1 <= data.hasta <= 31 and data.desde <= data.hasta):
        raise HTTPException(400, 'Rango de días inválido')
    if data.fuente not in ('maestro', 'auto'):
        raise HTTPException(400, 'Fuente inválida')
    if not os.path.exists(am.MAESTRO):
        raise HTTPException(400, 'No hay maestro BD DATA.xlsx. Súbelo primero.')
    with _bloqueo:
        if data.fuente == 'auto':
            am.descargar(am.AGENDADOS_FID, 'AGENDADOS')
            am.descargar(am.VENTA_FID, 'VENTA_DIARIA')
        else:
            am.ejecutar_sync(aplicar=True)
        ruta = _nombre_reporte(mes, data.anio, data.desde, data.hasta)
        try:
            res = rv.generar_reporte(mes=mes, anio=data.anio, desde=data.desde,
                                     hasta=data.hasta, fuente=data.fuente,
                                     salida=ruta)
        except Exception as e:  # noqa: BLE001
            raise HTTPException(500, f'Error generando el reporte: {e}')
    verificacion = None
    if os.path.exists(os.path.join(am.TMP_DIR, 'VENTA_DIARIA.xlsx')):
        try:
            verificacion = am.verificar_venta_vs_td(
                os.path.join(am.TMP_DIR, 'VENTA_DIARIA.xlsx'),
                data.anio, mes, data.desde, data.hasta)
            if verificacion.get('ok'):
                reporte_monto = res['totales']['mon']
                monto_coincide = abs(reporte_monto - verificacion['venta']) < 0.01
                verificacion['reporte_monto'] = reporte_monto
                verificacion['monto_coincide'] = monto_coincide
                if not monto_coincide:
                    verificacion['mensaje'] = (
                        f"El reporte maestro da S/ {reporte_monto:,.2f} pero VENTA 2026 / "
                        f"TD 2026 suman S/ {verificacion['venta']:,.2f} en el periodo. "
                        'Algún dato está mal (revisa PAGO TOTAL acumulado en el maestro).')
                    verificacion['coincide'] = False
        except Exception:  # noqa: BLE001
            verificacion = None
    return {'ok': True, 'archivo': os.path.basename(res['archivo']),
            'url': f'/api/reporte/download/{os.path.basename(res["archivo"])}',
            'totales': res['totales'], 'por_crm': res['por_crm'],
            'detalle': res['detalle'], 'verificacion': verificacion}


@app.get('/api/reportes')
async def listar_reportes():
    return _listar_reportes()


@app.get('/api/reporte/download/{archivo}')
async def descargar_reporte(archivo: str):
    p = os.path.join(REPORTES_DIR, os.path.basename(archivo))
    if not os.path.isfile(p):
        raise HTTPException(404, 'Reporte no encontrado')
    return FileResponse(p, filename=os.path.basename(p))


@app.delete('/api/reporte/{archivo}')
async def borrar_reporte(archivo: str):
    p = os.path.join(REPORTES_DIR, os.path.basename(archivo))
    if not os.path.isfile(p):
        raise HTTPException(404, 'Reporte no encontrado')
    os.remove(p)
    return {'ok': True}


# ============================================================
# Sincronizar maestro desde Drive
# ============================================================
@app.post('/api/sync')
async def sincronizar(data: SyncReq):
    with _bloqueo:
        resultado = am.ejecutar_sync(aplicar=data.aplicar)
    return resultado


# ============================================================
# Maestro BD DATA.xlsx (subir / descargar)
# ============================================================
@app.post('/api/maestro/upload')
async def subir_maestro(file: UploadFile = File(...)):
    contenido = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True)
        if 'BD DATA' not in wb.sheetnames:
            raise ValueError('El archivo no contiene la hoja "BD DATA"')
        wb.close()
    except Exception as e:  # noqa: BLE001
        raise HTTPException(400, f'Archivo inválido: {e}')
    os.makedirs(os.path.dirname(am.MAESTRO) or '.', exist_ok=True)
    if os.path.exists(am.MAESTRO):
        backup = am.MAESTRO.replace('.xlsx', f'_pre_upload_{datetime.now():%Y%m%d_%H%M%S}.xlsx')
        shutil.copy2(am.MAESTRO, backup)
    with open(am.MAESTRO, 'wb') as f:
        f.write(contenido)
    return {'ok': True, 'backup': backup if 'backup' in dir() else None}


@app.get('/api/maestro/download')
async def descargar_maestro():
    if not os.path.exists(am.MAESTRO):
        raise HTTPException(404, 'No hay maestro subido')
    return FileResponse(am.MAESTRO, filename='BD DATA.xlsx')


# ============================================================
# CRM: Agendados y Venta diaria (lectura / alta en Drive)
# ============================================================
def _col_order(filas):
    cols = set()
    for f in filas:
        cols.update(f.keys())
    return sorted(cols, key=lambda c: openpyxl.utils.column_index_from_string(c))


@app.get('/api/crm/agendados')
async def crm_agendados():
    try:
        data = crm.leer_agendados()
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo leer AGENDADOS de Drive: {e}')
    return {
        'ok': True,
        'filas': data['filas'],
        'total': data['total'],
        'descargado': data['descargado'],
        'columnas': data['columnas'],
        'valores': crm.valores_unicos_agendados(data['filas']),
    }


@app.post('/api/crm/agendados')
async def crm_agendados_nuevo(data: AgendadoReq):
    d = data.model_dump()
    _validar_fecha(d.get('dia'), d.get('mes'), d.get('anio'))
    _validar_fecha(d.get('dia_cita'), d.get('mes_cita'), d.get('anio_cita'),
                   requerido_mes=True)
    if not d.get('nombre') and not d.get('telefono'):
        raise HTTPException(400, 'Indica al menos el nombre o el teléfono del paciente')
    with _bloqueo:
        try:
            res = crm.agregar_agendado(d)
        except HTTPException:
            raise
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f'No se pudo guardar en AGENDADOS (Drive): {e}')
    return res


@app.get('/api/crm/venta')
async def crm_venta():
    try:
        data = crm.leer_venta()
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo leer VENTA DIARIA de Drive: {e}')
    hojas = {h: {'filas': v['filas'], 'total': len(v['filas']),
                 'columnas': v['columnas']}
             for h, v in data['hojas'].items()}
    todas = [f for v in data['hojas'].values() for f in v['filas']]
    return {
        'ok': True,
        'hojas': hojas,
        'descargado': data['descargado'],
        'valores': crm.valores_unicos_venta(todas),
    }


@app.post('/api/crm/venta')
async def crm_venta_nuevo(data: VentaReq):
    d = data.model_dump()
    _validar_fecha(d.get('dia'), d.get('mes'), d.get('anio'))
    if not d.get('nombre'):
        raise HTTPException(400, 'Indica el nombre del paciente')
    with _bloqueo:
        try:
            res = crm.agregar_venta(d)
        except HTTPException:
            raise
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f'No se pudo guardar en VENTA DIARIA (Drive): {e}')
    return res


# ============================================================
# CRM Plus: Pipeline kanban
# ============================================================
@app.get('/api/crm/pipeline')
async def crm_pipeline():
    try:
        return {'ok': True, 'etapas': cp.ETAPAS, 'tarjetas': cp.leer_tarjetas()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo leer el pipeline (Drive): {e}')


@app.post('/api/crm/pipeline')
async def crm_pipeline_nueva(data: TarjetaReq):
    try:
        res = cp.crear_tarjeta(data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo guardar la tarjeta (Drive): {e}')
    return {'ok': True, 'tarjeta': res}


@app.patch('/api/crm/pipeline/{tid}')
async def crm_pipeline_actualizar(tid: int, data: TarjetaReq):
    try:
        res = cp.actualizar_tarjeta(tid, data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo actualizar la tarjeta (Drive): {e}')
    if not res:
        raise HTTPException(404, 'Tarjeta no encontrada')
    return {'ok': True, 'tarjeta': res}


@app.delete('/api/crm/pipeline/{tid}')
async def crm_pipeline_borrar(tid: int):
    try:
        ok = cp.borrar_tarjeta(tid)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo borrar la tarjeta (Drive): {e}')
    if not ok:
        raise HTTPException(404, 'Tarjeta no encontrada')
    return {'ok': True}


# ============================================================
# CRM Plus: Tareas y recordatorios
# ============================================================
@app.get('/api/crm/tareas')
async def crm_tareas(estado: str = ''):
    try:
        return {'ok': True, 'tareas': cp.leer_tareas(estado=estado or None)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudieron leer las tareas (Drive): {e}')


@app.post('/api/crm/tareas')
async def crm_tareas_nueva(data: TareaReq):
    if not data.titulo:
        raise HTTPException(400, 'Indica el título de la tarea')
    try:
        res = cp.crear_tarea(data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo guardar la tarea (Drive): {e}')
    return {'ok': True, 'tarea': res}


@app.patch('/api/crm/tareas/{tid}')
async def crm_tareas_actualizar(tid: int, data: TareaReq):
    try:
        res = cp.actualizar_tarea(tid, data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo actualizar la tarea (Drive): {e}')
    if not res:
        raise HTTPException(404, 'Tarea no encontrada')
    return {'ok': True, 'tarea': res}


@app.delete('/api/crm/tareas/{tid}')
async def crm_tareas_borrar(tid: int):
    try:
        ok = cp.borrar_tarea(tid)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo borrar la tarea (Drive): {e}')
    if not ok:
        raise HTTPException(404, 'Tarea no encontrada')
    return {'ok': True}


# ============================================================
# CRM Plus: Notas de seguimiento
# ============================================================
@app.get('/api/crm/notas')
async def crm_notas(telefono: str = '', contacto: str = ''):
    try:
        return {'ok': True, 'notas': cp.leer_notas(
            telefono=telefono or None, contacto=contacto or None)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudieron leer las notas (Drive): {e}')


@app.post('/api/crm/notas')
async def crm_notas_nueva(data: NotaReq):
    if not data.texto:
        raise HTTPException(400, 'Escribe la nota')
    try:
        res = cp.agregar_nota(data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo guardar la nota (Drive): {e}')
    return {'ok': True, 'nota': res}


# ============================================================
# CRM Plus: Directorio, dashboard y actividades de hoy
# ============================================================
@app.get('/api/crm/pacientes')
async def crm_pacientes():
    try:
        return {'ok': True, 'pacientes': cp.leer_pacientes()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo armar el directorio (Drive): {e}')


@app.get('/api/crm/dashboard')
async def crm_dashboard():
    try:
        return {'ok': True, **cp.leer_dashboard()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular el dashboard (Drive): {e}')


@app.get('/api/crm/hoy')
async def crm_hoy():
    try:
        return {'ok': True, **cp.leer_hoy()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo armar el panel de hoy (Drive): {e}')
