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
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import alimentar_maestro as am          # noqa: E402
import analitica as ana                 # noqa: E402
import crm_drive as crm                 # noqa: E402
import crm_plus as cp                   # noqa: E402
import reporte_ventas_pdf as rv         # noqa: E402

DATA_DIR = os.environ.get('DATA_DIR', os.path.join(BASE_DIR, 'data'))
REPORTES_DIR = os.environ.get('REPORTES_DIR', os.path.join(DATA_DIR, 'reportes'))
BACKUP_DIR = os.environ.get('BACKUP_DIR', os.path.join(DATA_DIR, 'backups'))
MAX_BACKUPS = int(os.environ.get('MAX_BACKUPS', '12'))
os.makedirs(REPORTES_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)


def _crear_backup():
    """Copia maestro y CRM.xlsx a la carpeta de backups, conservando los últimos N."""
    hecha = []
    hoy = datetime.now().strftime('%Y%m%d_%H%M%S')
    for origen, nombre in ((am.ruta_maestro_local(), 'BD DATA.xlsx'),
                           (os.path.join(am.TMP_DIR, 'CRM.xlsx'), 'CRM.xlsx')):
        if not os.path.exists(origen):
            continue
        destino = os.path.join(BACKUP_DIR, f'{hoy}_{nombre}')
        shutil.copy2(origen, destino)
        hecha.append(destino)
    for f in sorted(os.listdir(BACKUP_DIR)):
        if not f.lower().endswith('.xlsx'):
            continue
        versiones = sorted(g for g in os.listdir(BACKUP_DIR)
                           if g.endswith(f[f.find('_'):]))
        while len(versiones) > MAX_BACKUPS:
            os.remove(os.path.join(BACKUP_DIR, versiones.pop(0)))
    return hecha

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


app.mount('/static', StaticFiles(directory=os.path.join(BASE_DIR, 'static')),
          name='static')


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


class CuotaReq(BaseModel):
    paciente: str = ''
    telefono: str = ''
    tratamiento: str = ''
    monto_total: Optional[float] = None
    n_cuotas: Optional[int] = None
    pagadas: Optional[int] = None
    monto_cuota: Optional[float] = None
    prox_fecha: str = ''
    estado: str = ''
    nota: str = ''


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
    maestro_ok = am.MAESTRO_FID or os.path.exists(am.ruta_maestro_local())
    return {
        'credenciales_ok': am.credenciales_disponibles(),
        'credenciales': am.CREDENCIALES,
        'maestro_ok': maestro_ok,
        'maestro': am.MAESTRO_FID or am.MAESTRO,
        'maestro_modificado': None,
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
    if not (am.MAESTRO_FID or os.path.exists(am.ruta_maestro_local())):
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
    try:
        _crear_backup()
    except Exception:  # noqa: BLE001
        pass
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
    ruta = os.path.join(am.TMP_DIR, 'BD DATA_upload.xlsx')
    with open(ruta, 'wb') as f:
        f.write(contenido)
    am.subir_maestro(ruta)
    return {'ok': True, 'maestro': am.MAESTRO_FID or am.MAESTRO}


@app.get('/api/maestro/download')
async def descargar_maestro():
    ruta = am.ruta_maestro_local()
    if not os.path.exists(ruta):
        raise HTTPException(404, 'No hay maestro subido')
    return FileResponse(ruta, filename='BD DATA.xlsx')


# ============================================================
# Analítica interactiva (maestro BD DATA)
# ============================================================
def _filtros(mes: str = '', anio: str = '', desde: str = '', hasta: str = ''):
    m = (mes or 'AGO').upper().replace('SEP', 'SET')
    if m not in rv.NOMBRES_MES:
        raise HTTPException(400, f'Mes inválido: {m}')
    try:
        a = int(anio) if anio not in (None, '') else datetime.now().year
        d = int(desde) if desde not in (None, '') else 1
        h = int(hasta) if hasta not in (None, '') else 31
    except (TypeError, ValueError):
        raise HTTPException(400, 'Parámetros numéricos inválidos')
    if not (1 <= d <= 31 and 1 <= h <= 31 and d <= h):
        raise HTTPException(400, 'Rango de días inválido')
    return m, a, d, h


@app.get('/api/analitica/kpis')
async def analitica_kpis(mes: str = '', anio: str = '',
                         desde: str = '', hasta: str = ''):
    m, a, d, h = _filtros(mes, anio, desde, hasta)
    try:
        return {'ok': True, **ana.kpis(m, a, d, h)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular KPIs: {e}')


@app.get('/api/analitica/serie')
async def analitica_serie(mes: str = '', anio: str = '',
                          desde: str = '', hasta: str = ''):
    m, a, d, h = _filtros(mes, anio, desde, hasta)
    try:
        return {'ok': True, **ana.serie_diaria(m, a, d, h)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular la serie: {e}')


@app.get('/api/analitica/perfil')
async def analitica_perfil(mes: str = '', anio: str = '',
                           desde: str = '', hasta: str = ''):
    m, a, d, h = _filtros(mes, anio, desde, hasta)
    try:
        return {'ok': True, **ana.perfil(m, a, d, h)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular el perfil: {e}')


@app.get('/api/analitica/comparativo')
async def analitica_comparativo(mes: str = '', anio: str = '',
                                desde: str = '', hasta: str = ''):
    m, a, d, h = _filtros(mes, anio, desde, hasta)
    try:
        return {'ok': True, **ana.comparativo(m, a, d, h)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular el comparativo: {e}')


@app.get('/api/analitica/recurrentes')
async def analitica_recurrentes(mes: str = '', anio: str = '',
                                desde: str = '', hasta: str = ''):
    m, a, d, h = _filtros(mes, anio, desde, hasta)
    try:
        return {'ok': True, **ana.recurrentes(m, a, d, h)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudieron calcular recurrentes: {e}')


@app.get('/api/analitica/historico')
async def analitica_historico():
    try:
        return {'ok': True, 'meses': ana.ventas_por_mes()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular el histórico: {e}')


@app.get('/api/analitica/paciente')
async def analitica_paciente(telefono: str = '', dni: str = ''):
    if not telefono and not dni:
        raise HTTPException(400, 'Indica teléfono o DNI del paciente')
    try:
        citas = ana.historial_paciente(telefono, dni or None)
        notas = cp.leer_notas(telefono=telefono) if telefono else []
        nombre = citas[0].get('nombre') if citas else ''
        return {'ok': True, 'telefono': telefono, 'nombre': nombre,
                'citas': citas, 'notas': notas}
    except HTTPException:
        raise
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo leer el historial del paciente: {e}')


@app.get('/api/analitica/reactivar')
async def analitica_reactivar(meses: int = 3):
    try:
        return {'ok': True, 'pacientes': ana.pacientes_a_reactivar(meses)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo calcular pacientes a reactivar: {e}')


# ============================================================
# Backups y exportaciones
# ============================================================
@app.post('/api/backup')
async def hacer_backup():
    try:
        hechas = _crear_backup()
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo hacer el backup: {e}')
    return {'ok': True, 'backups': [os.path.basename(h) for h in hechas]}


@app.get('/api/backups')
async def listar_backups():
    out = []
    for f in sorted(os.listdir(BACKUP_DIR), reverse=True):
        if f.lower().endswith('.xlsx'):
            p = os.path.join(BACKUP_DIR, f)
            out.append({'archivo': f, 'modificado': datetime.fromtimestamp(
                os.path.getmtime(p)).strftime('%d/%m/%Y %H:%M'),
                'tamano_kb': round(os.path.getsize(p) / 1024, 1)})
    return {'ok': True, 'backups': out}


@app.get('/api/backup/download/{archivo}')
async def descargar_backup(archivo: str):
    p = os.path.join(BACKUP_DIR, os.path.basename(archivo))
    if not os.path.isfile(p):
        raise HTTPException(404, 'Backup no encontrado')
    return FileResponse(p, filename=os.path.basename(p))


def _csv_response(columnas, filas, nombre):
    import csv
    import io as _io
    buf = _io.StringIO()
    w = csv.writer(buf)
    w.writerow(columnas)
    for f in filas:
        w.writerow([f.get(c) if not isinstance(f.get(c), (int, float)) or
                    isinstance(f.get(c), float) and f.get(c) == int(f.get(c))
                    else f.get(c) for c in columnas])
    return Response(buf.getvalue().encode('utf-8-sig'),
                    media_type='text/csv; charset=utf-8',
                    headers={'Content-Disposition':
                             f'attachment; filename="{nombre}.csv"'})


@app.get('/api/exportar/{tipo}')
async def exportar_csv(tipo: str):
    tipo = tipo.lower()
    try:
        if tipo == 'agendados':
            d = crm.leer_agendados()
            mapa = {'C': 'DIA', 'D': 'MES', 'E': 'AÑO', 'G': 'NOMBRE',
                    'H': 'RED SOCIAL', 'I': 'TELEFONO', 'K': 'AGENDADO POR',
                    'L': 'DIA2', 'M': 'MES3', 'N': 'AÑO4', 'O': 'CAMPAÑA',
                    'P': 'HORA', 'Q': 'ASISTENCIA'}
            filas = [{mapa.get(k, k): v for k, v in f.items()}
                     for f in d['filas']]
            cols = list(mapa.values())
            return _csv_response(cols, filas, 'agendados')
        if tipo == 'venta':
            d = crm.leer_venta()
            filas = [f for v in d['hojas'].values() for f in v['filas']]
            cols = sorted(filas[0].keys()) if filas else ['A']
            return _csv_response(cols, filas, 'venta_diaria')
        if tipo == 'cuotas':
            return _csv_response(
                ['ID', 'Paciente', 'Telefono', 'Tratamiento', 'Monto total',
                 'Cuotas', 'Pagadas', 'Saldo', 'Prox fecha', 'Estado'],
                [{'ID': c['id'], 'Paciente': c['paciente'], 'Telefono': c['telefono'],
                  'Tratamiento': c['tratamiento'], 'Monto total': c['monto_total'],
                  'Cuotas': c['n_cuotas'], 'Pagadas': c['pagadas'],
                  'Saldo': c['saldo'], 'Prox fecha': c['prox_fecha'],
                  'Estado': c['estado']} for c in cp.leer_cuotas()],
                'cuotas')
        if tipo == 'pacientes':
            filas = cp.leer_pacientes()
            cols = ['nombre', 'telefono', 'correo', 'crm', 'campana', 'citas',
                    'compras', 'total', 'proxima_cita', 'ultima_actividad', 'notas']
            return _csv_response(cols, filas, 'pacientes')
        if tipo == 'analitica':
            return _csv_response(
                ['Metrica', 'Actual', 'Mes anterior'],
                [{'Metrica': m, 'Actual': v, 'Mes anterior': ' '}
                 for m, v in ana.kpis('AGO', datetime.now().year, 1, 31).items()],
                'analitica')
    except HTTPException:
        raise
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo exportar: {e}')
    raise HTTPException(404, f'Tipo no soportado: {tipo}')


# ============================================================
# CRM: Agendados y Venta diaria (lectura / alta en Drive)
# ============================================================
def _col_order(filas):
    cols = set()
    for f in filas:
        cols.update(f.keys())
    return sorted(cols, key=lambda c: openpyxl.utils.column_index_from_string(c))


def _normalizar_cabecera(v):
    return (str(v or '').strip().upper()
            .replace('Ñ', 'N').replace('Á', 'A').replace('É', 'E')
            .replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U'))


def _valores_data(ruta, mapa):
    """Valores únicos desde la pestaña 'DATA' de un archivo de Drive.

    ``mapa`` asocia cabecera normalizada -> clave. Detecta la fila de
    cabecera (la que contiene las columnas buscadas) y lee los valores
    de las filas siguientes, igual que los dropdowns de la hoja en Drive.
    """
    try:
        if not os.path.exists(ruta):
            return {}
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        ws = wb['DATA'] if 'DATA' in wb.sheetnames else None
        if ws is None:
            return {}
        filas_ws = list(ws.iter_rows(min_row=1, max_row=8, values_only=True))
        cab = None
        cols = {}
        for r, fila in enumerate(filas_ws, start=1):
            cabeceras = {_normalizar_cabecera(v): c
                         for c, v in enumerate(fila, start=1)}
            encontradas = {k: cabeceras[k] for k in mapa if k in cabeceras}
            if encontradas:
                cab = r
                cols = encontradas
                break
        if cab is None:
            return {}
        unicos = {k: set() for k in cols}
        for fila in ws.iter_rows(min_row=cab + 1, values_only=True):
            for k, c in cols.items():
                if c <= len(fila):
                    v = fila[c - 1]
                    if isinstance(v, str) and v.strip():
                        unicos[k].add(v.strip())
        return {mapa[k]: sorted(vals) for k, vals in unicos.items()}
    except Exception:  # noqa: BLE001
        return {}
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass


def _valores_agendados():
    """Valores de CAMPAÑA y RED SOCIAL desde la pestaña 'DATA' del archivo
    AGENDADOS (la fuente de los dropdowns de la hoja en Drive)."""
    return _valores_data(os.path.join(am.TMP_DIR, 'AGENDADOS.xlsx'),
                         {'CAMPAÑAS'.replace('Ñ', 'N'): 'campana',
                          'RED SOCIAL': 'red_social'})


def _valores_venta():
    """Valores de los selects de VENTA DIARIA desde la pestaña 'DATA' del
    archivo VENTA_DIARIA (la fuente de los dropdowns de la hoja en Drive)."""
    return _valores_data(os.path.join(am.TMP_DIR, 'VENTA_DIARIA.xlsx'),
                         {'NUEVO/RECURRENTE': 'nuevo',
                          'DISTRITO/DEPARTAMENTO': 'distrito',
                          'SEXO': 'sexo',
                          'TRATAMIENTO': 'tratamiento',
                          'DOCTOR': 'doctor',
                          'STATUS': 'status',
                          'PAGO': 'pago'})


@app.get('/api/crm/agendados')
async def crm_agendados():
    try:
        data = crm.leer_agendados()
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo leer AGENDADOS de Drive: {e}')
    valores = crm.valores_unicos_agendados(data['filas'])
    extra = _valores_agendados()
    for campo in ('campana', 'red_social'):
        if extra.get(campo):
            valores[campo] = sorted(set(valores.get(campo, [])) | set(extra[campo]))
    return {
        'ok': True,
        'filas': data['filas'],
        'total': data['total'],
        'descargado': data['descargado'],
        'columnas': data['columnas'],
        'valores': valores,
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


@app.delete('/api/crm/agendados/{fila}')
async def crm_agendados_borrar(fila: int):
    with _bloqueo:
        try:
            return crm.borrar_agendado(fila)
        except ValueError as e:
            raise HTTPException(400, str(e))
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f'No se pudo borrar el agendado en Drive: {e}')


@app.put('/api/crm/agendados/{fila}')
async def crm_agendados_editar(fila: int, data: AgendadoReq):
    d = data.model_dump()
    _validar_fecha(d.get('dia'), d.get('mes'), d.get('anio'))
    _validar_fecha(d.get('dia_cita'), d.get('mes_cita'), d.get('anio_cita'),
                   requerido_mes=True)
    if not d.get('nombre') and not d.get('telefono'):
        raise HTTPException(400, 'Indica al menos el nombre o el teléfono del paciente')
    with _bloqueo:
        try:
            res = crm.editar_agendado(fila, d)
        except HTTPException:
            raise
        except ValueError as e:
            raise HTTPException(400, str(e))
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f'No se pudo editar el agendado en Drive: {e}')
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
    valores = crm.valores_unicos_venta(todas)
    extra = _valores_venta()
    for campo in extra:
        if extra[campo]:
            valores[campo] = sorted(set(valores.get(campo, [])) | set(extra[campo]))
    return {
        'ok': True,
        'hojas': hojas,
        'descargado': data['descargado'],
        'valores': valores,
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


@app.delete('/api/crm/venta/{fila}')
async def crm_venta_borrar(fila: int, hoja: str):
    with _bloqueo:
        try:
            return crm.borrar_venta(hoja, fila)
        except ValueError as e:
            raise HTTPException(400, str(e))
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f'No se pudo borrar la venta en Drive: {e}')


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
# CRM Plus: Cuotas / pagos a plazos
# ============================================================
@app.get('/api/crm/cuotas')
async def crm_cuotas(estado: str = '', telefono: str = ''):
    try:
        return {'ok': True, 'cuotas': cp.leer_cuotas(
            estado=estado or None, telefono=telefono or None)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudieron leer las cuotas (Drive): {e}')


@app.post('/api/crm/cuotas')
async def crm_cuotas_nueva(data: CuotaReq):
    if not data.paciente and not data.telefono:
        raise HTTPException(400, 'Indica el paciente o teléfono')
    if data.monto_total is not None and data.monto_total <= 0:
        raise HTTPException(400, 'El monto total debe ser mayor a 0')
    try:
        res = cp.crear_cuota(data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo guardar la cuota (Drive): {e}')
    return {'ok': True, 'cuota': res}


@app.patch('/api/crm/cuotas/{cid}')
async def crm_cuotas_actualizar(cid: int, data: CuotaReq):
    try:
        res = cp.actualizar_cuota(cid, data.model_dump(exclude_unset=True))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo actualizar la cuota (Drive): {e}')
    if not res:
        raise HTTPException(404, 'Cuota no encontrada')
    return {'ok': True, 'cuota': res}


@app.post('/api/crm/cuotas/{cid}/pago')
async def crm_cuotas_pago(cid: int):
    try:
        res = cp.registrar_pago_cuota(cid)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo registrar el pago (Drive): {e}')
    if not res:
        raise HTTPException(404, 'Cuota no encontrada')
    if res.get('error'):
        raise HTTPException(400, res['error'])
    return {'ok': True, 'cuota': res}


@app.delete('/api/crm/cuotas/{cid}')
async def crm_cuotas_borrar(cid: int):
    try:
        ok = cp.borrar_cuota(cid)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f'No se pudo borrar la cuota (Drive): {e}')
    if not ok:
        raise HTTPException(404, 'Cuota no encontrada')
    return {'ok': True}


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
