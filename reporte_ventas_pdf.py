#!/usr/bin/env python3
"""Reporte de ventas por campana de agendados en PDF (multi-pagina).

Mismos criterios que los pivotes del maestro BD DATA.xlsx:
  - Hoja AGENDADO : Agendados = filas con FECHA DE AGENDADO (DIA/MES/ANO, C/D/E)
                    dentro del periodo y con telefono.
  - Hoja ASISTIDO : Asistidos = filas con FECHA DE CITA (DIA2/MES3/ANO4, L/M/N)
                    dentro del periodo y ASISTENCIA = ASISTIO.
                    Compraron = de los asistidos, los que tienen PAGO TOTAL > 0.
                    Monto = suma de PAGO TOTAL.

Paginas del PDF:
  1. Resumen comercial   (KPIs, tasas, resumen por CRM)
  2. Flujo operativo     (embudo agendados -> venta y ticket promedio)
  3. Resumen de campanas (tablas por campana y por CRM)
  4. Campanas por canal  (graficas apiladas y detalle Kommo/WhatsApp/Organico)
  5. Metricas            (tratamientos, distritos, perfil del paciente)
  6. Hallazgos y sugerencias

Uso:
  python3 reporte_ventas_pdf.py                       # AGO 1-10, maestro directo
  python3 reporte_ventas_pdf.py --desde 11 --hasta 11 --fuente auto \
      --salida Reporte_Ventas_11_Agosto_2026.pdf
"""
import argparse
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import alimentar_maestro as am
import meta_ads as mads

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(BASE_DIR, 'Reporte_Ventas_1-10_Agosto_2026.pdf')
MES = 'AGO'
ANIO = 2026
D1, D2 = 1, 10
FUENTE = 'maestro'   # maestro = pivotes guardados | auto = integra AGENDADOS+VENTA (Drive)

FONDO_TITULO = '#1f3864'
AZUL = '#2f5597'
CELESTE = '#9dc3e6'
VERDE = '#70ad47'
NARANJA = '#ed7d31'
GRIS = '#595959'
ROJO = '#c00000'
FONDO_CLARO = '#eaf1f8'

CRM_ORDEN = ['KOMMO', 'WHATSAPP', 'ORGANICO', 'SIN CRM']
CANALES = ['KOMMO', 'WHATSAPP', 'ORGANICO']
COLOR_CANAL = {'KOMMO': AZUL, 'WHATSAPP': VERDE, 'ORGANICO': NARANJA}
PALETA = [AZUL, VERDE, NARANJA, CELESTE]


def canales_desde_datos(agg):
    """Canales reales presentes en los datos (top 3 por agendados), sin SIN CRM."""
    totales = Counter()
    for (_, crm), d in agg.items():
        totales[crm] += d['ag']
    top = [c for c, _ in totales.most_common(3) if str(c).strip().upper() != 'SIN CRM']
    return top or ['WHATSAPP']


# ============================================================
# Campañas reales vs "Otros" (no son campañas de anuncios)
# ============================================================
# Valores que puede tener la columna CAMPAÑA del maestro sin ser una campaña
# de verdad: tipos de venta que se registran por costumbre en esa celda, no
# nombres de campaña de Meta Ads.
CATEGORIAS_OTROS = {
    'EVALUACION', 'RETOQUE', 'RECURRENTE', 'RECOMENDADO', 'SESION',
    'ORGANICO REDES', 'VENTA SIN AGENDAR',
}
_RE_HORA = re.compile(r'^\d{1,2}[:.]\d{2}\s*(AM|PM)?$', re.I)


def _normalizar_categoria(s):
    s = unicodedata.normalize('NFKD', str(s or '').upper())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'[^A-Z0-9]+', ' ', s).strip()


def campana_valida(nombre):
    """False si ``nombre`` no es una campaña real: vacío, un tipo de venta
    (evaluación, retoque, recurrente...) o un dato corrupto (p.ej. una hora
    tipeada por error en la columna CAMPAÑA)."""
    s = str(nombre or '').strip()
    if not s or s == '(SIN CAMPANA)':
        return False
    if _RE_HORA.match(s):
        return False
    return _normalizar_categoria(s) not in CATEGORIAS_OTROS


def categoria_otros(nombre):
    """Nombre de categoría a mostrar en la tabla 'Otros' para un valor de
    CAMPAÑA que no pasó ``campana_valida``."""
    s = str(nombre or '').strip()
    if not s or s == '(SIN CAMPANA)':
        return 'SIN CAMPAÑA'
    if _RE_HORA.match(s):
        return 'SIN CAMPAÑA (dato inválido)'
    return s


def separar_campanas_otros(agg):
    """Divide un agg {(campana, crm): datos} en (campanas_reales, otros),
    agrupando en 'otros' lo que no es una campaña de anuncios (ver
    ``campana_valida``) para que no ensucie las tablas de campañas."""
    campanas, otros = {}, {}
    for (camp, crm), d in agg.items():
        if campana_valida(camp):
            campanas[(camp, crm)] = d
        else:
            cat = categoria_otros(camp)
            dest = otros.setdefault((cat, crm), dict(ag=0, as_=0, co=0, mon=0.0,
                                                      no_fueron=0, fueron_sin_compra=0))
            for k in dest:
                dest[k] += d.get(k, 0)
    return campanas, otros


COL_BM = {
    'CRM': 2, 'RED_SOCIAL': 7, 'DIA': 3, 'MES': 4, 'ANIO': 5, 'TELEFONO': 9,
    'DIA2': 12, 'MES3': 13, 'ANIO4': 14, 'CAMPANA': 15, 'ASISTENCIA': 16,
    'DISTRITO': 17, 'EDAD': 18, 'SEXO': 19,
    'TRAT': [20, 22, 24, 26], 'PAGO': [21, 23, 25, 27], 'PAGO_TOTAL': 28,
}
_CABECERA = {
    'DIA': 'DIA', 'MES': 'MES', 'AÑO': 'ANIO', 'TELEFONO': 'TELEFONO',
    'DIA2': 'DIA2', 'MES3': 'MES3', 'AÑO4': 'ANIO4', 'CAMPAÑA': 'CAMPANA',
    'ASISTENCIA': 'ASISTENCIA', 'DISTRITO': 'DISTRITO', 'EDAD': 'EDAD',
    'SEXO': 'SEXO', 'CRM': 'CRM', 'RED SOCIAL': 'RED_SOCIAL',
    'PAGO TOTAL': 'PAGO_TOTAL', 'DNI': 'DNI', 'NOMBRE': 'NOMBRE',
    'CORREO': 'CORREO',
    'AGENDADO POR': 'AGENDADO',
    'MOTIVO NO ASISTIO': 'MOTIVO_NO_ASISTIO',
    'MOTIVO NO COMPRA': 'MOTIVO_NO_COMPRA',
}


def detectar_columnas(ws):
    """Ubica las columnas del maestro por su cabecera (fila 4), tolerando el
    formato BM (con columna CRM) y el Derma Essenza (sin CRM). Si una columna
    no se detecta, usa la posición del formato BM."""
    col = dict(COL_BM)
    hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
    pos = {}
    for i, h in enumerate(hdr, 1):
        if h is None:
            continue
        hh = str(h).strip().upper()
        if hh in _CABECERA:
            pos[_CABECERA[hh]] = i
        elif hh.startswith('TRAT '):
            pos.setdefault('TRAT', []).append((int(hh[5:]), i))
        elif hh.startswith('PAGO '):
            pos.setdefault('PAGO', []).append((int(hh[5:]), i))
    for k, v in pos.items():
        col[k] = [i for _, i in sorted(v)] if isinstance(v, list) else v
    col['ES_BM'] = 'CRM' in pos
    col['CANAL'] = pos.get('CRM') or pos.get('RED_SOCIAL') or COL_BM['CRM']
    return col


COL = dict(COL_BM)

NOMBRES_MES = {'ENE': 'Enero', 'FEB': 'Febrero', 'MAR': 'Marzo', 'ABR': 'Abril',
               'MAY': 'Mayo', 'JUN': 'Junio', 'JUL': 'Julio', 'AGO': 'Agosto',
               'SET': 'Septiembre', 'OCT': 'Octubre', 'NOV': 'Noviembre', 'DIC': 'Diciembre'}


def en_periodo(dia):
    return isinstance(dia, (int, float)) and D1 <= int(dia) <= D2


def pago_total(ws, r):
    p = ws.cell(row=r, column=COL['PAGO_TOTAL']).value
    s = sum(x for x in (ws.cell(row=r, column=c).value for c in COL['PAGO'])
            if isinstance(x, (int, float)))
    if isinstance(p, (int, float)) and p > 0:
        return p
    return s


def _cita_pasada(ws, r):
    """True si la fecha de cita de la fila ya ocurrió (comparada con hoy)."""
    try:
        anio = int(ws.cell(row=r, column=COL['ANIO4']).value)
        mes = ws.cell(row=r, column=COL['MES3']).value
        dia = int(ws.cell(row=r, column=COL['DIA2']).value)
        if not mes:
            return False
        mes = str(mes).strip().upper()
        if mes == 'SEP':
            mes = 'SET'
        if mes not in NOMBRES_MES:
            return False
        hoy = datetime.now()
        mi = list(NOMBRES_MES).index(mes) + 1
        return (anio, mi, dia) < (hoy.year, hoy.month, hoy.day)
    except (TypeError, ValueError):
        return False


def monto(sol):
    return f"S/ {sol:,.0f}" if sol else "S/ 0"


def build_data():
    global COL
    if FUENTE == 'auto':
        maestro_ws = am.leer_maestro(am.ruta_maestro_local())
        COL = detectar_columnas(maestro_ws)
        agendados, ag_col = am.leer_agendados(os.path.join(am.TMP_DIR, 'AGENDADOS.xlsx'))
        venta = am.leer_venta(os.path.join(am.TMP_DIR, 'VENTA_DIARIA.xlsx'))
        col = am.detectar_maestro(maestro_ws)
        calc = am.Calculo(maestro_ws, agendados, venta, col=col, ag_col=ag_col)
        tmp = os.path.join(am.TMP_DIR, 'simulado_reporte.xlsx')
        am.aplicar_xml(am.ruta_maestro_local(), tmp, calc.new_rows, calc.updates,
                       col=col, ag_col=ag_col)
        ws = openpyxl.load_workbook(tmp, data_only=True)['BD DATA']
    else:
        ws = am.leer_maestro(am.ruta_maestro_local())
        COL = detectar_columnas(ws)
    agg = defaultdict(lambda: dict(ag=0, as_=0, co=0, mon=0.0,
                                   no_fueron=0, fueron_sin_compra=0))
    for r in range(5, ws.max_row + 1):
        camp = str(ws.cell(row=r, column=COL['CAMPANA']).value or '').strip() or '(SIN CAMPANA)'
        crm = ws.cell(row=r, column=COL['CANAL']).value or 'SIN CRM'
        d = agg[(camp, crm)]
        if (ws.cell(row=r, column=COL['ANIO']).value == ANIO
                and ws.cell(row=r, column=COL['MES']).value == MES
                and en_periodo(ws.cell(row=r, column=COL['DIA']).value)
                and ws.cell(row=r, column=COL['TELEFONO']).value):
            d['ag'] += 1
            asist = str(ws.cell(row=r, column=COL['ASISTENCIA']).value or '').strip()
            if asist != 'ASISTIO' and _cita_pasada(ws, r):
                d['no_fueron'] += 1
        if (ws.cell(row=r, column=COL['ANIO4']).value == ANIO
                and ws.cell(row=r, column=COL['MES3']).value == MES
                and en_periodo(ws.cell(row=r, column=COL['DIA2']).value)
                and ws.cell(row=r, column=COL['ASISTENCIA']).value == 'ASISTIO'):
            d['as_'] += 1
            p = pago_total(ws, r)
            d['mon'] += p
            if p > 0:
                d['co'] += 1
            else:
                d['fueron_sin_compra'] += 1
    return agg


def datos_analiticos():
    """Metricas de perfil de los asistentes del periodo."""
    global COL
    ws = am.leer_maestro(am.ruta_maestro_local())
    COL = detectar_columnas(ws)
    trat = Counter(); dist = Counter(); sexo = Counter()
    edades = []; montos = []; crm = Counter(); camp = Counter()
    for r in range(5, ws.max_row + 1):
        if (ws.cell(row=r, column=COL['ANIO4']).value == ANIO
                and ws.cell(row=r, column=COL['MES3']).value == MES
                and en_periodo(ws.cell(row=r, column=COL['DIA2']).value)
                and ws.cell(row=r, column=COL['ASISTENCIA']).value == 'ASISTIO'):
            crm[str(ws.cell(row=r, column=COL['CANAL']).value or 'SIN CRM')] += 1
            camp[str(ws.cell(row=r, column=COL['CAMPANA']).value or '').strip() or '(SIN)'] += 1
            for cc in COL['TRAT']:
                t = ws.cell(row=r, column=cc).value
                if t:
                    trat[str(t).strip().upper()] += 1
            d = ws.cell(row=r, column=COL['DISTRITO']).value
            if d:
                dist[str(d).strip().title()] += 1
            e = ws.cell(row=r, column=COL['EDAD']).value
            if isinstance(e, (int, float)):
                edades.append(e)
            s = ws.cell(row=r, column=COL['SEXO']).value
            if s:
                sexo[str(s).strip().upper()] += 1
            p = pago_total(ws, r)
            if p > 0:
                montos.append(p)
    return {'trat': trat, 'dist': dist, 'sexo': sexo, 'edades': edades,
            'montos': montos, 'crm': crm, 'camp': camp}


def datos_ejecutivas():
    """Ranking por ejecutiva (AGENDADO POR) del periodo, desde el maestro."""
    global COL
    ws = am.leer_maestro(am.ruta_maestro_local())
    COL = detectar_columnas(ws)
    c_ag = COL.get('AGENDADO')
    if not c_ag:
        return []
    c_ej = c_ag
    agrup = {}
    for r in range(5, ws.max_row + 1):
        ej = str(ws.cell(row=r, column=c_ej).value or '').strip() or 'SIN EJECUTIVA'
        d = agrup.setdefault(ej, {'ag': 0, 'as_': 0, 'co': 0, 'mon': 0.0,
                                  'no_fueron': 0, 'fueron_sin_compra': 0})
        if (ws.cell(row=r, column=COL['ANIO']).value == ANIO
                and ws.cell(row=r, column=COL['MES']).value == MES
                and en_periodo(ws.cell(row=r, column=COL['DIA']).value)
                and ws.cell(row=r, column=COL['TELEFONO']).value):
            d['ag'] += 1
            asist = str(ws.cell(row=r, column=COL['ASISTENCIA']).value or '').strip()
            if asist != 'ASISTIO' and _cita_pasada(ws, r):
                d['no_fueron'] += 1
        if (ws.cell(row=r, column=COL['ANIO4']).value == ANIO
                and ws.cell(row=r, column=COL['MES3']).value == MES
                and en_periodo(ws.cell(row=r, column=COL['DIA2']).value)
                and ws.cell(row=r, column=COL['ASISTENCIA']).value == 'ASISTIO'):
            d['as_'] += 1
            p = pago_total(ws, r)
            d['mon'] += p
            if p > 0:
                d['co'] += 1
            else:
                d['fueron_sin_compra'] += 1
    out = []
    for ej, d in agrup.items():
        out.append({'ejecutiva': ej, 'ag': d['ag'], 'as_': d['as_'], 'co': d['co'],
                    'mon': round(d['mon'], 2), 'no_fueron': d['no_fueron'],
                    'fueron_sin_compra': d['fueron_sin_compra'],
                    'asistencia_pct': pct(d['as_'], d['ag']),
                    'conversion_pct': pct(d['co'], d['as_']),
                    'ticket': d['mon'] / d['co'] if d['co'] else 0})
    out.sort(key=lambda x: (x['mon'], x['ag']), reverse=True)
    return out


def datos_motivos():
    """Distribución de motivos de no asistencia y no compra del periodo."""
    global COL
    ws = am.leer_maestro(am.ruta_maestro_local())
    COL = detectar_columnas(ws)
    c_m_as = COL.get('MOTIVO_NO_ASISTIO')
    c_m_co = COL.get('MOTIVO_NO_COMPRA')
    no_asistio = Counter()
    no_compra = Counter()
    for r in range(5, ws.max_row + 1):
        agendado_ok = (ws.cell(row=r, column=COL['ANIO']).value == ANIO
                       and ws.cell(row=r, column=COL['MES']).value == MES
                       and en_periodo(ws.cell(row=r, column=COL['DIA']).value))
        asistido_ok = (ws.cell(row=r, column=COL['ANIO4']).value == ANIO
                       and ws.cell(row=r, column=COL['MES3']).value == MES
                       and en_periodo(ws.cell(row=r, column=COL['DIA2']).value))
        asist = str(ws.cell(row=r, column=COL['ASISTENCIA']).value or '').strip()
        if agendado_ok and asist != 'ASISTIO' and _cita_pasada(ws, r) and c_m_as:
            m = str(ws.cell(row=r, column=c_m_as).value or '').strip()
            if m:
                no_asistio[m.upper()] += 1
        if asistido_ok and asist == 'ASISTIO' and pago_total(ws, r) == 0 and c_m_co:
            m = str(ws.cell(row=r, column=c_m_co).value or '').strip()
            if m:
                no_compra[m.upper()] += 1
    return {'no_asistio': dict(no_asistio.most_common()),
            'no_compra': dict(no_compra.most_common())}


def _periodo_anterior(mes, anio, desde, hasta):
    """Mes previo con el mismo rango de días (variación comparable)."""
    orden = list(NOMBRES_MES)
    i = orden.index(mes)
    if i == 0:
        return orden[-1], anio - 1, desde, hasta
    return orden[i - 1], anio, desde, hasta


def _variacion(actual, anterior):
    if anterior is None or anterior == 0:
        return None
    return round((actual - anterior) / anterior * 100.0, 1)


def totales_periodo(mes, anio, desde, hasta):
    """Totales del periodo sobre el maestro actual (sin tocar globales)."""
    ws = am.leer_maestro(am.ruta_maestro_local())
    col = detectar_columnas(ws)
    tot = {'ag': 0, 'as_': 0, 'co': 0, 'mon': 0.0,
           'no_fueron': 0, 'fueron_sin_compra': 0}
    for r in range(5, ws.max_row + 1):
        if (ws.cell(row=r, column=col['ANIO']).value == anio
                and ws.cell(row=r, column=col['MES']).value == mes
                and isinstance(ws.cell(row=r, column=col['DIA']).value, (int, float))
                and desde <= int(ws.cell(row=r, column=col['DIA']).value) <= hasta
                and ws.cell(row=r, column=col['TELEFONO']).value):
            tot['ag'] += 1
            asist = str(ws.cell(row=r, column=col['ASISTENCIA']).value or '').strip()
            if asist != 'ASISTIO' and _cita_pasada(ws, r):
                tot['no_fueron'] += 1
        if (ws.cell(row=r, column=col['ANIO4']).value == anio
                and ws.cell(row=r, column=col['MES3']).value == mes
                and isinstance(ws.cell(row=r, column=col['DIA2']).value, (int, float))
                and desde <= int(ws.cell(row=r, column=col['DIA2']).value) <= hasta
                and ws.cell(row=r, column=col['ASISTENCIA']).value == 'ASISTIO'):
            tot['as_'] += 1
            p = pago_total(ws, r)
            tot['mon'] += p
            if p > 0:
                tot['co'] += 1
            else:
                tot['fueron_sin_compra'] += 1
    return tot


def pct(a, b):
    return 100.0 * a / b if b else 0.0


# ============================================================
# Integración Meta Ads + maestro (tabla por campaña)
# ============================================================
META_A_MAESTRO = {
    'TOXINA 2026': 'TOXINA FULL FACE',
    'CONSULTA GRATIS 2026': 'CONSULTA GRATUITA',
    'ACIDO HIALURONICO': 'ACIDO HIALURONICO',
}


def _ruta_meta_dir():
    base = os.environ.get('DATA_DIR', os.path.join(BASE_DIR, 'data'))
    return os.path.join(base, 'meta_ads')


def _norm_tokens(s):
    s = unicodedata.normalize('NFKD', str(s).upper())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return set(t for t in s.split() if len(t) >= 3)


def _fuzzy_maestro(meta_nombre, camp_maestros):
    """Encuentra la campaña del maestro más parecida por tokens compartidos."""
    mt = _norm_tokens(meta_nombre) - {'2026'}
    if not mt:
        return None
    best, best_n = None, 0
    for cm in camp_maestros:
        n = len(mt & _norm_tokens(cm))
        if n > best_n:
            best, best_n = cm, n
    return best if best_n > 0 else None


def _cargar_meta_campanas():
    """Por campaña de la carga de Meta Ads más reciente. None si no hay cargas."""
    try:
        idx = mads.listar(_ruta_meta_dir())
    except Exception:
        return None
    if not idx:
        return None
    try:
        det = mads.detalle(_ruta_meta_dir(), idx[0]['id'])
    except Exception:
        return None
    return det.get('por_campania') or []


def _fila_meta(nombre, mc, d):
    gasto = mc.get('gasto') or 0
    leads = mc.get('resultados') or 0
    costo_res = mc.get('costo_resultado')
    if costo_res is None:
        costo_res = round(gasto / leads, 2) if leads else 0
    ag, as_, co, mon = d['ag'], d['as_'], d['co'], d['mon']
    return dict(campania=nombre, gasto=gasto, leads=leads, costo_res=costo_res,
                ag=ag, pct_ag=pct(ag, leads), as_=as_, pct_as=pct(as_, ag),
                co=co, mon=mon, ticket=mon / as_ if as_ else 0, organica=False)


def _gasto_ads_total():
    por_meta = _cargar_meta_campanas()
    if not por_meta:
        return None
    return round(sum((mc.get('gasto') or 0) for mc in por_meta), 2)


def build_campana_meta(agg):
    """Une cada campaña de Meta Ads con su embudo del maestro (ag→as→co→monto)."""
    camps = defaultdict(lambda: dict(ag=0, as_=0, co=0, mon=0.0))
    for (camp, crm), d in agg.items():
        for k in ('ag', 'as_', 'co', 'mon'):
            camps[camp][k] += d[k]
    cero = dict(ag=0, as_=0, co=0, mon=0.0)
    filas = []
    usadas = set()
    por_meta = _cargar_meta_campanas()
    if por_meta is not None:
        for mc in por_meta:
            nombre = mc.get('campania') or '—'
            maes = META_A_MAESTRO.get(nombre) or _fuzzy_maestro(nombre, list(camps))
            d = camps.get(maes, cero) if maes else cero
            if maes:
                usadas.add(maes)
            filas.append(_fila_meta(nombre, mc, d))
        filas.sort(key=lambda x: -x['gasto'])
    for camp, d in sorted(camps.items(), key=lambda kv: -kv[1]['ag']):
        if camp in usadas or (not d['ag'] and not d['as_']):
            continue
        filas.append(dict(campania=camp + ' (organica)', gasto=0, leads=0,
                          costo_res=0, ag=d['ag'], pct_ag=0,
                          as_=d['as_'], pct_as=pct(d['as_'], d['ag']),
                          co=d['co'], mon=d['mon'],
                          ticket=d['mon'] / d['as_'] if d['as_'] else 0,
                          organica=True))
    return filas


def estilo_tabla(ax, datos, col_w, header=None, color_titulo=FONDO_TITULO,
                 fontsize=8.5, scale=1.7, alinear_izq_col0=True, altura_fila=None,
                 max_row_h=None):
    """scale ahora limita la altura de fila (fraccion de la altura del eje);
    la tabla siempre llena el eje en horizontal, y en vertical hasta ese
    tope, quedando pegada arriba en vez de centrada con un hueco grande."""
    ax.axis('off')
    n_filas = len(datos) + (1 if header else 0)
    tope = max_row_h if max_row_h is not None else 0.11 * (scale / 1.7)
    alto_tabla = min(1.0, n_filas * tope)
    tabla = ax.table(cellText=datos, colLabels=header, cellLoc='center',
                     bbox=[0, 1 - alto_tabla, 1, alto_tabla], colWidths=col_w)
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(fontsize)
    for (i, j), cel in tabla.get_celld().items():
        cel.set_edgecolor('#bfbfbf')
        if altura_fila:
            cel.set_height(altura_fila)
        if i == 0:
            cel.set_facecolor(color_titulo)
            cel.set_text_props(color='white', fontweight='bold')
        elif alinear_izq_col0 and j == 0:
            cel.set_text_props(fontweight='bold', ha='left')
        elif i % 2 == 0:
            cel.set_facecolor('#f2f2f2')
    return tabla


def caja_kpi(fig, x, y, w, h, titulo, valor, sub, color, var_pct=None):
    ax = fig.add_axes([x, y, w, h]); ax.axis('off')
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=color, edgecolor='none',
                               transform=ax.transAxes))
    fs = 24 if len(valor) <= 3 else 22
    ax.text(0.5, 0.72, valor, ha='center', va='center', fontsize=fs,
            fontweight='bold', color='white', transform=ax.transAxes)
    ax.text(0.5, 0.44, titulo, ha='center', va='center', fontsize=10,
            color='white', fontweight='bold', transform=ax.transAxes)
    sub_txt = sub
    if var_pct is not None:
        signo = '▲ +' if var_pct > 0 else ('▼ ' if var_pct < 0 else '▶ ')
        sub_txt = f"{signo}{var_pct:+.1f}% vs periodo ant.\n{sub}"
    ax.text(0.5, 0.16, sub_txt, ha='center', va='center', fontsize=7.5,
            color='white', alpha=0.9, transform=ax.transAxes)


def pagina_resumen(pdf, tot, agg, variacion=None, gasto_ads=None):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')

    ax.text(0.5, 0.96, 'REPORTE DE VENTAS', ha='center', fontsize=24,
            fontweight='bold', color=FONDO_TITULO, transform=fig.transFigure)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO} - Campaña de agendados',
            ha='center', fontsize=12, color=GRIS, transform=fig.transFigure)
    ax.add_line(plt.Line2D([0.15, 0.85], [0.90, 0.90], transform=fig.transFigure,
                           color=CELESTE, lw=2))

    v = variacion or {}
    kpis = [
        ('AGENDADOS', str(tot['ag']), 'citas agendadas', AZUL, v.get('ag')),
        ('ASISTIERON', str(tot['as_']), f"{pct(tot['as_'], tot['ag']):.0f}% de asistencia", VERDE, v.get('as_')),
        ('COMPRARON', str(tot['co']), f"{pct(tot['co'], tot['as_']):.0f}% de conversion", NARANJA, v.get('co')),
        ('MONTO TOTAL', monto(tot['mon']), 'ventas del periodo', FONDO_TITULO, v.get('mon')),
    ]
    n = len(kpis)
    w = 0.21
    gap = (0.92 - n * w) / (n + 1)
    for i, (t, val, s, c, vp) in enumerate(kpis):
        x = gap + i * (w + gap)
        caja_kpi(fig, x, 0.70, w, 0.16, t, val, s, c, var_pct=vp)

    ax.add_patch(plt.Rectangle((0.06, 0.46), 0.88, 0.17, facecolor=FONDO_CLARO,
                               edgecolor=CELESTE, linewidth=1, transform=fig.transFigure))
    ax.text(0.09, 0.605, 'Indicadores clave', fontsize=11, fontweight='bold',
            color=FONDO_TITULO, transform=fig.transFigure)
    if tot['co']:
        ticket = tot['mon'] / tot['co']
        lineas = [
            f"Ticket promedio: S/ {ticket:,.0f} por compra  |  Asistencia: {pct(tot['as_'], tot['ag']):.0f}%  |  Conversion a compra: {pct(tot['co'], tot['as_']):.0f}%",
            f"Agendaron pero no fueron: {tot['no_fueron']}  |  Fueron pero no compraron: {tot['fueron_sin_compra']}.",
            f"Cada {tot['ag']} agendados generan {tot['co']} ventas por S/ {tot['mon']:,.0f}.",
        ]
        if gasto_ads and gasto_ads > 0:
            roas = tot['mon'] / gasto_ads
            cac = gasto_ads / tot['co']
            lineas.append(f"Inversión Meta Ads: S/ {gasto_ads:,.0f}  |  ROAS: {roas:.2f}x  |  CAC: S/ {cac:,.0f} por compra")
    else:
        lineas = ['No se registraron ventas en el periodo.',
                  f"Agendaron pero no fueron: {tot['no_fueron']}  |  Fueron pero no compraron: {tot['fueron_sin_compra']}."]
        if gasto_ads and gasto_ads > 0:
            lineas.append(f"Inversión Meta Ads: S/ {gasto_ads:,.0f} (sin ventas registradas aún)")
    ax.text(0.09, 0.565, lineas[0], fontsize=9.5, color=GRIS, transform=fig.transFigure)
    if len(lineas) > 1:
        ax.text(0.09, 0.535, lineas[1], fontsize=9.5, color=GRIS, transform=fig.transFigure)
    if len(lineas) > 2:
        ax.text(0.09, 0.505, lineas[2], fontsize=9.5, color=GRIS, transform=fig.transFigure)
    if len(lineas) > 3:
        ax.text(0.09, 0.475, lineas[3], fontsize=9.5, color=AZUL, fontweight='bold', transform=fig.transFigure)

    ax.text(0.08, 0.415, 'Resumen por CRM', fontsize=12, fontweight='bold',
            color=FONDO_TITULO, transform=fig.transFigure)
    cr = []
    for c in CRM_ORDEN:
        if c in agg_tot_crm:
            d = agg_tot_crm[c]
            cr.append([c, d['ag'], d['as_'], d['co'],
                       f"{pct(d['co'], d['as_']):.0f}%", monto(d['mon'])])
    cr.append(['TOTAL', tot['ag'], tot['as_'], tot['co'],
               f"{pct(tot['co'], tot['as_']):.0f}%", monto(tot['mon'])])
    estilo_tabla(fig.add_axes([0.08, 0.08, 0.84, 0.32]), cr,
                 [0.26, 0.15, 0.15, 0.14, 0.14, 0.16],
                 header=['CRM', 'Agendados', 'Asistieron', 'Compraron', 'Conv.', 'Monto'],
                 fontsize=9, scale=1.5)

    ax.text(0.08, 0.06, 'Fuente: maestro BD DATA.xlsx (hojas AGENDADO y ASISTIDO, '
                         'mismas cifras de los pivotes).', fontsize=8, color=GRIS,
            transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_flujo(pdf, tot):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Flujo operativo: del agendado a la venta', ha='center',
            fontsize=16, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    etapas = [('AGENDADOS', tot['ag'], AZUL),
              ('ASISTIERON', tot['as_'], VERDE),
              ('COMPRARON', tot['co'], NARANJA)]
    maxv = max(tot['ag'], 1)
    xs = [0.12, 0.17, 0.22]
    y = 0.80
    bar_h = 0.085
    for i, (nom, val, c) in enumerate(etapas):
        w = max(0.16, 0.62 * val / maxv)
        x0 = xs[i]
        x1 = x0 + w
        ax.add_patch(plt.Rectangle((x0, y), w, bar_h, facecolor=c, edgecolor='none',
                                   transform=fig.transFigure, clip_on=False))
        ax.text(x1 + 0.02, y + bar_h / 2, f'{nom}: {val}', ha='left', va='center',
                fontsize=13, fontweight='bold', color=c, transform=fig.transFigure)
        if i < len(etapas) - 1:
            conv = pct(etapas[i + 1][1], val)
            y_conv = y - 0.06
            ax.annotate('', xy=(x1 - 0.03, y - 0.013), xytext=(x0 + 0.03, y - 0.013),
                        arrowprops=dict(arrowstyle='->', color=GRIS, lw=1.2),
                        transform=fig.transFigure)
            ax.text(0.06, y_conv, f'{conv:.0f}%', ha='left', va='center', fontsize=13,
                    fontweight='bold', color=FONDO_TITULO, transform=fig.transFigure)
            ax.text(0.12, y_conv - 0.030, 'convierte a la siguiente etapa', ha='left',
                    va='center', fontsize=9, color=GRIS, transform=fig.transFigure)
        y -= 0.23

    ax.add_patch(plt.Rectangle((0.16, 0.155), 0.68, 0.14, facecolor=FONDO_TITULO,
                               edgecolor='none', transform=fig.transFigure, clip_on=False))
    ax.text(0.5, 0.255, monto(tot['mon']), ha='center', va='center', fontsize=26,
            fontweight='bold', color='white', transform=fig.transFigure)
    ax.text(0.5, 0.215, 'MONTO VENDIDO EN EL PERIODO', ha='center', va='center',
            fontsize=10, color='white', transform=fig.transFigure)

    if tot['co']:
        ticket = tot['mon'] / tot['co']
        lineas = [
            f"Ticket promedio por compra: S/ {ticket:,.0f}",
            f"Ventas por cada asistente: S/ {tot['mon'] / max(tot['as_'], 1):,.0f}",
            f"{pct(tot['co'], tot['ag']):.0f}% de los agendados termino comprando.",
        ]
    else:
        lineas = ['No hubo ventas en el periodo.']
    ax.text(0.5, 0.115, '   |   '.join(lineas), ha='center', fontsize=9.5,
            color=GRIS, transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_campanas(pdf, agg):
    """``agg`` debe ser solo campañas reales (ver ``separar_campanas_otros``);
    los tipos de venta que no son campaña van en ``pagina_otros``."""
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Resumen de campanas', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    por_crm = defaultdict(dict)
    camps = defaultdict(lambda: dict(ag=0, as_=0, co=0, mon=0.0))
    for (camp, crm), d in agg.items():
        por_crm[crm][camp] = d
        for k in ('ag', 'as_', 'co', 'mon'):
            camps[camp][k] += d[k]
    camp_ag = sum(d['ag'] for d in camps.values())
    camp_as = sum(d['as_'] for d in camps.values())
    camp_co = sum(d['co'] for d in camps.values())
    camp_mon = sum(d['mon'] for d in camps.values())

    camp_list = sorted(camps.items(), key=lambda kv: -kv[1]['ag'])
    filas = [[c, d['ag'], d['as_'], d['co'], f"{pct(d['co'], d['as_']):.0f}%",
              monto(d['mon'])] for c, d in camp_list[:8]]
    if len(camp_list) > 8:
        resto = [d for _, d in camp_list[8:]]
        filas.append(['Otras campanas', sum(d['ag'] for d in resto),
                      sum(d['as_'] for d in resto),
                      sum(d['co'] for d in resto),
                      f"{pct(sum(d['co'] for d in resto), sum(d['as_'] for d in resto)):.0f}%",
                      monto(sum(d['mon'] for d in resto))])
    filas.append(['TOTAL', camp_ag, camp_as, camp_co,
                  f"{pct(camp_co, camp_as):.0f}%", monto(camp_mon)])

    if not camps:
        ax.text(0.5, 0.6, 'No hubo campañas de anuncios en este periodo.\n'
                          'Ver "Otros" para evaluaciones, retoques, recurrentes, etc.',
                ha='center', fontsize=10, color=GRIS, transform=fig.transFigure)
        pdf.savefig(fig); plt.close(fig)
        return

    heading_y = 0.895
    axc_top = 0.87
    alto_camp = min(0.50, max(0.12, 0.043 * (len(filas) + 1)))
    ax.text(0.08, heading_y, 'Por campana (principales)', fontsize=12,
            fontweight='bold', color=FONDO_TITULO)
    axc = fig.add_axes([0.08, axc_top - alto_camp, 0.84, alto_camp])
    estilo_tabla(axc, filas, [0.36, 0.13, 0.13, 0.13, 0.11, 0.17],
                 header=['Campana', 'Agend.', 'Asist.', 'Compr.', 'Conv.', 'Monto'],
                 fontsize=9, scale=1.0)

    crm_activos = [c for c in CRM_ORDEN if c in por_crm]
    crm_filas = []
    for c in crm_activos:
        v = por_crm[c]
        ag = sum(d['ag'] for d in v.values())
        as_ = sum(d['as_'] for d in v.values())
        co = sum(d['co'] for d in v.values())
        mon = sum(d['mon'] for d in v.values())
        crm_filas.append([c, ag, as_, co, f"{pct(co, as_):.0f}%", monto(mon)])
    crm_filas.append(['TOTAL', camp_ag, camp_as, camp_co,
                      f"{pct(camp_co, camp_as):.0f}%", monto(camp_mon)])

    crm_heading_y = axc_top - alto_camp - 0.05
    alto_crm = min(0.30, max(0.10, 0.043 * (len(crm_filas) + 1)))
    ax.text(0.08, crm_heading_y, 'Por CRM', fontsize=12, fontweight='bold', color=FONDO_TITULO)
    axd = fig.add_axes([0.08, crm_heading_y - 0.03 - alto_crm, 0.84, alto_crm])
    estilo_tabla(axd, crm_filas, [0.30, 0.15, 0.15, 0.13, 0.12, 0.18],
                 header=['CRM', 'Agend.', 'Asist.', 'Compr.', 'Conv.', 'Monto'],
                 fontsize=9, scale=1.0)
    pdf.savefig(fig); plt.close(fig)


def pagina_otros(pdf, agg_otros):
    """Tipos de venta que no son campañas de anuncios: evaluaciones, retoques,
    recurrentes, recomendados, orgánico/redes, sesiones y filas con la
    columna CAMPAÑA vacía o con un dato inválido. No tienen presupuesto ni
    leads de Meta Ads porque no vienen de un anuncio pagado."""
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Otros (no son campanas de anuncios)', ha='center', fontsize=15,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)
    ax.text(0.5, 0.895, 'Evaluaciones, retoques, recurrentes, recomendados, organico/redes, '
                        'sesiones y datos sin campana valida.', ha='center',
            fontsize=8.5, color=GRIS, transform=fig.transFigure)

    cats = defaultdict(lambda: dict(ag=0, as_=0, co=0, mon=0.0))
    for (cat, crm), d in agg_otros.items():
        for k in ('ag', 'as_', 'co', 'mon'):
            cats[cat][k] += d[k]

    if not cats:
        ax.text(0.5, 0.6, 'No hubo ventas sin campaña en este periodo: todo lo\n'
                          'registrado corresponde a una campaña de anuncios real.',
                ha='center', fontsize=10, color=GRIS, transform=fig.transFigure)
        pdf.savefig(fig); plt.close(fig)
        return

    cat_list = sorted(cats.items(), key=lambda kv: -kv[1]['ag'])
    filas = [[c, d['ag'], d['as_'], f"{pct(d['as_'], d['ag']):.0f}%", d['co'],
              f"{pct(d['co'], d['as_']):.0f}%", monto(d['mon'])] for c, d in cat_list]
    tot = dict(ag=sum(d['ag'] for d in cats.values()), as_=sum(d['as_'] for d in cats.values()),
               co=sum(d['co'] for d in cats.values()), mon=sum(d['mon'] for d in cats.values()))
    filas.append(['TOTAL', tot['ag'], tot['as_'], f"{pct(tot['as_'], tot['ag']):.0f}%",
                  tot['co'], f"{pct(tot['co'], tot['as_']):.0f}%", monto(tot['mon'])])

    alto = min(0.55, max(0.12, 0.043 * (len(filas) + 1)))
    axc = fig.add_axes([0.08, 0.86 - alto, 0.84, alto])
    estilo_tabla(axc, filas, [0.28, 0.11, 0.11, 0.11, 0.11, 0.11, 0.17],
                 header=['Categoria', 'Agend.', 'Asist.', '% Asist.', 'Compr.',
                         '% Conv.', 'Monto'],
                 fontsize=9, scale=1.0)
    pdf.savefig(fig); plt.close(fig)


def pagina_campanas_meta(pdf, filas):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Rendimiento por campana: Meta Ads + embudo', ha='center',
            fontsize=15, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.928, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO} | '
                        'gasto y leads de la carga de Meta Ads', ha='center',
            fontsize=10, color=GRIS)

    total = dict(gasto=0, leads=0, ag=0, as_=0, co=0, mon=0.0)
    datos = []
    for f in filas:
        fila = [f['campania'][:18],
                monto(f['gasto']) if f['gasto'] else '—',
                str(f['leads']) if f['leads'] else '—',
                f"S/ {f['costo_res']:.2f}" if f['leads'] else '—',
                str(f['ag']),
                f"{f['pct_ag']:.0f}%" if f['leads'] else '—',
                str(f['as_']),
                f"{f['pct_as']:.0f}%",
                str(f['co']),
                monto(f['mon']),
                monto(f['ticket']) if f['as_'] else '—']
        datos.append(fila)
        if not f['organica']:
            total['gasto'] += f['gasto']; total['leads'] += f['leads']
        total['ag'] += f['ag']; total['as_'] += f['as_']
        total['co'] += f['co']; total['mon'] += f['mon']
    datos.append(['TOTAL',
                  monto(total['gasto']), str(total['leads']),
                  f"S/ {total['gasto'] / total['leads']:.2f}" if total['leads'] else '—',
                  str(total['ag']),
                  f"{pct(total['ag'], total['leads']):.0f}%" if total['leads'] else '—',
                  str(total['as_']), f"{pct(total['as_'], total['ag']):.0f}%",
                  str(total['co']), monto(total['mon']),
                  monto(total['mon'] / total['as_']) if total['as_'] else '—'])

    axc = fig.add_axes([0.03, 0.10, 0.94, 0.72])
    estilo_tabla(axc, datos,
                 [0.21, 0.07, 0.045, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07, 0.085, 0.095],
                 header=['Campana', 'Gasto\nMeta', 'Leads', 'Costo\n/lead', 'Agend.',
                         'Agend.\n/lead', 'Asist.', '%\nAsist.', 'Compr.',
                         'Monto', 'Ticket\n/asist.'],
                 fontsize=7.3, scale=3.6)
    ax.text(0.045, 0.055, 'Agend/lead = % de agendados logrados por cada lead de Meta.  '
                          'Ticket/asist. = ventas entre los que asistieron.',
            fontsize=7.5, color=GRIS, transform=fig.transFigure)
    ax.text(0.045, 0.035, 'El gasto y los leads de Meta son del rango del reporte cargado.',
            fontsize=7.5, color=GRIS, transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_campana_canal(pdf, agg):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Campanas por canal', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO, transform=fig.transFigure)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS, transform=fig.transFigure)

    por = defaultdict(dict)
    camps = defaultdict(lambda: dict(ag=0, as_=0, co=0, mon=0.0))
    for (camp, crm), d in agg.items():
        if crm not in CANALES:
            continue
        por[camp][crm] = d
        for k in ('ag', 'as_', 'co', 'mon'):
            camps[camp][k] += d[k]
    camp_list = sorted(camps.items(), key=lambda kv: -kv[1]['ag'])
    if not camp_list:
        ax.text(0.5, 0.6, 'No hubo campañas de anuncios en este periodo.',
                ha='center', fontsize=10, color=GRIS, transform=fig.transFigure)
        pdf.savefig(fig); plt.close(fig)
        return
    top = camp_list[:8]
    cero = dict(ag=0, as_=0, co=0, mon=0.0)

    def apiladas(axg, metric, titulo, fmt):
        ypos = np.arange(len(top))[::-1]
        left = np.zeros(len(top))
        for crm in CANALES:
            vals = np.array([por[c].get(crm, cero)[metric] for c, _ in top])
            axg.barh(ypos, vals, left=left, height=0.62,
                     color=COLOR_CANAL[crm], label=crm)
            left += vals
        axg.set_yticks(ypos)
        axg.set_yticklabels([c[:15] for c, _ in top], fontsize=6.5)
        axg.set_title(titulo, fontsize=11, fontweight='bold',
                      color=FONDO_TITULO, pad=8)
        axg.legend(loc='lower right', fontsize=6, ncol=3, frameon=True,
                   facecolor='white', edgecolor='none')
        for yi, total in zip(ypos, left):
            if total:
                axg.text(total, yi, fmt(total), va='center', ha='left',
                         fontsize=6.5, color=GRIS)
        axg.set_xlim(0, max(left) * 1.16)
        axg.set_xticks([])
        axg.spines[['top', 'right']].set_visible(False)

    ax1 = fig.add_axes([0.16, 0.50, 0.34, 0.38])
    apiladas(ax1, 'ag', 'Agendados por campana', lambda v: f'{int(v):,}')
    ax2 = fig.add_axes([0.62, 0.50, 0.28, 0.38])
    apiladas(ax2, 'mon', 'Monto vendido (S/)', lambda v: f'{v:,.0f}')

    ax.text(0.08, 0.48, 'Detalle por canal', fontsize=12, fontweight='bold',
            color=FONDO_TITULO, transform=fig.transFigure)

    filas = []
    for camp, d in top[:7]:
        row = [camp]
        for crm in CANALES:
            v = por[camp].get(crm, cero)
            row += [str(v['ag']), str(v['as_']), str(v['co']), f"{v['mon']:,.0f}"]
        filas.append(row)
    resto = camp_list[7:]

    def suma_por_canal(lista, k):
        return [sum(por[c].get(crm, cero)[k] for c, _ in lista) for crm in CANALES]

    if resto:
        ot = [suma_por_canal(resto, k) for k in ('ag', 'as_', 'co', 'mon')]
        fila = ['Otras campanas']
        for i in range(len(CANALES)):
            fila += [str(ot[0][i]), str(ot[1][i]), str(ot[2][i]), f"{ot[3][i]:,.0f}"]
        filas.append(fila)
    tc = [suma_por_canal(camp_list, k) for k in ('ag', 'as_', 'co', 'mon')]
    fila = ['TOTAL']
    for i in range(len(CANALES)):
        fila += [str(tc[0][i]), str(tc[1][i]), str(tc[2][i]), f"{tc[3][i]:,.0f}"]
    filas.append(fila)

    bx, by, bw, bh = 0.08, 0.06, 0.84, 0.40
    nrows = 2 + len(filas)
    row_h = bh / nrows
    widths = [0.19] + [0.0675] * (len(CANALES) * 4)
    xs = [bx]
    for w in widths:
        xs.append(xs[-1] + w * bw)

    def celda(x, y, w, h, texto, fc='white', tc='#333333', fs=7.5, fw='normal',
              ha='center', xp=0.0):
        ax.add_patch(plt.Rectangle((x, y), w, h, facecolor=fc,
                                   edgecolor='#bfbfbf', linewidth=0.4,
                                   transform=fig.transFigure))
        tx = x + w / 2 if ha == 'center' else x + xp
        ax.text(tx, y + h / 2, texto, ha=ha, va='center', fontsize=fs,
                color=tc, fontweight=fw, transform=fig.transFigure)

    ytop = by + bh
    celda(bx, ytop - row_h, widths[0] * bw, row_h, 'Campana',
          FONDO_TITULO, 'white', 8, 'bold')
    for i, crm in enumerate(CANALES):
        x0 = xs[1 + i * 4]
        x1 = xs[1 + i * 4 + 4]
        celda(x0, ytop - row_h, x1 - x0, row_h, crm,
              FONDO_TITULO, 'white', 8, 'bold')
    etiquetas = ['Ag', 'As', 'Co', 'S/'] * len(CANALES)
    for j, lab in enumerate(etiquetas):
        celda(xs[1 + j], ytop - 2 * row_h, widths[1 + j] * bw, row_h, lab,
              CELESTE, FONDO_TITULO, 7, 'bold')
    celda(bx, ytop - 2 * row_h, widths[0] * bw, row_h, '', CELESTE)
    for i, fila in enumerate(filas):
        y = ytop - (i + 3) * row_h
        es_total = fila[0] == 'TOTAL'
        fc = FONDO_TITULO if es_total else ('white' if i % 2 == 0 else '#f2f2f2')
        tc = 'white' if es_total else '#333333'
        fw = 'bold' if es_total else 'normal'
        celda(bx, y, widths[0] * bw, row_h, fila[0], fc, tc, 7, fw,
              'left', 0.008)
        for j in range(len(CANALES) * 4):
            celda(xs[1 + j], y, widths[1 + j] * bw, row_h, fila[1 + j],
                  fc, tc, 7, fw)
    pdf.savefig(fig); plt.close(fig)


def pagina_metricas(pdf, analitica):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Metricas del periodo', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    trat = analitica['trat'].most_common(8)
    if trat:
        ax1 = fig.add_axes([0.30, 0.52, 0.62, 0.30])
        names = [str(t[0])[:20] for t in trat]
        vals = [t[1] for t in trat]
        ypos = np.arange(len(names))[::-1]
        ax1.barh(ypos, vals, color=AZUL)
        ax1.set_yticks(ypos); ax1.set_yticklabels(names, fontsize=7.5)
        ax1.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        ax1.set_title('Tratamientos mas frecuentes', fontsize=12, fontweight='bold',
                      color=FONDO_TITULO, pad=10)
        for yi, va in zip(ypos, vals):
            ax1.text(va + 0.05, yi, str(va), va='center', fontsize=9, color=GRIS)
        ax1.tick_params(axis='x', labelsize=8)
        ax1.spines[['top', 'right']].set_visible(False)

    dist = analitica['dist'].most_common(8)
    if dist:
        ax2 = fig.add_axes([0.24, 0.06, 0.24, 0.38])
        names2 = [str(d[0])[:13] for d in dist]
        vals2 = [d[1] for d in dist]
        ypos2 = np.arange(len(names2))[::-1]
        ax2.barh(ypos2, vals2, color=VERDE)
        ax2.set_yticks(ypos2); ax2.set_yticklabels(names2, fontsize=6.8)
        ax2.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        ax2.set_title('Distritos', fontsize=11, fontweight='bold', color=FONDO_TITULO, pad=8)
        ax2.tick_params(axis='x', labelsize=7)
        ax2.spines[['top', 'right']].set_visible(False)

    ax3 = fig.add_axes([0.54, 0.06, 0.36, 0.38]); ax3.axis('off')
    ax3.set_title('Perfil del paciente', fontsize=11, fontweight='bold',
                  color=FONDO_TITULO, pad=8)
    edades = analitica['edades']
    sexo = analitica['sexo']
    ed_med = round(sum(edades) / len(edades), 1) if edades else 'n/d'
    ed_min = min(edades) if edades else '-'
    ed_max = max(edades) if edades else '-'
    sf = sexo.get('F', 0)
    sm = sexo.get('M', 0)
    stot = sf + sm or 1
    ax3.text(0.02, 0.95, 'Edad media', fontsize=9, fontweight='bold', color=GRIS)
    ax3.text(0.02, 0.82, f'{ed_med} anos (rango {ed_min}-{ed_max})', fontsize=12,
             fontweight='bold', color=FONDO_TITULO)

    ax3.text(0.02, 0.68, 'Bandas de edad', fontsize=9, fontweight='bold', color=GRIS)
    bandas = {'18-25': 0, '26-35': 0, '36-45': 0, '46-55': 0, '56+': 0}
    for e in edades:
        if e < 26:
            bandas['18-25'] += 1
        elif e < 36:
            bandas['26-35'] += 1
        elif e < 46:
            bandas['36-45'] += 1
        elif e < 56:
            bandas['46-55'] += 1
        else:
            bandas['56+'] += 1
    ed_tot = len(edades) or 1
    colores_banda = [AZUL, VERDE, NARANJA, CELESTE, GRIS]
    x0, y0, alto = 0.02, 0.58, 0.055
    for (nom, cnt), col in zip(bandas.items(), colores_banda):
        w = 0.96 * cnt / ed_tot
        if w > 0:
            ax3.add_patch(plt.Rectangle((x0, y0), w, alto, facecolor=col,
                                        edgecolor='white', linewidth=0.6,
                                        transform=ax3.transAxes))
        x0 += w
    for i, ((nom, cnt), col) in enumerate(zip(bandas.items(), colores_banda)):
        ax3.text(0.02 + (i % 2) * 0.50, 0.46 - (i // 2) * 0.10,
                 f'{nom}: {pct(cnt, ed_tot):.0f}%', fontsize=7.8, color=col, fontweight='bold')

    ax3.text(0.02, 0.12, 'Sexo', fontsize=9, fontweight='bold', color=GRIS)
    ax3.text(0.02, 0.00, f'F: {sf} ({pct(sf, stot):.0f}%)   |   M: {sm} ({pct(sm, stot):.0f}%)',
             fontsize=10, fontweight='bold', color=FONDO_TITULO)
    pdf.savefig(fig); plt.close(fig)


def pagina_ejecutivas(pdf, ejecutivas):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Performance por ejecutiva (Agendado por)', ha='center',
            fontsize=16, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    if not ejecutivas:
        ax.text(0.5, 0.5, 'No hay datos de ejecutivas en el periodo',
                ha='center', fontsize=12, color=GRIS)
        pdf.savefig(fig); plt.close(fig)
        return

    filas = []
    tot_ag_e = sum(e['ag'] for e in ejecutivas)
    tot_as_e = sum(e['as_'] for e in ejecutivas)
    tot_co_e = sum(e['co'] for e in ejecutivas)
    tot_mon_e = sum(e['mon'] for e in ejecutivas)
    tot_nf_e = sum(e['no_fueron'] for e in ejecutivas)
    for e in ejecutivas:
        filas.append([
            e['ejecutiva'][:18],
            e['ag'],
            e['as_'],
            f"{e['asistencia_pct']:.0f}%",
            e['co'],
            f"{e['conversion_pct']:.0f}%",
            e['no_fueron'],
            monto(e['mon']),
            monto(e['ticket']) if e['co'] else '—',
        ])
    filas.append([
        'TOTAL',
        tot_ag_e,
        tot_as_e,
        f"{pct(tot_as_e, tot_ag_e):.0f}%",
        tot_co_e,
        f"{pct(tot_co_e, tot_as_e):.0f}%",
        tot_nf_e,
        monto(tot_mon_e),
        monto(tot_mon_e / tot_co_e) if tot_co_e else '—',
    ])

    header = ['Ejecutiva', 'Agend.', 'Asist.', '%\nAsist.', 'Compr.',
              '%\nConv.', 'No\nAsist.', 'Monto', 'Ticket']
    col_w = [0.20, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.13, 0.13]
    sw = sum(col_w)
    col_w = [w / sw for w in col_w]

    ax_t = fig.add_axes([0.06, 0.50, 0.88, 0.35])
    estilo_tabla(ax_t, filas, col_w, header=header, fontsize=7.8, scale=2.2)

    valid_ej = [e for e in ejecutivas if e['ejecutiva'] != 'SIN EJECUTIVA'] or ejecutivas
    if valid_ej:
        ax_g = fig.add_axes([0.10, 0.10, 0.80, 0.32])
        names = [e['ejecutiva'][:16] for e in valid_ej][:8]
        montos_ej = [e['mon'] for e in valid_ej][:8]
        ypos = np.arange(len(names))[::-1]
        ax_g.barh(ypos, montos_ej, color=AZUL, height=0.55)
        ax_g.set_yticks(ypos)
        ax_g.set_yticklabels(names, fontsize=8.5)
        ax_g.set_title('Ventas generadas por ejecutiva (S/)', fontsize=11,
                       fontweight='bold', color=FONDO_TITULO, pad=8)
        max_m = max(montos_ej) if montos_ej else 100
        for yi, va in zip(ypos, montos_ej):
            ax_g.text(va + max(max_m * 0.02, 1), yi, monto(va),
                      va='center', fontsize=8.5, color=FONDO_TITULO, fontweight='bold')
        ax_g.tick_params(axis='x', labelsize=8)
        ax_g.spines[['top', 'right']].set_visible(False)
        ax_g.set_xlim(0, max_m * 1.25 if max_m > 0 else 100)

    pdf.savefig(fig); plt.close(fig)


def pagina_motivos(pdf, motivos):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Motivos de pérdida', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    m_as = motivos.get('no_asistio', {})
    m_co = motivos.get('no_compra', {})

    # Bloque 1: No Asistencia (arriba)
    ax.text(0.08, 0.88, '1. Motivos de NO Asistencia (agendaron pero no acudieron)',
            fontsize=12, fontweight='bold', color=FONDO_TITULO)
    tot_as_loss = sum(m_as.values())
    if m_as:
        filas_as = []
        for mot, cnt in m_as.items():
            filas_as.append([mot, cnt, f"{cnt / tot_as_loss * 100.0:.1f}%"])
        filas_as.append(['TOTAL', tot_as_loss, '100.0%'])
        ax_t1 = fig.add_axes([0.08, 0.58, 0.42, 0.26])
        estilo_tabla(ax_t1, filas_as, [0.55, 0.22, 0.23],
                     header=['Motivo', 'Cant.', '%'], fontsize=8, scale=1.3)

        ax_g1 = fig.add_axes([0.56, 0.58, 0.36, 0.26])
        names = list(m_as.keys())[:6]
        vals = list(m_as.values())[:6]
        ypos = np.arange(len(names))[::-1]
        ax_g1.barh(ypos, vals, color=ROJO, alpha=0.85, height=0.55)
        ax_g1.set_yticks(ypos)
        ax_g1.set_yticklabels(names, fontsize=7.5)
        ax_g1.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        ax_g1.spines[['top', 'right']].set_visible(False)
        for yi, va in zip(ypos, vals):
            ax_g1.text(va + 0.05, yi, str(va), va='center', fontsize=8, color=GRIS)
    else:
        ax.add_patch(plt.Rectangle((0.08, 0.60), 0.84, 0.24, facecolor=FONDO_CLARO,
                                   edgecolor=CELESTE, linewidth=1, transform=fig.transFigure))
        ax.text(0.5, 0.73, 'Sin motivos de no asistencia registrados en el periodo.',
                ha='center', va='center', fontsize=10, color=GRIS, transform=fig.transFigure)
        ax.text(0.5, 0.68, 'Tipificación en la columna "MOTIVO NO ASISTIO" (col AB) del maestro BD DATA.',
                ha='center', va='center', fontsize=8.5, color=AZUL, transform=fig.transFigure)

    # Bloque 2: No Compra (abajo)
    ax.text(0.08, 0.48, '2. Motivos de NO Compra (asistieron pero no compraron)',
            fontsize=12, fontweight='bold', color=FONDO_TITULO)
    tot_co_loss = sum(m_co.values())
    if m_co:
        filas_co = []
        for mot, cnt in m_co.items():
            filas_co.append([mot, cnt, f"{cnt / tot_co_loss * 100.0:.1f}%"])
        filas_co.append(['TOTAL', tot_co_loss, '100.0%'])
        ax_t2 = fig.add_axes([0.08, 0.18, 0.42, 0.26])
        estilo_tabla(ax_t2, filas_co, [0.55, 0.22, 0.23],
                     header=['Motivo', 'Cant.', '%'], fontsize=8, scale=1.3)

        ax_g2 = fig.add_axes([0.56, 0.18, 0.36, 0.26])
        names2 = list(m_co.keys())[:6]
        vals2 = list(m_co.values())[:6]
        ypos2 = np.arange(len(names2))[::-1]
        ax_g2.barh(ypos2, vals2, color=NARANJA, alpha=0.85, height=0.55)
        ax_g2.set_yticks(ypos2)
        ax_g2.set_yticklabels(names2, fontsize=7.5)
        ax_g2.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        ax_g2.spines[['top', 'right']].set_visible(False)
        for yi, va in zip(ypos2, vals2):
            ax_g2.text(va + 0.05, yi, str(va), va='center', fontsize=8, color=GRIS)
    else:
        ax.add_patch(plt.Rectangle((0.08, 0.20), 0.84, 0.24, facecolor=FONDO_CLARO,
                                   edgecolor=CELESTE, linewidth=1, transform=fig.transFigure))
        ax.text(0.5, 0.33, 'Sin motivos de no compra registrados en el periodo.',
                ha='center', va='center', fontsize=10, color=GRIS, transform=fig.transFigure)
        ax.text(0.5, 0.28, 'Tipificación en la columna "MOTIVO NO COMPRA" (col AC) del maestro BD DATA.',
                ha='center', va='center', fontsize=8.5, color=AZUL, transform=fig.transFigure)

    ax.text(0.08, 0.06, 'Fuente: maestro BD DATA.xlsx (columnas AB y AC).',
            fontsize=8, color=GRIS, transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_comparativo_historico(pdf, comp, historico):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Comparativo y tendencia', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO} vs periodos comparables',
            ha='center', fontsize=11, color=GRIS)

    actual = comp['actual']
    ant = comp['anterior_mes']
    anio_ant = comp['mismo_mes_anio_anterior']

    ax1 = fig.add_axes([0.12, 0.62, 0.76, 0.27])
    metricas = ['agendados', 'asistidos', 'compraron']
    labels_m = ['Agendados', 'Asistieron', 'Compraron']
    periodos = [('Periodo actual', actual, FONDO_TITULO),
                ('Mes anterior\n(mismos dias)', ant, CELESTE),
                ('Mismo periodo\nano anterior', anio_ant, GRIS)]
    x = np.arange(len(metricas))
    width = 0.25
    for i, (nombre, d, color) in enumerate(periodos):
        vals = [d.get(m, 0) for m in metricas]
        ax1.bar(x + (i - 1) * width, vals, width, label=nombre, color=color)
    ax1.set_xticks(x); ax1.set_xticklabels(labels_m, fontsize=9)
    ax1.legend(fontsize=6.5, loc='upper right', ncol=1)
    ax1.set_title('Embudo: periodo actual vs comparables', fontsize=11,
                  fontweight='bold', color=FONDO_TITULO, pad=10)
    ax1.tick_params(axis='y', labelsize=8)
    ax1.spines[['top', 'right']].set_visible(False)

    def var(a, b):
        return f"{((a - b) / b * 100):+.1f}%" if b else 'n/d'

    filas = [
        ['Agendados', actual['agendados'], ant.get('agendados', 0),
         anio_ant.get('agendados', 0), var(actual['agendados'], ant.get('agendados', 0))],
        ['Asistieron', actual['asistidos'], ant.get('asistidos', 0),
         anio_ant.get('asistidos', 0), var(actual['asistidos'], ant.get('asistidos', 0))],
        ['Compraron', actual['compraron'], ant.get('compraron', 0),
         anio_ant.get('compraron', 0), var(actual['compraron'], ant.get('compraron', 0))],
        ['Monto (S/)', f"{actual['monto']:,.0f}", f"{ant.get('monto', 0):,.0f}",
         f"{anio_ant.get('monto', 0):,.0f}", var(actual['monto'], ant.get('monto', 0))],
    ]
    axt = fig.add_axes([0.09, 0.42, 0.82, 0.15])
    estilo_tabla(axt, filas, [0.24, 0.19, 0.21, 0.21, 0.15],
                 header=['Metrica', 'Actual', 'Mes ant.', 'Ano ant.', 'Var. vs mes ant.'],
                 fontsize=8, scale=1.25)

    if historico:
        ax2 = fig.add_axes([0.10, 0.08, 0.82, 0.27])
        etiquetas = [f"{h['mes']} {str(h['anio'])[2:]}" for h in historico]
        montos_h = [h['monto'] for h in historico]
        ax2.plot(range(len(etiquetas)), montos_h, marker='o', color=AZUL,
                 linewidth=2, markersize=4)
        ax2.fill_between(range(len(etiquetas)), montos_h, alpha=0.12, color=AZUL)
        ax2.set_xticks(range(len(etiquetas)))
        ax2.set_xticklabels(etiquetas, fontsize=6.5, rotation=45, ha='right')
        ax2.set_title(f'Tendencia de ventas ({len(historico)} meses)', fontsize=11,
                      fontweight='bold', color=FONDO_TITULO, pad=8)
        ax2.spines[['top', 'right']].set_visible(False)
        ax2.tick_params(axis='y', labelsize=7)
    pdf.savefig(fig); plt.close(fig)


def pagina_evolucion_diaria(pdf, serie):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Evolucion diaria del periodo', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    dias = serie['dias']
    montos_d = serie['monto']
    if not dias or (not any(montos_d) and not any(serie['agendados'])):
        ax.text(0.5, 0.5, 'Sin actividad diaria registrada en el periodo.',
                ha='center', fontsize=11, color=GRIS, transform=fig.transFigure)
        pdf.savefig(fig); plt.close(fig)
        return

    ax1 = fig.add_axes([0.10, 0.56, 0.82, 0.32])
    ax1.bar(dias, montos_d, color=FONDO_TITULO, width=0.6)
    ax1.set_title('Monto vendido por dia (S/)', fontsize=12, fontweight='bold',
                  color=FONDO_TITULO, pad=8)
    ax1.set_xticks(dias)
    ax1.tick_params(axis='both', labelsize=7.5)
    ax1.spines[['top', 'right']].set_visible(False)
    for d, v in zip(dias, montos_d):
        if v:
            ax1.text(d, v, f'{v:,.0f}', ha='center', va='bottom', fontsize=6.5, color=GRIS)

    ax2 = fig.add_axes([0.10, 0.16, 0.82, 0.32])
    ax2.plot(dias, serie['agendados'], marker='o', color=AZUL, label='Agendados',
             linewidth=1.6, markersize=4)
    ax2.plot(dias, serie['asistidos'], marker='o', color=VERDE, label='Asistieron',
             linewidth=1.6, markersize=4)
    ax2.plot(dias, serie['compraron'], marker='o', color=NARANJA, label='Compraron',
             linewidth=1.6, markersize=4)
    ax2.set_xticks(dias)
    ax2.tick_params(axis='both', labelsize=7.5)
    ax2.legend(fontsize=8, loc='upper right')
    ax2.set_title('Agendados, asistencia y compras por dia', fontsize=12,
                  fontweight='bold', color=FONDO_TITULO, pad=8)
    ax2.spines[['top', 'right']].set_visible(False)

    if any(montos_d):
        mejor_i = max(range(len(dias)), key=lambda i: montos_d[i])
        ax.text(0.5, 0.06, f'Mejor dia del periodo: {dias[mejor_i]} de {NOMBRES_MES[MES]} '
                           f'con S/ {montos_d[mejor_i]:,.0f} en ventas.',
                ha='center', fontsize=9.5, color=AZUL, fontweight='bold',
                transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_recurrentes(pdf, rec):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Pacientes recurrentes y valor de vida (LTV)', ha='center',
            fontsize=15, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, 'Historico completo de la base de datos (no limitado al periodo)',
            ha='center', fontsize=10, color=GRIS)

    total_p = rec['total_pacientes']
    total_r = rec['total_recurrentes']
    kpis = [
        ('PACIENTES UNICOS', str(total_p), 'con al menos 1 compra', AZUL),
        ('RECURRENTES', str(total_r), f'{pct(total_r, total_p):.1f}% del total', VERDE),
        ('LTV PROMEDIO', monto(rec['ltv_promedio']), 'valor de vida por paciente', FONDO_TITULO),
    ]
    n = len(kpis)
    w = 0.26
    gap = (0.92 - n * w) / (n + 1)
    for i, (t, val, s, c) in enumerate(kpis):
        x = gap + i * (w + gap)
        caja_kpi(fig, x, 0.74, w, 0.15, t, val, s, c)

    ax.text(0.08, 0.65, 'Top pacientes recurrentes (por numero de compras)', fontsize=12,
            fontweight='bold', color=FONDO_TITULO)
    pacientes = rec['pacientes'][:15]
    if pacientes:
        filas = [[p['nombre'][:26], p['compras'], monto(p['monto']),
                  monto(p['monto'] / p['compras']), p['ultima']] for p in pacientes]
        axt = fig.add_axes([0.06, 0.10, 0.90, 0.52])
        estilo_tabla(axt, filas, [0.30, 0.13, 0.19, 0.19, 0.19],
                     header=['Paciente', 'Compras', 'Monto total', 'Ticket prom.', 'Ultima visita'],
                     fontsize=8, scale=1.3)
    else:
        ax.text(0.5, 0.4, 'Aun no hay pacientes con 2 o mas compras registradas.',
                ha='center', fontsize=11, color=GRIS, transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_reactivacion(pdf, react, meses_umbral=3):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Oportunidad de reactivacion de pacientes', ha='center',
            fontsize=15, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'Pacientes que asistieron alguna vez pero no vuelven hace mas '
                        f'de {meses_umbral} meses (a hoy)', ha='center',
            fontsize=10, color=GRIS)

    ax.add_patch(plt.Rectangle((0.08, 0.80), 0.84, 0.11, facecolor=FONDO_TITULO,
                               edgecolor='none', transform=fig.transFigure))
    ax.text(0.5, 0.855, f'{len(react)} pacientes para contactar', ha='center', va='center',
            fontsize=18, fontweight='bold', color='white', transform=fig.transFigure)

    if react:
        filas = [[p['nombre'][:26], p['telefono'], p['ultima_cita'], p['meses_sin_volver']]
                 for p in react[:25]]
        axt = fig.add_axes([0.08, 0.20, 0.84, 0.55])
        estilo_tabla(axt, filas, [0.34, 0.24, 0.22, 0.20],
                     header=['Paciente', 'Telefono', 'Ultima cita', 'Meses sin volver'],
                     fontsize=8, scale=1.3)
        if len(react) > 25:
            ax.text(0.5, 0.145, f'Mostrando los 25 casos mas antiguos de {len(react)} totales.',
                    ha='center', fontsize=8.5, color=GRIS, transform=fig.transFigure)
    else:
        ax.text(0.5, 0.5, 'No hay pacientes pendientes de reactivar en este momento.',
                ha='center', fontsize=11, color=GRIS, transform=fig.transFigure)
    ax.text(0.5, 0.065, 'Contactar a esta lista es una oportunidad directa de venta:',
            ha='center', fontsize=9, color=AZUL, fontweight='bold', transform=fig.transFigure)
    ax.text(0.5, 0.045, 'son pacientes que ya confiaron en el consultorio.',
            ha='center', fontsize=9, color=AZUL, fontweight='bold', transform=fig.transFigure)
    pdf.savefig(fig); plt.close(fig)


def pagina_hallazgos(pdf, tot, analitica, rec=None, react=None, serie=None):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')
    ax.text(0.5, 0.955, 'Hallazgos y sugerencias', ha='center', fontsize=16,
            fontweight='bold', color=FONDO_TITULO)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO}', ha='center',
            fontsize=11, color=GRIS)

    hallazgos = []
    asis = pct(tot['as_'], tot['ag'])
    conv = pct(tot['co'], tot['as_'])
    hallazgos.append(f"La asistencia fue del {asis:.0f}% ({tot['as_']} de {tot['ag']} "
                     f"agendados) y la conversion a compra del {conv:.0f}%.")
    if tot['co']:
        ticket = tot['mon'] / tot['co']
        hallazgos.append(f"Ticket promedio de {monto(ticket)} por compra.")
    mejor_crm = max(((c, v) for c, v in agg_tot_crm.items() if v['co']),
                    key=lambda kv: kv[1]['co'], default=None)
    if mejor_crm:
        hallazgos.append(f"{mejor_crm[0]} genera la mayor parte de las ventas "
                         f"({mejor_crm[1]['co']} compras, S/ {mejor_crm[1]['mon']:,.0f}).")
    top_trat = analitica['trat'].most_common(1)
    if top_trat:
        t, n = top_trat[0]
        hallazgos.append(f"El servicio mas solicitado fue {t} ({n} veces).")
    if rec and rec.get('total_pacientes'):
        pct_rec = pct(rec['total_recurrentes'], rec['total_pacientes'])
        hallazgos.append(f"{pct_rec:.0f}% de los pacientes de la base son recurrentes "
                         f"(2+ compras), con un LTV promedio de {monto(rec['ltv_promedio'])}.")
    if react:
        hallazgos.append(f"{len(react)} pacientes no vuelven hace mas de 3 meses: "
                         'oportunidad directa de reactivacion.')
    if serie and any(serie.get('monto', [])):
        montos_s = serie['monto']
        mejor_i = max(range(len(montos_s)), key=lambda i: montos_s[i])
        hallazgos.append(f"El mejor dia del periodo fue el {serie['dias'][mejor_i]} "
                         f"con S/ {montos_s[mejor_i]:,.0f} en ventas.")
    if tot['no_fueron']:
        hallazgos.append(f"{tot['no_fueron']} pacientes agendaron pero no fueron a "
                         f"su cita ({pct(tot['no_fueron'], tot['ag']):.0f}% de los "
                         'agendados).')
    if tot['fueron_sin_compra']:
        hallazgos.append(f"{tot['fueron_sin_compra']} pacientes fueron a su cita pero "
                         f"no compraron ({pct(tot['fueron_sin_compra'], tot['as_']):.0f}"
                         '% de los asistentes).')
    top_camp = analitica['camp'].most_common(1)
    if top_camp:
        c, n = top_camp[0]
        hallazgos.append(f"La campana con mas asistentes fue {c} ({n}).")
    top_dist = analitica['dist'].most_common(1)
    if top_dist:
        d, n = top_dist[0]
        hallazgos.append(f"Mayor concentracion de pacientes: {d} ({n} asistentes).")
    hallazgos = hallazgos[:8]

    sugerencias = []
    if asis < 65:
        sugerencias.append('Reforzar confirmacion y recordatorios (dia previo) para '
                           'subir la asistencia por encima del 65%.')
    else:
        sugerencias.append('Mantener el protocolo de recordatorios; la asistencia '
                           'ya supera el 65%.')
    if conv < 50:
        sugerencias.append('Capacitar en upselling durante la consulta para elevar '
                           'la conversion a compra.')
    else:
        sugerencias.append('La conversion es solida; sube el ticket con paquetes y '
                           'promociones.')
    if react:
        sugerencias.append(f'Contactar a los {len(react)} pacientes inactivos con una '
                           'promocion de regreso puede generar ventas rapidas.')
    if top_trat:
        t, n = top_trat[0]
        sugerencias.append(f'Reserva mas stock y agenda para {t}, el mas demandado.')
    sugerencias = sugerencias[:4]

    def bloque(titulo, items, y0, color):
        ax.text(0.08, y0, titulo, fontsize=12, fontweight='bold', color=color)
        y = y0 - 0.045
        for it in items:
            ax.text(0.08, y, '\u2022  ' + it, fontsize=9.5, color=GRIS, va='top',
                    wrap=True)
            y -= 0.056
        return y

    y_fin = bloque('Hallazgos', hallazgos, 0.89, FONDO_TITULO)
    bloque('Sugerencias', sugerencias, y_fin - 0.03, VERDE)
    pdf.savefig(fig); plt.close(fig)


agg_tot_crm = {}
tot_ag = tot_as = tot_co = 0
tot_mon = 0.0


def generar_reporte(mes='AGO', anio=2026, desde=1, hasta=10, fuente='maestro',
                    salida=None):
    """Genera el PDF del reporte y devuelve dict con resumen (reutilizable CLI/web)."""
    global MES, ANIO, D1, D2, FUENTE, SALIDA, tot_ag, tot_as, tot_co, tot_mon
    global CRM_ORDEN, CANALES, COLOR_CANAL
    MES, ANIO, D1, D2, FUENTE = mes, int(anio), int(desde), int(hasta), fuente
    if salida:
        SALIDA = salida if os.path.isabs(salida) else os.path.join(BASE_DIR, salida)
    agg_tot_crm.clear()
    agg = build_data()
    if COL.get('ES_BM'):
        CRM_ORDEN = ['KOMMO', 'WHATSAPP', 'ORGANICO', 'SIN CRM']
        CANALES = ['KOMMO', 'WHATSAPP', 'ORGANICO']
    else:
        CANALES = canales_desde_datos(agg)
        CRM_ORDEN = CANALES + ['SIN CRM']
    COLOR_CANAL = {c: PALETA[i % len(PALETA)] for i, c in enumerate(CANALES)}
    tot = dict(ag=0, as_=0, co=0, mon=0.0, no_fueron=0, fueron_sin_compra=0)
    for d in agg.values():
        tot['ag'] += d['ag']; tot['as_'] += d['as_']
        tot['co'] += d['co']; tot['mon'] += d['mon']
        tot['no_fueron'] += d['no_fueron']
        tot['fueron_sin_compra'] += d['fueron_sin_compra']
    for c in CRM_ORDEN:
        sel = {k: v for k, v in agg.items() if k[1] == c}
        if sel:
            agg_tot_crm[c] = dict(ag=sum(v['ag'] for v in sel.values()),
                                  as_=sum(v['as_'] for v in sel.values()),
                                  co=sum(v['co'] for v in sel.values()),
                                  mon=sum(v['mon'] for v in sel.values()),
                                  no_fueron=sum(v['no_fueron'] for v in sel.values()),
                                  fueron_sin_compra=sum(v['fueron_sin_compra'] for v in sel.values()))
    tot_ag, tot_as, tot_co, tot_mon = tot['ag'], tot['as_'], tot['co'], tot['mon']
    agg_campanas, agg_otros = separar_campanas_otros(agg)
    analitica = datos_analiticos()
    filas_meta = build_campana_meta(agg_campanas)
    ejecutivas = datos_ejecutivas()
    motivos = datos_motivos()
    gasto_ads = _gasto_ads_total()

    # Variación vs periodo anterior equivalente (mismo rango de días)
    pm, pa, pd, ph = _periodo_anterior(MES, ANIO, D1, D2)
    tot_ant = totales_periodo(pm, pa, pd, ph)
    variacion = {
        'ag': _variacion(tot['ag'], tot_ant['ag']),
        'as_': _variacion(tot['as_'], tot_ant['as_']),
        'co': _variacion(tot['co'], tot_ant['co']),
        'mon': _variacion(tot['mon'], tot_ant['mon']),
    }

    # Analítica adicional (comparativos, tendencia, recurrencia, reactivación).
    # Import diferido para evitar el ciclo de imports con analitica.py, que a
    # su vez importa este módulo.
    import analitica as ana
    comp = ana.comparativo(MES, ANIO, D1, D2)
    historico = ana.ventas_por_mes()
    serie = ana.serie_diaria(MES, ANIO, D1, D2)
    rec = ana.recurrentes(MES, ANIO, D1, D2)
    react = ana.pacientes_a_reactivar(meses=3)

    with PdfPages(SALIDA) as pdf:
        pagina_resumen(pdf, tot, agg, variacion=variacion, gasto_ads=gasto_ads)
        pagina_comparativo_historico(pdf, comp, historico)
        pagina_flujo(pdf, tot)
        pagina_evolucion_diaria(pdf, serie)
        pagina_campanas(pdf, agg_campanas)
        if filas_meta:
            pagina_campanas_meta(pdf, filas_meta)
        pagina_campana_canal(pdf, agg_campanas)
        pagina_otros(pdf, agg_otros)
        if ejecutivas:
            pagina_ejecutivas(pdf, ejecutivas)
        pagina_metricas(pdf, analitica)
        pagina_recurrentes(pdf, rec)
        pagina_reactivacion(pdf, react)
        pagina_motivos(pdf, motivos)
        pagina_hallazgos(pdf, tot, analitica, rec=rec, react=react, serie=serie)
    return {'archivo': SALIDA, 'totales': tot,
            'por_crm': {c: dict(v) for c, v in agg_tot_crm.items()},
            'detalle': {f'{k[0]} | {k[1]}': dict(v) for k, v in sorted(
                agg.items(), key=lambda kv: -kv[1]['ag'])},
            'por_campana_meta': filas_meta,
            'ejecutivas': ejecutivas,
            'motivos': motivos,
            'variacion': variacion,
            'recurrentes': {'total_pacientes': rec['total_pacientes'],
                            'total_recurrentes': rec['total_recurrentes'],
                            'ltv_promedio': rec['ltv_promedio']},
            'reactivacion': {'pendientes': len(react)}}


def main(argv=None):
    ap = argparse.ArgumentParser(description='Reporte de ventas por campana de agendados (PDF)')
    ap.add_argument('--mes', default='AGO')
    ap.add_argument('--anio', type=int, default=2026)
    ap.add_argument('--desde', type=int, default=1)
    ap.add_argument('--hasta', type=int, default=10)
    ap.add_argument('--fuente', choices=['maestro', 'auto'], default='maestro',
                    help='maestro = pivotes guardados; auto = integra AGENDADOS+VENTA de Drive')
    ap.add_argument('--salida', default=None)
    a = ap.parse_args(argv)
    res = generar_reporte(mes=a.mes, anio=a.anio, desde=a.desde, hasta=a.hasta,
                          fuente=a.fuente, salida=a.salida)
    tot = res['totales']
    print(f"PDF generado: {res['archivo']}")
    print(f"Agendados: {tot['ag']} | Asistieron: {tot['as_']} | Compraron: {tot['co']} | "
          f"Monto: S/ {tot['mon']:,.0f}".replace(',', ' '))


if __name__ == '__main__':
    main()
