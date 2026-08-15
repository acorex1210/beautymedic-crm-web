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
        filas.append(dict(campania=camp + ' (sin Meta)', gasto=0, leads=0,
                          costo_res=0, ag=d['ag'], pct_ag=0,
                          as_=d['as_'], pct_as=pct(d['as_'], d['ag']),
                          co=d['co'], mon=d['mon'],
                          ticket=d['mon'] / d['as_'] if d['as_'] else 0,
                          organica=True))
    return filas


def estilo_tabla(ax, datos, col_w, header=None, color_titulo=FONDO_TITULO,
                 fontsize=8.5, scale=1.7, alinear_izq_col0=True, altura_fila=None):
    ax.axis('off')
    tabla = ax.table(cellText=datos, colLabels=header, cellLoc='center',
                     loc='center', colWidths=col_w)
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(fontsize)
    tabla.scale(1, scale)
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


def caja_kpi(fig, x, y, w, h, titulo, valor, sub, color):
    ax = fig.add_axes([x, y, w, h]); ax.axis('off')
    ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=color, edgecolor='none',
                               transform=ax.transAxes))
    fs = 24 if len(valor) <= 3 else 22
    ax.text(0.5, 0.72, valor, ha='center', va='center', fontsize=fs,
            fontweight='bold', color='white', transform=ax.transAxes)
    ax.text(0.5, 0.42, titulo, ha='center', va='center', fontsize=10,
            color='white', fontweight='bold', transform=ax.transAxes)
    ax.text(0.5, 0.12, sub, ha='center', va='center', fontsize=8,
            color='white', alpha=0.9, transform=ax.transAxes)


def pagina_resumen(pdf, tot, agg):
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis('off')

    ax.text(0.5, 0.96, 'REPORTE DE VENTAS', ha='center', fontsize=24,
            fontweight='bold', color=FONDO_TITULO, transform=fig.transFigure)
    ax.text(0.5, 0.925, f'{D1} al {D2} de {NOMBRES_MES[MES]} {ANIO} - Campaña de agendados',
            ha='center', fontsize=12, color=GRIS, transform=fig.transFigure)
    ax.add_line(plt.Line2D([0.15, 0.85], [0.90, 0.90], transform=fig.transFigure,
                           color=CELESTE, lw=2))

    kpis = [
        ('AGENDADOS', str(tot['ag']), 'citas agendadas en el periodo', AZUL),
        ('ASISTIERON', str(tot['as_']), f"{pct(tot['as_'], tot['ag']):.0f}% de asistencia", VERDE),
        ('COMPRARON', str(tot['co']), f"{pct(tot['co'], tot['as_']):.0f}% de conversion", NARANJA),
        ('MONTO TOTAL', monto(tot['mon']), 'ventas del periodo', FONDO_TITULO),
    ]
    n = len(kpis)
    w = 0.21
    gap = (0.92 - n * w) / (n + 1)
    for i, (t, v, s, c) in enumerate(kpis):
        x = gap + i * (w + gap)
        caja_kpi(fig, x, 0.70, w, 0.16, t, v, s, c)

    ax.add_patch(plt.Rectangle((0.06, 0.48), 0.88, 0.15, facecolor=FONDO_CLARO,
                               edgecolor=CELESTE, linewidth=1, transform=fig.transFigure))
    if tot['co']:
        ticket = tot['mon'] / tot['co']
        ax.text(0.09, 0.585, 'Indicadores clave', fontsize=11, fontweight='bold',
                color=FONDO_TITULO, transform=fig.transFigure)
        lineas = [
            f"Ticket promedio: S/ {ticket:,.0f} por compra  |  Asistencia: {pct(tot['as_'], tot['ag']):.0f}%  |  Conversion a compra: {pct(tot['co'], tot['as_']):.0f}%",
            f"Agendaron pero no fueron: {tot['no_fueron']}  |  Fueron pero no compraron: {tot['fueron_sin_compra']}.",
            f"Cada {tot['ag']} agendados generan {tot['co']} ventas por S/ {tot['mon']:,.0f}.",
        ]
    else:
        lineas = ['No se registraron ventas en el periodo.',
                  f"Agendaron pero no fueron: {tot['no_fueron']}  |  Fueron pero no compraron: {tot['fueron_sin_compra']}."]
    ax.text(0.09, 0.545, lineas[0], fontsize=9.5, color=GRIS, transform=fig.transFigure)
    if len(lineas) > 1:
        ax.text(0.09, 0.512, lineas[1], fontsize=9.5, color=GRIS, transform=fig.transFigure)
    if len(lineas) > 2:
        ax.text(0.09, 0.479, lineas[2], fontsize=9.5, color=GRIS, transform=fig.transFigure)

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
    filas.append(['TOTAL', tot_ag, tot_as, tot_co,
                  f"{pct(tot_co, tot_as):.0f}%", monto(tot_mon)])
    ax.text(0.08, 0.90, 'Por campana (principales)', fontsize=12,
            fontweight='bold', color=FONDO_TITULO)
    axc = fig.add_axes([0.08, 0.42, 0.84, 0.46])
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
    crm_filas.append(['TOTAL', tot_ag, tot_as, tot_co,
                      f"{pct(tot_co, tot_as):.0f}%", monto(tot_mon)])
    ax.text(0.08, 0.38, 'Por CRM', fontsize=12, fontweight='bold', color=FONDO_TITULO)
    axd = fig.add_axes([0.08, 0.08, 0.84, 0.27])
    estilo_tabla(axd, crm_filas, [0.30, 0.15, 0.15, 0.13, 0.12, 0.18],
                 header=['CRM', 'Agend.', 'Asist.', 'Compr.', 'Conv.', 'Monto'],
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
        fila = [f['campania'],
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

    axc = fig.add_axes([0.045, 0.10, 0.91, 0.78])
    estilo_tabla(axc, datos,
                 [0.21, 0.075, 0.055, 0.08, 0.075, 0.07, 0.075, 0.07, 0.075, 0.09, 0.09],
                 header=['Campana', 'Gasto Meta', 'Leads', 'Costo/lead', 'Agendados',
                         'Agend/lead', 'Asistieron', 'Asistencia', 'Compraron',
                         'Monto', 'Ticket/asist.'],
                 fontsize=8, scale=6.0)
    ax.text(0.045, 0.045, 'Agend/lead = % de agendados logrados por cada lead de Meta. '
                          'Ticket/asist. = ventas entre los que asistieron. '
                          'El gasto y los leads de Meta son del rango del reporte cargado.',
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
        ax1 = fig.add_axes([0.10, 0.48, 0.80, 0.42])
        names = [str(t[0])[:22] for t in trat]
        vals = [t[1] for t in trat]
        ypos = np.arange(len(names))[::-1]
        ax1.barh(ypos, vals, color=AZUL)
        ax1.set_yticks(ypos); ax1.set_yticklabels(names, fontsize=8)
        ax1.set_title('Tratamientos mas frecuentes', fontsize=12, fontweight='bold',
                      color=FONDO_TITULO, pad=10)
        for yi, va in zip(ypos, vals):
            ax1.text(va + 0.05, yi, str(va), va='center', fontsize=9, color=GRIS)
        ax1.tick_params(axis='x', labelsize=8)
        ax1.spines[['top', 'right']].set_visible(False)

    dist = analitica['dist'].most_common(8)
    if dist:
        ax2 = fig.add_axes([0.10, 0.06, 0.38, 0.38])
        names2 = [str(d[0])[:14] for d in dist]
        vals2 = [d[1] for d in dist]
        ypos2 = np.arange(len(names2))[::-1]
        ax2.barh(ypos2, vals2, color=VERDE)
        ax2.set_yticks(ypos2); ax2.set_yticklabels(names2, fontsize=7)
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
    ax3.text(0.02, 0.88, 'Edad media', fontsize=9, fontweight='bold', color=GRIS)
    ax3.text(0.02, 0.72, f'{ed_med} anos', fontsize=16, fontweight='bold', color=FONDO_TITULO)
    ax3.text(0.02, 0.55, 'Rango de edad', fontsize=9, fontweight='bold', color=GRIS)
    ax3.text(0.02, 0.40, f'{ed_min} - {ed_max} anos', fontsize=12, color=FONDO_TITULO)
    ax3.text(0.02, 0.20, f'F: {sf}  ({pct(sf, stot):.0f}%)', fontsize=10, color=AZUL)
    ax3.text(0.02, 0.08, f'M: {sm}  ({pct(sm, stot):.0f}%)', fontsize=10, color=VERDE)
    pdf.savefig(fig); plt.close(fig)


def pagina_hallazgos(pdf, tot, analitica):
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
    top_trat = analitica['trat'].most_common(1)
    if top_trat:
        t, n = top_trat[0]
        hallazgos.append(f"El servicio mas solicitado fue {t} ({n} veces).")
    top_dist = analitica['dist'].most_common(1)
    if top_dist:
        d, n = top_dist[0]
        hallazgos.append(f"Mayor concentracion de pacientes: {d} ({n} asistentes).")
    mejor_crm = max(((c, v) for c, v in agg_tot_crm.items() if v['co']),
                    key=lambda kv: kv[1]['co'], default=None)
    if mejor_crm:
        hallazgos.append(f"{mejor_crm[0]} genera la mayor parte de las ventas "
                         f"({mejor_crm[1]['co']} compras, S/ {mejor_crm[1]['mon']:,.0f}).")
    if tot['co']:
        ticket = tot['mon'] / tot['co']
        hallazgos.append(f"Ticket promedio de {monto(ticket)} por compra.")

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
    if top_trat:
        t, n = top_trat[0]
        sugerencias.append(f'Reserva mas stock y agenda para {t}, el mas demandado.')
    sugerencias.append('Monitorea el embudo semanalmente para detectar caidas de '
                       'asistencia o conversion a tiempo.')

    def bloque(titulo, items, y0, color):
        ax.text(0.08, y0, titulo, fontsize=12, fontweight='bold', color=color)
        y = y0 - 0.05
        for it in items:
            ax.text(0.08, y, '\u2022  ' + it, fontsize=10, color=GRIS, va='top')
            y -= 0.065

    bloque('Hallazgos', hallazgos, 0.89, FONDO_TITULO)
    bloque('Sugerencias', sugerencias, 0.47, VERDE)
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
    analitica = datos_analiticos()
    filas_meta = build_campana_meta(agg)
    with PdfPages(SALIDA) as pdf:
        pagina_resumen(pdf, tot, agg)
        pagina_flujo(pdf, tot)
        pagina_campanas(pdf, agg)
        if filas_meta:
            pagina_campanas_meta(pdf, filas_meta)
        pagina_campana_canal(pdf, agg)
        pagina_metricas(pdf, analitica)
        pagina_hallazgos(pdf, tot, analitica)
    return {'archivo': SALIDA, 'totales': tot,
            'por_crm': {c: dict(v) for c, v in agg_tot_crm.items()},
            'detalle': {f'{k[0]} | {k[1]}': dict(v) for k, v in sorted(
                agg.items(), key=lambda kv: -kv[1]['ag'])},
            'por_campana_meta': filas_meta}


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
