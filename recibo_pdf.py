# -*- coding: utf-8 -*-
"""recibo_pdf.py
=================
Genera un recibo interno en PDF para una venta de VENTA DIARIA (una o varias
líneas de tratamiento de una misma compra). Es un comprobante de cortesía
para el paciente, NO una boleta o factura electrónica válida ante SUNAT
(eso requeriría un proveedor de servicios electrónicos (PSE) homologado,
como Nubefact o Bizlinks, con RUC y certificado digital propios).
"""
import os
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

FONDO_TITULO = '#1f3864'
VERDE = '#70ad47'
GRIS = '#595959'
BORDE = '#c9d3e0'

ANCHO_CM = 14.8       # A5 de ancho
ALTO_HEADER_CM = 2.6
ALTO_LINEA_DATO_CM = 0.72
ALTO_FILA_ITEM_CM = 0.85
ALTO_FOOTER_CM = 2.1
MARGEN_CM = 0.9


def numero_recibo(hoja, fila):
    anio = ''.join(c for c in str(hoja) if c.isdigit())[-2:]
    return f'R{anio}-{fila:06d}'


def generar_recibo(datos, ruta, brand_nombre='Derma Essenza'):
    """Genera el PDF del recibo en ``ruta``.

    ``datos`` = {
      'hoja': str, 'fila': int, 'fecha': 'DD/MM/AAAA', 'paciente': str,
      'telefono': str, 'doctor': str, 'pago': str,
      'lineas': [{'tratamiento': str, 'venta': float}, ...],
    }
    """
    lineas = [ln for ln in (datos.get('lineas') or []) if ln.get('tratamiento')]
    total = sum(ln.get('venta') or 0 for ln in lineas)
    numero = numero_recibo(datos.get('hoja') or '', datos.get('fila') or 0)

    datos_persona = [(et, val) for et, val in (
        ('Paciente', datos.get('paciente')), ('Teléfono', datos.get('telefono')),
        ('Doctor', datos.get('doctor')), ('Forma de pago', datos.get('pago')),
    ) if val]

    alto_cm = (ALTO_HEADER_CM + 0.5
              + len(datos_persona) * ALTO_LINEA_DATO_CM + 0.4
              + 0.7  # encabezado de tabla
              + max(1, len(lineas)) * ALTO_FILA_ITEM_CM + 0.3
              + 1.0  # fila de total
              + ALTO_FOOTER_CM + MARGEN_CM)
    alto_cm = max(alto_cm, 14.0)

    fig = plt.figure(figsize=(ANCHO_CM / 2.54, alto_cm / 2.54))
    fig.patch.set_facecolor('white')
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis('off')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    def yf(cm_desde_arriba):
        return 1 - cm_desde_arriba / alto_cm

    # Encabezado
    y0, y1 = yf(0), yf(ALTO_HEADER_CM)
    ax.add_patch(plt.Rectangle((0, y1), 1, y0 - y1, transform=ax.transAxes,
                               color=FONDO_TITULO, zorder=0))
    ax.text(0.06, yf(1.05), brand_nombre, transform=ax.transAxes, ha='left', va='center',
           fontsize=18, fontweight='bold', color='white')
    ax.text(0.06, yf(1.85), 'Recibo interno de venta', transform=ax.transAxes,
           ha='left', va='center', fontsize=10, color='white')
    ax.text(0.94, yf(1.05), numero, transform=ax.transAxes, ha='right', va='center',
           fontsize=13, fontweight='bold', color='white')
    ax.text(0.94, yf(1.85), datos.get('fecha') or '', transform=ax.transAxes,
           ha='right', va='center', fontsize=10, color='white')

    cm = ALTO_HEADER_CM + 0.5

    for etiqueta, valor in datos_persona:
        ax.text(0.06, yf(cm), etiqueta, transform=ax.transAxes, ha='left', va='center',
               fontsize=9, color=GRIS)
        ax.text(0.35, yf(cm), str(valor), transform=ax.transAxes, ha='left', va='center',
               fontsize=10.5, color='#111', fontweight='bold')
        cm += ALTO_LINEA_DATO_CM

    cm += 0.15
    ax.plot([0.06, 0.94], [yf(cm), yf(cm)], color=BORDE, linewidth=1, transform=ax.transAxes)
    cm += 0.45

    ax.text(0.06, yf(cm), 'TRATAMIENTO', transform=ax.transAxes, ha='left', va='center',
           fontsize=9, fontweight='bold', color=GRIS)
    ax.text(0.94, yf(cm), 'MONTO (S/)', transform=ax.transAxes, ha='right', va='center',
           fontsize=9, fontweight='bold', color=GRIS)
    cm += 0.35
    ax.plot([0.06, 0.94], [yf(cm), yf(cm)], color=BORDE, linewidth=1, transform=ax.transAxes)
    cm += 0.55

    for ln in lineas:
        ax.text(0.06, yf(cm), str(ln.get('tratamiento') or ''), transform=ax.transAxes,
               ha='left', va='center', fontsize=10.5, color='#111')
        monto = ln.get('venta')
        ax.text(0.94, yf(cm), (f"{monto:,.2f}" if monto else '—'), transform=ax.transAxes,
               ha='right', va='center', fontsize=10.5, color='#111')
        cm += ALTO_FILA_ITEM_CM

    cm += 0.1
    ax.plot([0.06, 0.94], [yf(cm), yf(cm)], color=BORDE, linewidth=1, transform=ax.transAxes)
    cm += 0.75

    ax.text(0.62, yf(cm), 'TOTAL', transform=ax.transAxes, ha='left', va='center',
           fontsize=12, fontweight='bold', color=FONDO_TITULO)
    ax.text(0.94, yf(cm), f'S/ {total:,.2f}', transform=ax.transAxes, ha='right', va='center',
           fontsize=14, fontweight='bold', color=VERDE)

    cm += 0.9
    nota = ('Este documento es un comprobante interno de cortesía y no constituye una\n'
           'boleta ni factura electrónica válida ante SUNAT.')
    ax.text(0.5, yf(cm), nota, transform=ax.transAxes, ha='center', va='center',
           fontsize=8, color=GRIS, style='italic', linespacing=1.5)
    cm += 0.85
    ax.text(0.5, yf(cm), f'Generado el {datetime.now().strftime("%d/%m/%Y %H:%M")}',
           transform=ax.transAxes, ha='center', va='center', fontsize=7, color='#999')

    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    fig.savefig(ruta, format='pdf')
    plt.close(fig)
    return ruta
