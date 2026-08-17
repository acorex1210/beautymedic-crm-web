# -*- coding: utf-8 -*-
"""whatsapp_leads.py
==================
Webhook de solo-lectura para WhatsApp Business Cloud API (Meta).

No envía mensajes ni inicia conversaciones: únicamente recibe los mensajes
entrantes (vía "coexistencia", sin afectar la app de WhatsApp Business que
ya usa la clínica para responder), detecta la campaña de origen y crea/
actualiza una tarjeta en el Pipeline.

Detección de campaña, en orden de prioridad:
  1. ``referral`` de Meta Ads (Click-to-WhatsApp) — viene exacto en el payload.
  2. Palabra clave del primer mensaje, según el mapa editable en
     ``whatsapp_campanas.json``.
  3. "SIN CAMPAÑA / ORGÁNICO" si no coincide nada.

Variables de entorno:
  WHATSAPP_VERIFY_TOKEN  token que se define al configurar el webhook en
                         Meta for Developers (cualquier string secreto).
  WHATSAPP_APP_SECRET    app secret de la app de Meta, para validar la firma
                         de cada envío (opcional mientras no esté configurado).
"""
import hashlib
import hmac
import json
import os
import re

import crm_plus as cp

DATA_DIR = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data'))
CAMPANAS_PATH = os.path.join(DATA_DIR, 'whatsapp_campanas.json')

VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN', '')
APP_SECRET = os.environ.get('WHATSAPP_APP_SECRET', '')

SIN_CAMPANA = 'SIN CAMPAÑA / ORGÁNICO'


def leer_campanas():
    """Mapa {palabra_clave: nombre_campaña}, editable desde la web."""
    if not os.path.exists(CAMPANAS_PATH):
        return {}
    try:
        with open(CAMPANAS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return {}


def guardar_campanas(mapa):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CAMPANAS_PATH, 'w', encoding='utf-8') as f:
        json.dump(mapa, f, ensure_ascii=False, indent=2)


def detectar_campana(texto, referral=None):
    if referral:
        nombre = referral.get('headline') or referral.get('body') or referral.get('source_id')
        if nombre:
            return f'META ADS: {nombre}'
    texto_l = (texto or '').upper()
    for palabra, campana in leer_campanas().items():
        if palabra and palabra.upper() in texto_l:
            return campana
    return SIN_CAMPANA


def verificar_firma(payload_bytes, signature_header):
    """Valida X-Hub-Signature-256. Si no hay APP_SECRET configurado aún, deja pasar."""
    if not APP_SECRET:
        return True
    if not signature_header or not signature_header.startswith('sha256='):
        return False
    esperado = hmac.new(APP_SECRET.encode(), payload_bytes, hashlib.sha256).hexdigest()
    recibido = signature_header.split('=', 1)[1]
    return hmac.compare_digest(esperado, recibido)


def _tarjeta_activa(telefono):
    for t in cp.leer_tarjetas():
        if t.get('telefono') == telefono and t.get('etapa') not in ('GANADO', 'PERDIDO'):
            return t
    return None


def procesar_mensaje(telefono, texto, referral=None):
    """Crea una tarjeta NUEVO en el pipeline, o actualiza la nota si ya hay una activa."""
    campana = detectar_campana(texto, referral)
    nota = f'WhatsApp: "{(texto or "")[:120]}"'
    existente = _tarjeta_activa(telefono)
    if existente:
        return cp.actualizar_tarjeta(existente['id'], {'nota': nota})
    return cp.crear_tarjeta({
        'telefono': telefono, 'etapa': 'NUEVO', 'crm': 'WHATSAPP',
        'campana': campana, 'nota': nota,
    })


def procesar_webhook_payload(payload):
    """Recorre el payload de la Cloud API y procesa cada mensaje entrante de texto."""
    resultados = []
    for entry in payload.get('entry', []):
        for change in entry.get('changes', []):
            valor = change.get('value', {})
            for msg in valor.get('messages', []):
                telefono = re.sub(r'\D', '', msg.get('from') or '')
                if not telefono:
                    continue
                texto = ''
                if msg.get('type') == 'text':
                    texto = msg.get('text', {}).get('body', '')
                referral = msg.get('referral')
                resultados.append(procesar_mensaje(telefono, texto, referral))
    return resultados
