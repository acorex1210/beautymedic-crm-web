#!/usr/bin/env bash
# Arranca la web de Derma Essenza en local + túnel público (Tailscale Funnel).
#
# Uso: bash iniciar_derma.sh
# Después de ejecutarlo, la web queda en:
#   - Local:  http://127.0.0.1:8011
#   - Pública: https://macbook-neo-de-andre.tailab4d2b.ts.net
set -u
cd "$(dirname "$0")"

echo "=== 1) Levantando servidor local (uvicorn) ==="
if curl -s -o /dev/null --max-time 3 http://127.0.0.1:8011/; then
  echo "Servidor ya está corriendo en el puerto 8011."
else
  bash run_derma_local.sh > /tmp/uvicorn_derma.log 2>&1 &
  echo "Servidor iniciado en segundo plano (log: /tmp/uvicorn_derma.log)."
fi
sleep 4
curl -s -o /dev/null --max-time 5 http://127.0.0.1:8011/ && echo "OK: servidor local respondiendo." || { echo "ERROR: el servidor no responde. Revisa /tmp/uvicorn_derma.log"; exit 1; }

echo
echo "=== 2) Verificando Tailscale ==="
if tailscale status >/dev/null 2>&1; then
  echo "Tailscale conectado."
else
  echo "Tailscale no está iniciado. Abre la app Tailscale y vuelve a ejecutar este script."
  exit 1
fi

echo
echo "=== 3) Publicando en internet (Funnel) ==="
if tailscale funnel status 2>/dev/null | grep -q "8011"; then
  echo "Funnel ya estaba activo."
else
  tailscale funnel --bg 8011 >/dev/null 2>&1 || { echo "No se pudo habilitar Funnel."; exit 1; }
  echo "Funnel habilitado."
fi

echo
echo "==================== LISTO ===================="
echo "Local:    http://127.0.0.1:8011"
echo "Pública:  https://macbook-neo-de-andre.tailab4d2b.ts.net"
echo "==============================================="
echo
echo "Para detener la web pública:"
echo "  tailscale funnel --https=443 off"
echo "  kill \$(pgrep -f 'uvicorn app:app')"
