#!/usr/bin/env bash
# Levanta la web de Derma Essenza en local con uvicorn (sin hosting de pago).
# Uso: bash run_derma_local.sh [puerto]   (default 8011)
set -euo pipefail
cd "$(dirname "$0")"

PORT="${1:-8011}"

export CREDENCIALES="${CREDENCIALES:-$HOME/credenciales-derma.json}"
export TMP_DIR="${TMP_DIR:-/tmp}"

# Maestro de ventas (BD DATA de Derma Essenza)
export MAESTRO_PATH="${MAESTRO_PATH:-$HOME/Downloads/BD DATA DERMA ESSENZA.xlsx}"

# Archivos de Google Drive de Derma Essenza
export AGENDADOS_FID="1So_1Fh744c3K9kss2oA1twjBLJpgrSxZCu2lqhWpqJM"
export VENTA_FID="1TDM7ZFV6Jdsqc6i4CadNkwPQNdrIBhu7"

# Marca de la web
export BRAND_NOMBRE="Derma Essenza"
export BRAND_BADGE="DE"
export CRM_ARCHIVO="CRM DERMA ESSENZA.xlsx"

echo "=== Derma Essenza local: http://127.0.0.1:$PORT ==="
exec python3 -m uvicorn app:app --host 0.0.0.0 --port "$PORT"
