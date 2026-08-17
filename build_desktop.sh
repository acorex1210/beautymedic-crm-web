#!/usr/bin/env bash
# Script para compilar la aplicación de escritorio de Derma Essenza.
# Funciona en macOS y Windows (con ajustes menores de rutas).
set -euo pipefail
cd "$(dirname "$0")"

echo "=== Compilando aplicación de escritorio ==="
# Determinar el separador de rutas para PyInstaller (":" en macOS/Linux, ";" en Windows)
SEP=":"
if [[ "${OSTYPE:-}" == "msys" || "${OSTYPE:-}" == "win32" ]]; then
  SEP=";"
fi

# Ruta al ejecutable de PyInstaller
PYINSTALLER_BIN="pyinstaller"
if [ -f "$HOME/Library/Python/3.9/bin/pyinstaller" ]; then
  PYINSTALLER_BIN="$HOME/Library/Python/3.9/bin/pyinstaller"
fi

$PYINSTALLER_BIN --name "Derma Essenza CRM" \
  --onefile \
  --noconsole \
  --add-data "templates${SEP}templates" \
  --add-data "static${SEP}static" \
  --clean \
  desktop.py

echo "=== ¡Compilación completada! ==="
echo "Encontrarás la aplicación en la carpeta: ./dist/"
