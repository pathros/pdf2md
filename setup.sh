#!/usr/bin/env bash
# Configura el entorno con pyenv + pyenv-virtualenv e instala dependencias.
set -euo pipefail

PY_VERSION="3.12.11"
ENV_NAME="pdf2md"

cd "$(dirname "$0")"

# Asegura que pyenv esté disponible en este shell.
if ! command -v pyenv >/dev/null 2>&1; then
  echo "pyenv no está instalado o no está en el PATH." >&2
  exit 1
fi
eval "$(pyenv init -)"
eval "$(pyenv virtualenv-init -)" 2>/dev/null || true

# Instala la versión de Python si falta (-s: omite si ya existe, sin preguntar).
echo ">> Verificando Python $PY_VERSION ..."
pyenv install -s "$PY_VERSION"

# Crea el virtualenv si no existe.
if ! pyenv virtualenvs --bare | grep -qx "${ENV_NAME}"; then
  echo ">> Creando virtualenv $ENV_NAME ..."
  pyenv virtualenv "$PY_VERSION" "$ENV_NAME"
else
  echo ">> El virtualenv $ENV_NAME ya existe."
fi

# Activa automáticamente el env al entrar a esta carpeta.
pyenv local "$ENV_NAME"

echo ">> Instalando dependencias ..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "Listo. Ejecuta la app con:"
echo "  ./run.sh"
echo "o bien:"
echo "  python pdf2md.py"
