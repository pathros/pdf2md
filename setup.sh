# Configura el entorno con uv (preferido) o pyenv + pyenv-virtualenv.
set -euo pipefail

PY_VERSION="3.12"
ENV_NAME="pdf2md"

cd "$(dirname "$0")"

if command -v uv >/dev/null 2>&1; then
  echo ">> Usando uv para Python $PY_VERSION e instalación de dependencias ..."
  uv python install "$PY_VERSION"
  uv sync --python "$PY_VERSION"
  echo ""
  echo "Listo. Ejecuta la app con:"
  echo "  ./run.sh"
  echo "o bien:"
  echo "  uv run python pdf2md.py"
  exit 0
fi

echo ">> uv no está en el PATH; se intenta con pyenv (entorno macOS/Linux previo)."
echo "   En computadoras nuevas instala uv:"
echo "     curl -LsSf https://astral.sh/uv/install.sh | sh"
echo ""

if ! command -v pyenv >/dev/null 2>&1; then
  echo "Ni uv ni pyenv están instalados." >&2
  exit 1
fi

eval "$(pyenv init -)"
eval "$(pyenv virtualenv-init -)" 2>/dev/null || true

echo ">> Verificando Python ${PY_VERSION}.11 ..."
pyenv install -s "${PY_VERSION}.11"

if ! pyenv virtualenvs --bare | grep -qx "${ENV_NAME}"; then
  echo ">> Creando virtualenv $ENV_NAME ..."
  pyenv virtualenv "${PY_VERSION}.11" "$ENV_NAME"
else
  echo ">> El virtualenv $ENV_NAME ya existe."
fi

pyenv local "$ENV_NAME"

echo ">> Instalando dependencias ..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "Listo. Ejecuta la app con:"
echo "  ./run.sh"
echo "o bien:"
echo "  python pdf2md.py"
