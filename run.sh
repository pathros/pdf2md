# Lanza pdf2md con su propio entorno. No hace falta `source`/activate.
set -euo pipefail
cd "$(dirname "$0")"

# Si quedaste en un venv de pyenv (u otro), uv avisa y lo ignora.
# Lo limpiamos para que siempre use el .venv del proyecto.
unset VIRTUAL_ENV

if command -v uv >/dev/null 2>&1; then
  exec uv run python pdf2md.py "$@"
fi

if [[ -x .venv/bin/python ]]; then
  exec .venv/bin/python pdf2md.py "$@"
fi

if command -v pyenv >/dev/null 2>&1; then
  eval "$(pyenv init -)"
fi
exec python pdf2md.py "$@"
