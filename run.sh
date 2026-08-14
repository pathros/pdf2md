# Lanza pdf2md: uv si está disponible; si no, el virtualenv de pyenv.
set -euo pipefail
cd "$(dirname "$0")"

if command -v uv >/dev/null 2>&1; then
  exec uv run python pdf2md.py "$@"
fi

if command -v pyenv >/dev/null 2>&1; then
  eval "$(pyenv init -)"
fi
exec python pdf2md.py "$@"
