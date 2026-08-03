#!/usr/bin/env bash
# Lanza pdf2md usando el virtualenv de pyenv de esta carpeta.
set -euo pipefail
cd "$(dirname "$0")"
if command -v pyenv >/dev/null 2>&1; then
  eval "$(pyenv init -)"
fi
exec python pdf2md.py "$@"
