@echo off
REM Lanza pdf2md desde cmd.exe usando uv.
cd /d "%~dp0"
uv run python pdf2md.py %*
