@echo off
REM Transcribir-PDF.bat
REM Arrastra un PDF sobre este archivo, o:
REM   Transcribir-PDF.bat "C:\ruta\documento.pdf"
REM   Transcribir-PDF.bat "C:\ruta\documento.pdf" "D:\Salidas"
REM
REM Por defecto guarda un .docx en:
REM   %USERPROFILE%\Documents\Transcripciones-PDF
REM (o la carpeta de PDF2MD_OUTPUT_DIR si está definida)

setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" (
    echo.
    echo  Transcribir-PDF — IA local ^(Ollama / Qwen^)
    echo  -------------------------------------------
    echo  Arrastra un PDF sobre este archivo, o ejecuta:
    echo.
    echo    Transcribir-PDF.bat documento.pdf
    echo    Transcribir-PDF.bat documento.pdf "D:\CarpetaDeSalida"
    echo.
    echo  Salida por defecto:
    echo    %USERPROFILE%\Documents\Transcripciones-PDF
    echo.
    pause
    exit /b 1
)

set "PDF=%~1"
if not exist "%PDF%" (
    echo No existe el archivo: %PDF%
    pause
    exit /b 1
)

set "OUTDIR=%USERPROFILE%\Documents\Transcripciones-PDF"
if defined PDF2MD_OUTPUT_DIR set "OUTDIR=%PDF2MD_OUTPUT_DIR%"
if not "%~2"=="" set "OUTDIR=%~2"

set "UV="
where uv >nul 2>&1 && set "UV=uv"
if not defined UV if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV=%USERPROFILE%\.cargo\bin\uv.exe"

if not defined UV (
    echo uv no está en el PATH. Ejecuta primero setup.ps1
    echo Instalación:
    echo   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    pause
    exit /b 1
)

echo.
echo Transcribiendo:
echo   PDF     %PDF%
echo   Salida  %OUTDIR%
echo   Formato Word ^(.docx^)
echo.
echo Deja Ollama en ejecución. En CPU el OCR puede tardar minutos por página.
echo.

"%UV%" run python pdf2md.py pdf "%PDF%" --ocr auto --format docx -o "%OUTDIR%"
set "ERR=%ERRORLEVEL%"

echo.
if %ERR%==0 (
    echo Listo. Revisa la carpeta:
    echo   %OUTDIR%
) else (
    echo Falló la transcripción ^(código %ERR%^).
)

if not defined PDF2MD_NOPAUSE pause
exit /b %ERR%
