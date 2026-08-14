# Configura pdf2md en Windows con uv (Python 3.12 + dependencias).
# Uso (PowerShell):
#   powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Get-UvCommand {
    $cmd = Get-Command uv -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $candidates = @(
        (Join-Path $env:USERPROFILE ".local\bin\uv.exe"),
        (Join-Path $env:USERPROFILE ".cargo\bin\uv.exe")
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

$uv = Get-UvCommand
if (-not $uv) {
    Write-Host "uv no está instalado o no está en el PATH."
    Write-Host "En computadoras nuevas, instálalo con:"
    Write-Host '  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"'
    Write-Host "Cierra y vuelve a abrir la terminal, luego ejecuta de nuevo .\setup.ps1"
    exit 1
}

Write-Host ">> Usando uv: $uv"
Write-Host ">> Instalando Python 3.12 (si falta) y dependencias ..."
& $uv python install 3.12
& $uv sync --python 3.12

Write-Host ""
Write-Host "Listo. Para transcribir un PDF a Word, arrastra el archivo sobre:"
Write-Host "  Transcribir-PDF.bat"
Write-Host "o en PowerShell:"
Write-Host "  .\run.ps1"
Write-Host "  uv run python pdf2md.py"
Write-Host ""
Write-Host "Ollama debe estar en ejecución con un modelo de visión."
Write-Host "Recomendado en 16 GB RAM:  ollama pull qwen3.5:9b"
Write-Host "El .docx se guarda en Documentos\Transcripciones-PDF (salvo que pases otra carpeta)."
