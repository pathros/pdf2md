# Lanza pdf2md en Windows con el entorno de uv.
# Uso:  .\run.ps1
#       .\run.ps1 documento.pdf
#       .\run.ps1 images foto1.png foto2.jpg

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
    Write-Host "uv no está instalado. Ejecuta primero .\setup.ps1"
    Write-Host 'Instalación de uv: powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"'
    exit 1
}

& $uv run python pdf2md.py @args
exit $LASTEXITCODE
