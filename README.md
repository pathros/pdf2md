# pdf2md

Transcribe **cualquier PDF** (digital o escaneado, **n páginas**) a **Markdown (.md)** o **Word (.docx)** con IA **local** (Ollama + Qwen visión). También acepta imágenes.

Herramienta de terminal para **Windows, macOS y Linux**. En Windows el atajo es arrastrar el PDF sobre `Transcribir-PDF.bat`.

## Alcance

| Entrada | Qué hace | ¿Tope de páginas? |
|---------|----------|-------------------|
| **PDF digital** (texto seleccionable) | Extrae texto, encabezados y tablas con **PyMuPDF / pymupdf4llm**. | No. Página a página. |
| **PDF escaneado** | OCR de visión con **Ollama / Qwen**. **No usa Tesseract.** | No. Una página a la vez (el límite es el **tiempo**, no la RAM). |
| **PDF mixto** | `--ocr auto`: texto donde hay capa; OCR solo en páginas-imagen. | Igual. |
| **Imágenes** | Mismo OCR; un solo archivo de salida. | N archivos. |

Objetivo: transcripción **fiel** (nombres, cifras, tablas), con mínimos errores, **sin enviar el PDF a la nube**.

En CPU (p. ej. Intel i5, 16 GB) el OCR va **lento** (minutos por página). La terminal muestra **barra de progreso** y las **últimas 5 líneas** que se van transcribiendo (estilo log).

## Salida

| Sistema | Por defecto | También |
|---------|-------------|---------|
| **Windows** | `.docx` en `%USERPROFILE%\Documents\Transcripciones-PDF` | Otra carpeta con `-o` o el 2.º argumento del `.bat` |
| **macOS / Linux** | `.md` en `~/Downloads` (en el menú pregunta md o docx) | `--format docx` o `--format md` |

## Requisitos

- [uv](https://docs.astral.sh/uv/) (gestión de Python; se instala en computadoras nuevas)
- [Ollama](https://ollama.com) en ejecución, con un **modelo de visión** (Qwen 3.5/3.6 aceptan imagen)

En macOS/Linux, si ya tenías el entorno con `pyenv`, `setup.sh` sigue funcionando como respaldo.

### Modelos y RAM (Windows 11 · Intel i5 · 16 GB)

`qwen3.5` es **multimodal** (texto + imagen), así que sirve para el OCR.

| Modelo | Peso (Q4 en Ollama) | ¿En 16 GB RAM? |
|--------|---------------------|----------------|
| **`qwen3.5:9b`** | ~6.6 GB | **Sí.** Es el recomendado. Cierra Chrome/Cursor pesados si se pone justo. |
| `qwen3.5:4b` | ~3.4 GB | Sí, más rápido y peor OCR. |
| **`qwen3.5:27b`** | **~17 GB** | **No de forma usable.** El modelo ya supera la RAM. Hace falta **24 GB+** (ideal 32 GB) y, mejor, GPU con 16 GB+ VRAM. |

```bash
ollama pull qwen3.5:9b
```

Si no pasas `--model`, pdf2md elige solo: con menos de 24 GB de RAM prefiere `qwen3.5:9b`.

## Instalación

### Windows 11 (PowerShell)

Si `uv` aún no está (computadora nueva):

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Cierra y abre la terminal, clona el repo y configura:

```powershell
git clone https://github.com/pathros/pdf2md.git
cd pdf2md
powershell -ExecutionPolicy Bypass -File .\setup.ps1
ollama pull qwen3.5:9b
```

### macOS / Linux

```bash
git clone https://github.com/pathros/pdf2md.git
cd pdf2md
./setup.sh
```

uv en computadoras nuevas (macOS/Linux):

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

## Uso

### Windows — transcribir a Word

Arrastra el PDF sobre **`Transcribir-PDF.bat`**, o:

```bat
Transcribir-PDF.bat "C:\Users\yo\Escritorio\acta.pdf"
Transcribir-PDF.bat "C:\Users\yo\Escritorio\acta.pdf" "D:\Salidas"
```

- 1.er argumento: PDF (típicamente escaneado; si tiene texto digital, también lo usa).
- 2.º argumento (opcional): carpeta de salida. Si se omite: `Documentos\Transcripciones-PDF`.
- Resultado: `acta.docx`.

También:

```powershell
.\run.ps1
uv run python pdf2md.py pdf documento.pdf --format docx
```

### macOS / Linux

No hace falta activar ningún venv a mano. El alias o `./run.sh` usan `uv run`
y el `.venv` del proyecto solos:

```bash
pdf2md                    # alias → ~/python/pdf2md/run.sh
# o:
./run.sh
```

El menú pide modo y **formato** (`.md` o `.docx`). Arrastra el archivo a la terminal.

```bash
./run.sh documento.pdf
./run.sh pdf escaneado.pdf --format docx -o ~/Documents/salidas/
./run.sh pdf doc.pdf --ocr never --format md
./run.sh pdf doc.pdf --model qwen3.8:27b   # forzar un modelo concreto
```

Con ≥24 GB de RAM, si no pasas `--model`, prioriza `qwen3.8:27b` cuando esté
instalado (luego `qwen3.6:27b`, etc.).
### Modo imágenes (OCR)

```bash
./run.sh images foto1.png foto2.jpg
./run.sh images a.png --format docx
```

Formatos de imagen: `.png`, `.jpg`, `.jpeg`, `.webp`, `.gif`, `.bmp`, `.tif`, `.tiff`.

### Opciones

| Opción | Descripción |
|--------|-------------|
| `-o, --output` | Ruta o **carpeta** de salida |
| `--format` | `md` o `docx` |
| `--ocr` | Solo PDF: `auto` (default) / `always` / `never` |
| `--model` | Modelo de visión en Ollama |
| `--host` | Host de Ollama (predet. `http://localhost:11434`) |

### Variables de entorno

| Variable | Uso |
|----------|-----|
| `PDF2MD_VISION_MODEL` | Modelo de visión de Ollama |
| `PDF2MD_OUTPUT_DIR` | Carpeta de salida por defecto |
| `PDF2MD_FORMAT` | `md` o `docx` |
| `OLLAMA_HOST` | Host de Ollama |

## Cómo funciona

1. **PDF digital:** pymupdf4llm extrae la capa de texto (sin Tesseract).
2. **PDF escaneado:** cada página se renderiza a PNG y se manda a Ollama (`/api/generate`, streaming). El modelo no resume ni traduce.
3. **n páginas:** el bucle es una página a la vez; un PDF de 200 páginas es válido, tardará proporcionalmente.
4. **Word:** se arma Markdown interno y se exporta a `.docx` (títulos, listas, tablas, salto de página por hoja del PDF).

## Estructura

```
pdf2md/
├── pdf2md.py
├── Transcribir-PDF.bat   # Windows: arrastra el PDF → .docx
├── setup.ps1 / run.ps1 / run.cmd
├── setup.sh / run.sh
├── pyproject.toml / uv.lock / requirements.txt
└── README.md
```

## Licencia

Uso personal / proyecto open source a tu criterio. Añade un archivo `LICENSE` si publicas el repo con una licencia concreta (MIT, Apache-2.0, etc.).
