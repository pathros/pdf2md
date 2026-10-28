# pdf2md

Transcribe **cualquier PDF** (digital o escaneado, **n páginas**) a **Markdown (.md)** o **Word (.docx)** con IA **local** (Ollama + Qwen visión). También acepta imágenes, Word, Excel, PowerPoint y Markdown.

Herramienta de terminal para **Windows, macOS y Linux**. En Windows el atajo es arrastrar el PDF sobre `Transcribir-PDF.bat`.

## Alcance

| Entrada | Qué hace | ¿Tope de páginas? |
|---------|----------|-------------------|
| **PDF digital** (texto seleccionable) | Extrae texto, encabezados y tablas con **PyMuPDF / pymupdf4llm**. | No. Página a página. |
| **PDF escaneado** | OCR de visión con **Ollama / Qwen**. **No usa Tesseract.** | No. Una página a la vez (el límite es el **tiempo**, no la RAM). |
| **PDF mixto** | `--ocr auto`: texto donde hay capa; OCR solo en páginas-imagen. | Igual. |
| **Imágenes** | Mismo OCR; un solo archivo de salida. | N archivos. |
| **Word (.docx)** | Pasa a Markdown con pandoc. Sin OCR. | — |
| **Excel (.xlsx, .xlsm)** | Cada hoja visible → tabla Markdown. Sin OCR. | No. Hoja por hoja. |
| **PowerPoint (.pptx, .pptm)** | Cada diapositiva → Markdown (título, viñetas, tablas, notas). Sin OCR. | No. Diapositiva a diapositiva. |
| **Markdown (.md)** | Pasa a Word (`.docx`): títulos, listas, tablas, citas y código. Sin OCR. | — |

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

### Word → Markdown

Requiere [pandoc](https://pandoc.org) (`brew install pandoc`). En el menú interactivo
elige **3**, o:

```bash
./run.sh word documento.docx
./run.sh documento.docx          # si la ruta termina en .docx, se infiere solo
```

El `.md` sale a `~/Downloads` por defecto (o `-o`).

### Excel → Markdown

Cada hoja **visible** de un `.xlsx` o `.xlsm` se vuelve una tabla de Markdown. No usa Ollama ni pandoc. En el menú interactivo elige **4**, o:

```bash
./run.sh excel datos.xlsx
./run.sh datos.xlsx                       # si la ruta termina en .xlsx, se infiere solo
./run.sh excel datos.xlsx --cell-width 40
./run.sh excel datos.xlsx --cell-width 0  # no parte celdas largas
```

El `.md` sale a `~/Downloads` por defecto (o `-o`). El formato `.xls` (Excel 97-2003) no se lee: hay que guardarlo como `.xlsx`.

Al armar la tabla:

- El carácter `|` dentro de una celda se escribe como `\|`, para que no se confunda con el separador de columnas. También se escapan la barra invertida, la comilla invertida y el asterisco.
- Un salto de línea dentro de la celda pasa a `&lt;br&gt;`. Un salto real partiría la fila.
- Si el texto de una celda supera el ancho (`--cell-width`, 60 por defecto), se parte en ese límite con `&lt;br&gt;`. La fila sigue siendo **una sola línea** de Markdown; el corte se ve al previsualizar.
- `<`, `>` y `&` se escapan para que no se lean como HTML. El `&lt;br&gt;` de los cortes se inserta después, así que sí funciona.
- Fechas como `YYYY-MM-DD`. Porcentajes como `15%` (Excel guarda 0.15). Booleanos como `TRUE` / `FALSE`.
- Si una fórmula no trae el valor que Excel calculó, se copia la fórmula (`=A1+B1`).
- Una fila suelta encima de la tabla (título) queda como párrafo, no como encabezado de columnas.
- Un rango combinado pequeño repite su valor en cada celda, porque Markdown no combina celdas. Un título fusionado a lo ancho se deja en una sola celda y pasa a párrafo.
- Un hueco grande de filas o columnas vacías abre **otra tabla**, para no generar cientos de celdas vacías.
- Hojas ocultas, imágenes y gráficos del libro no se exportan.

### PowerPoint → Markdown

Cada diapositiva de un `.pptx` o `.pptm` (también `.ppsx`) se vuelve una sección de Markdown. No usa Ollama ni pandoc. En el menú interactivo elige **5**, o:

```bash
./run.sh pptx presentacion.pptx
./run.sh presentacion.pptx                # si la ruta termina en .pptx, se infiere solo
./run.sh pptx presentacion.pptx --cell-width 40
```

El `.md` sale a `~/Downloads` por defecto (o `-o`). El formato `.ppt` (PowerPoint 97-2003) no se lee: hay que guardarlo como `.pptx`.

- El título de la diapositiva queda en el encabezado (`## Diapositiva 1. …`). El texto del cuerpo sale como lista, con el nivel de la viñeta.
- Las tablas usan las mismas precauciones que Excel: `\|`, salto de línea como `&lt;br&gt;` y corte de celdas largas con `--cell-width`.
- Si el gráfico trae categorías y series, se vuelcan a una tabla. Si no, queda el título del gráfico.
- Las notas del orador van al final de la diapositiva, bajo `### Notas`.
- Una diapositiva oculta se marca `(oculta)` y sí se incluye.
- Las imágenes sin texto alternativo y los diagramas SmartArt no se transcriben: el archivo no trae ese texto.

### Markdown → Word

Un `.md` (también `.markdown`) se vuelve un `.docx`. No usa Ollama ni pandoc: es el mismo exportador con el que un PDF pasa a Word.

Arranca el menú con `./run.sh` (o el alias `pdf2md`) y elige **6**. Ahí pide la ruta: arrastra el `.md` a la terminal y pulsa Enter. El `.docx` sale a `~/Downloads`.

```bash
./run.sh
```

Si ya tienes la ruta a mano, también vale pasarla directo (`./run.sh notas.md`, o `-o` para otra carpeta). Siempre es Word; `--format` no aplica en este modo.

- Títulos, listas (con el nivel de la sangría), tablas, citas (`>`) y bloques de código.
- Varias líneas seguidas, sin línea en blanco, forman un solo párrafo.
- Un `<br>` dentro de una celda (como los que escribe Excel o PowerPoint) pasa a un salto de línea en Word. `\|`, `\\` y las entidades `&amp;` `&lt;` `&gt;` se deshacen.
- El bloque `---` inicial de metadatos (fuente, generado_por) no se copia al Word.
- Una imagen `![texto](ruta)` deja el texto; no incrusta el archivo.

En el menú interactivo, **0** cancela sin traceback (Ctrl+C también sale limpio).

Formatos de imagen: `.png`, `.jpg`, `.jpeg`, `.webp`, `.gif`, `.bmp`, `.tif`, `.tiff`.

### Opciones

| Opción | Descripción |
|--------|-------------|
| `-o, --output` | Ruta o **carpeta** de salida |
| `--format` | `md` o `docx` (PDF e imágenes). El modo Markdown siempre escribe `.docx` |
| `--cell-width` | Excel y PowerPoint: ancho máximo de una celda (predet. 60; `0` = no partir) |
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
4. **Word:** se arma Markdown interno y se exporta a `.docx` (títulos, listas, tablas, salto de página por hoja del PDF). Word → Markdown usa pandoc.
5. **Excel:** openpyxl lee las celdas (valor calculado, o la fórmula si no hay caché) y escribe tablas Markdown escapando `|` y partiendo celdas largas con `&lt;br&gt;`.
6. **PowerPoint:** python-pptx lee cada diapositiva (título, viñetas, tablas, gráficos con datos y notas) y escribe Markdown. Las imágenes sin texto alternativo no se transcriben.
7. **Markdown → Word:** python-docx arma el `.docx` (títulos, listas, tablas, citas, código). Un `<br>` de una celda pasa a salto de línea.

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
