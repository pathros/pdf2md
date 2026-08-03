# pdf2md

Convierte el contenido textual de un **PDF** o de **imágenes** a **Markdown (.md)** limpio y bien estructurado, pensado para que cualquier **LLM** lo digiera fácilmente.

Herramienta de terminal para macOS/Linux: arrastras el archivo, eliges el modo y obtienes un `.md` en `~/Downloads` (por defecto), con barra de progreso.

## Por qué existe (ahorro de tokens)

Los **PDF escaneados** (páginas como imagen, sin texto seleccionable) suelen consumir **muchísimos tokens** cuando los subes directo a ChatGPT, Claude u otros modelos de OpenAI/Anthropic: la IA tiene que “leer” la imagen página por página y eso encarece el uso.

`pdf2md` está pensado para un flujo más barato:

1. Lees el PDF escaneado (o las capturas) **en local** con un modelo de visión vía Ollama.
2. Obtienes un `.md` con el texto ya estructurado.
3. Usas ese Markdown en Claude, ChatGPT u otra IA en la nube.

Así reduces el gasto de tokens: en la nube procesas texto, no páginas escaneadas pesadas.

## Qué hace

| Entrada | Comportamiento |
|---------|----------------|
| **PDF con texto** | Extrae texto, encabezados y tablas con **PyMuPDF / pymupdf4llm** (rápido). |
| **PDF escaneado** | Detecta páginas sin capa de texto y las transcribe con **OCR de visión vía Ollama**. **No usa Tesseract.** |
| **Imágenes** | Transcribe N fotos/capturas con el mismo OCR y las junta en un solo `.md`. |

## Requisitos

- macOS o Linux con [`pyenv`](https://github.com/pyenv/pyenv) y [`pyenv-virtualenv`](https://github.com/pyenv/pyenv-virtualenv)
- [Ollama](https://ollama.com) en ejecución (`ollama serve`) con un modelo de visión

Modelos recomendados:

```bash
ollama pull qwen3.6:27b
# o, especializado en OCR → Markdown:
ollama pull benhaotang/Nanonets-OCR-s:F16
```

## Instalación

```bash
git clone https://github.com/pathros/pdf2md.git
cd pdf2md
./setup.sh
```

Esto crea el virtualenv `pdf2md` (Python 3.12.11) con `pyenv virtualenv`, lo activa en la carpeta (`pyenv local`) e instala las dependencias de `requirements.txt`.

## Uso

```bash
./run.sh
```

Elige **1) PDF** o **2) Imágenes** y arrastra el/los archivo(s) a la terminal cuando lo pida.

Alias opcional (ajusta la ruta a donde clonaste el repo):

```bash
echo "alias pdf2md='$HOME/python/pdf2md/run.sh'" >> ~/.zshrc
source ~/.zshrc
pdf2md
```

### Modo PDF

```bash
./run.sh                          # menú interactivo
./run.sh documento.pdf            # PDF directo
./run.sh pdf documento.pdf
```

### Modo imágenes (OCR)

```bash
./run.sh images
# te pide las rutas → arrastra N imágenes → Enter

./run.sh images foto1.png foto2.jpg captura.webp
./run.sh img ~/Desktop/a.png ~/Desktop/b.png -o ~/Downloads/notas.md
```

Formatos: `.png`, `.jpg`, `.jpeg`, `.webp`, `.gif`, `.bmp`, `.tif`, `.tiff`.

- 1 imagen → el `.md` se llama como la imagen (`foto.md`)
- Varias → `ocr-imagenes-YYYYMMDD-HHMMSS.md`

### Opciones

| Opción | Descripción |
|--------|-------------|
| `-o, --output` | Ruta o carpeta de salida `.md` (por defecto `~/Downloads`) |
| `--ocr` | Solo PDF: `auto` / `always` / `never` |
| `--model` | Modelo de visión en Ollama (predet. `qwen3.6:27b`) |
| `--host` | Host de Ollama (predet. `http://localhost:11434`) |

```bash
./run.sh pdf escaneado.pdf --ocr always
./run.sh images a.png b.png --model benhaotang/Nanonets-OCR-s:F16
./run.sh pdf doc.pdf --ocr never -o ~/Documents/salidas/
```

### Variables de entorno

| Variable | Uso |
|----------|-----|
| `PDF2MD_VISION_MODEL` | Modelo de visión de Ollama por defecto |
| `PDF2MD_OUTPUT_DIR` | Carpeta de salida por defecto |
| `OLLAMA_HOST` | Host de Ollama |

## Cómo funciona el OCR

1. **PDF:** marca páginas sin capa de texto (con imágenes) como escaneadas. El texto embebido se extrae con pymupdf4llm. Las páginas escaneadas se renderizan a PNG y se envían a Ollama.
2. **Imágenes:** cada archivo se manda a `/api/generate` de Ollama con un prompt de transcripción fiel a Markdown (sin traducir ni resumir) y se concatena en un único `.md`.

## Estructura

```
pdf2md/
├── pdf2md.py          # CLI principal
├── run.sh             # Lanza la app con el virtualenv de pyenv
├── setup.sh           # Crea el entorno e instala dependencias
├── requirements.txt
└── README.md
```

## Licencia

Uso personal / proyecto open source a tu criterio. Añade un archivo `LICENSE` si publicas el repo con una licencia concreta (MIT, Apache-2.0, etc.).
