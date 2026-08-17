#!/usr/bin/env python3
"""pdf2md - Transcribe PDF o imágenes a Markdown o Word con IA local.

Modos:
  pdf     Extrae texto/tablas con pymupdf4llm. Páginas escaneadas → OCR con
          Ollama (visión). Sin límite de páginas (una a la vez).
  images  Transcribe N imágenes con el mismo OCR y las junta en un archivo.

Salida: Windows → .docx por defecto; macOS/Linux → .md o .docx (--format).
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shlex
import sys
from datetime import datetime
from pathlib import Path

import requests
from docx import Document
from rich.console import Console, Group
from rich.live import Live
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.prompt import Prompt
from rich.text import Text

import pymupdf  # PyMuPDF
import pymupdf4llm

console = Console()

# --- Configuración por defecto (sobreescribible por CLI o variables de entorno) ---
DEFAULT_OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
# Fallback si Ollama no responde o no hay modelo conocido instalado.
DEFAULT_VISION_MODEL = os.environ.get("PDF2MD_VISION_MODEL", "qwen3.5:9b")

# Preferencia cuando hay varios modelos de visión en Ollama.
# En máquinas con <24 GB de RAM se prioriza el 9b (cabe en 16 GB).
VISION_MODELS_LOW_RAM = (
    "qwen3.5:9b",
    "qwen3.5:latest",
    "qwen3.5:4b",
    "qwen3.6:9b",
    "benhaotang/Nanonets-OCR-s:F16",
)
VISION_MODELS_HIGH_RAM = (
    "qwen3.8:27b",
    "qwen3.6:27b",
    "qwen3.5:27b",
    *VISION_MODELS_LOW_RAM,
)
# Por debajo de esto, un 27B Q4 (~17 GB) pelea con el SO y suele ir a swap.
HIGH_RAM_GB = 24.0

# Si una página tiene menos de estos caracteres de texto real y contiene imágenes,
# se considera escaneada (candidata a OCR).
OCR_TEXT_THRESHOLD = 20

# DPI de render para el OCR con Ollama (zoom 2.5 ~= 180 dpi).
OCR_RENDER_ZOOM = 2.5

IMAGE_EXTENSIONS = frozenset({
    ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tif", ".tiff",
})

OCR_PROMPT = (
    "Eres un motor de OCR. Transcribe TODO el texto visible de esta imagen a "
    "Markdown limpio y bien estructurado. Respeta el idioma original (no traduzcas). "
    "Copia con fidelidad nombres, fechas, cifras, sellos y puntuación. "
    "Usa encabezados (#), listas y tablas de Markdown cuando corresponda y conserva "
    "el orden de lectura. Si un fragmento es ilegible, escribe [ilegible]; no inventes. "
    "NO resumas y NO agregues comentarios ni explicaciones: "
    "devuelve únicamente la transcripción en Markdown."
)

PREVIEW_LINE_COUNT = 5
PREVIEW_LINE_WIDTH = 110

PROGRESS_COLUMNS = (
    SpinnerColumn(),
    TextColumn("[progress.description]{task.description}"),
    BarColumn(),
    TextColumn("{task.completed}/{task.total}"),
    TimeElapsedColumn(),
    TimeRemainingColumn(),
)


def default_output_dir() -> Path:
    env = os.environ.get("PDF2MD_OUTPUT_DIR")
    if env:
        return Path(os.path.expanduser(env))
    if os.name == "nt":
        return Path.home() / "Documents" / "Transcripciones-PDF"
    return Path.home() / "Downloads"


def default_output_format() -> str:
    env = (os.environ.get("PDF2MD_FORMAT") or "").strip().lower()
    if env in ("md", "markdown"):
        return "md"
    if env in ("docx", "doc", "word"):
        return "docx"
    return "docx" if os.name == "nt" else "md"


DEFAULT_OUTPUT_DIR = default_output_dir()


# --------------------------------------------------------------------------- #
# Hardware / modelo de visión
# --------------------------------------------------------------------------- #
def total_ram_gb() -> float | None:
    """RAM física total en GB, o None si no se pudo leer."""
    try:
        if sys.platform == "darwin":
            import subprocess

            out = subprocess.check_output(["sysctl", "-n", "hw.memsize"], text=True)
            return int(out.strip()) / (1024 ** 3)
        if sys.platform == "win32":
            import ctypes

            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]

            stat = MEMORYSTATUSEX()
            stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
                return None
            return stat.ullTotalPhys / (1024 ** 3)
        with open("/proc/meminfo", encoding="utf-8") as fh:
            for line in fh:
                if line.startswith("MemTotal:"):
                    return int(line.split()[1]) / (1024 ** 2)
    except Exception:  # noqa: BLE001
        return None
    return None


def list_ollama_model_names(host: str) -> set[str]:
    try:
        resp = requests.get(f"{host}/api/tags", timeout=5)
        resp.raise_for_status()
    except requests.RequestException:
        return set()
    return {m.get("name", "") for m in resp.json().get("models", []) if m.get("name")}


def model_is_installed(names: set[str], model: str) -> bool:
    """True solo si el tag pedido está instalado (no basta la misma familia)."""
    if model in names:
        return True
    if ":" not in model and f"{model}:latest" in names:
        return True
    return False


def resolve_vision_model(host: str, requested: str | None) -> str:
    """Elige el modelo de visión: CLI / env / lo instalado según RAM."""
    if requested:
        return requested
    env = os.environ.get("PDF2MD_VISION_MODEL")
    if env:
        return env

    ram = total_ram_gb()
    preferred = VISION_MODELS_HIGH_RAM if (ram or 0) >= HIGH_RAM_GB else VISION_MODELS_LOW_RAM
    names = list_ollama_model_names(host)
    for candidate in preferred:
        if model_is_installed(names, candidate):
            return candidate
    return DEFAULT_VISION_MODEL


def warn_if_model_too_heavy(model: str) -> None:
    ram = total_ram_gb()
    if ram is None:
        return
    heavy = any(tag in model.lower() for tag in (":27b", ":35b", ":122b"))
    if heavy and ram < HIGH_RAM_GB:
        console.print(
            f"[yellow]Aviso:[/yellow] '{model}' pesa ~17 GB+ (Q4). "
            f"Esta máquina tiene ~{ram:.0f} GB de RAM; Windows + Ollama "
            "suelen quedarse cortos y el 27b irá a swap o no cargará. "
            "Recomendado: [bold]qwen3.5:9b[/bold] (~6.6 GB)."
        )


# --------------------------------------------------------------------------- #
# Utilidades de ruta / entrada
# --------------------------------------------------------------------------- #
def _strip_wrapping_quotes(value: str) -> str:
    s = value.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
        return s[1:-1]
    return s


def clean_dropped_path(raw: str) -> str:
    """Normaliza una ruta arrastrada a la terminal (Finder, Explorer o PowerShell)."""
    s = raw.strip()
    if s.startswith("&"):
        s = s[1:].strip()
    s = _strip_wrapping_quotes(s)
    if os.name != "nt":
        for esc in ("\\ ", "\\(", "\\)", "\\&", "\\'", "\\\""):
            s = s.replace(esc, esc[1])
    return os.path.expanduser(s.strip())


def parse_dropped_paths(raw: str) -> list[Path]:
    """Separa varias rutas arrastradas (Finder en macOS o Explorer en Windows)."""
    raw = raw.strip()
    if not raw:
        return []
    posix = os.name != "nt"
    to_split = raw.replace("& ", " ") if os.name == "nt" else raw
    try:
        tokens = shlex.split(to_split, posix=posix)
    except ValueError:
        tokens = to_split.split()
    paths: list[Path] = []
    for token in tokens:
        cleaned = clean_dropped_path(token)
        if cleaned:
            paths.append(Path(cleaned))
    return paths


def is_image_path(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def prompt_for_mode() -> str:
    console.print(
        Panel.fit(
            "[bold]1[/bold]  PDF → Markdown / Word\n"
            "[bold]2[/bold]  Imágenes → Markdown / Word (OCR con Ollama)",
            title="pdf2md",
            border_style="cyan",
        )
    )
    while True:
        choice = Prompt.ask(
            "[cyan]Modo[/cyan]",
            choices=["1", "2", "pdf", "images", "img", "imgs"],
            default="1",
            show_choices=False,
        ).strip().lower()
        if choice in ("1", "pdf"):
            return "pdf"
        if choice in ("2", "images", "img", "imgs"):
            return "images"
        console.print("[yellow]Elige 1 (PDF) o 2 (Imágenes).[/yellow]")


def prompt_for_pdf() -> Path:
    """Pide interactivamente la ruta del PDF hasta obtener una válida."""
    console.print(
        Panel.fit(
            "[bold]Arrastra el archivo PDF a esta terminal[/bold] y pulsa Enter.\n"
            "También puedes pegar o escribir la ruta.",
            title="pdf2md · PDF",
            border_style="cyan",
        )
    )
    while True:
        raw = Prompt.ask("[cyan]Ruta del PDF[/cyan]")
        if not raw.strip():
            console.print("[yellow]No escribiste nada. Intenta de nuevo.[/yellow]")
            continue
        path = Path(clean_dropped_path(raw))
        if not path.exists():
            console.print(f"[red]No existe:[/red] {path}")
            continue
        if path.suffix.lower() != ".pdf":
            console.print("[red]El archivo no parece un .pdf[/red]")
            continue
        return path


def prompt_for_images() -> list[Path]:
    """Pide una o más rutas de imagen (arrastrar desde Finder o Explorer)."""
    console.print(
        Panel.fit(
            "[bold]Arrastra una o más imágenes a esta terminal[/bold] y pulsa Enter.\n"
            "Las rutas quedan separadas por espacio (Finder/Explorer lo hace solo).\n"
            f"Formatos: {', '.join(sorted(IMAGE_EXTENSIONS))}",
            title="pdf2md · Imágenes",
            border_style="cyan",
        )
    )
    while True:
        raw = Prompt.ask("[cyan]Rutas de imágenes[/cyan]")
        paths = parse_dropped_paths(raw)
        if not paths:
            console.print("[yellow]No escribiste nada. Intenta de nuevo.[/yellow]")
            continue
        ok, err = validate_image_paths(paths)
        if err:
            for msg in err:
                console.print(f"[red]{msg}[/red]")
            continue
        return ok


def validate_image_paths(paths: list[Path]) -> tuple[list[Path], list[str]]:
    ok: list[Path] = []
    err: list[str] = []
    seen: set[Path] = set()
    for path in paths:
        resolved = path.expanduser()
        if resolved in seen:
            continue
        seen.add(resolved)
        if not resolved.exists():
            err.append(f"No existe: {resolved}")
            continue
        if not resolved.is_file():
            err.append(f"No es un archivo: {resolved}")
            continue
        if not is_image_path(resolved):
            err.append(f"No es una imagen soportada: {resolved}")
            continue
        ok.append(resolved)
    return ok, err


# --------------------------------------------------------------------------- #
# Detección de páginas escaneadas
# --------------------------------------------------------------------------- #
def detect_scanned_pages(pdf_path: Path) -> list[bool]:
    """Devuelve, por página, si parece escaneada (sin capa de texto, con imagen)."""
    doc = pymupdf.open(pdf_path)
    flags: list[bool] = []
    try:
        for page in doc:
            text = page.get_text("text").strip()
            has_images = bool(page.get_images(full=True))
            flags.append(len(text) < OCR_TEXT_THRESHOLD and has_images)
    finally:
        doc.close()
    return flags


# --------------------------------------------------------------------------- #
# OCR con Ollama (modelo de visión) + vista en vivo
# --------------------------------------------------------------------------- #
class TranscriptionDisplay:
    """Barra de progreso + últimas N líneas (estilo log de Docker)."""

    def __init__(self, total: int, label: str = "Procesando") -> None:
        self.progress = Progress(*PROGRESS_COLUMNS, expand=True)
        self.task = self.progress.add_task(label, total=total)
        self._buffer = ""
        self._live: Live | None = None

    def _preview_lines(self) -> list[str]:
        lines = self._buffer.splitlines()
        last = lines[-PREVIEW_LINE_COUNT:] if lines else []
        last = [ln[:PREVIEW_LINE_WIDTH] for ln in last]
        while len(last) < PREVIEW_LINE_COUNT:
            last.append("")
        return last

    def _renderable(self) -> Group:
        body = "\n".join(self._preview_lines()) or " "
        panel = Panel(
            Text(body, style="dim"),
            title="Transcripción (en vivo)",
            subtitle="últimas 5 líneas",
            border_style="cyan",
            height=PREVIEW_LINE_COUNT + 2,
        )
        return Group(self.progress, panel)

    def _refresh(self) -> None:
        if self._live is not None:
            self._live.update(self._renderable())

    def __enter__(self) -> TranscriptionDisplay:
        self._live = Live(
            self._renderable(),
            console=console,
            refresh_per_second=8,
            vertical_overflow="crop",
        )
        self._live.__enter__()
        return self

    def __exit__(self, *args: object) -> None:
        if self._live is not None:
            self._live.__exit__(*args)

    def set_status(self, description: str) -> None:
        self.progress.update(self.task, description=description)
        self._refresh()

    def start_unit(self, description: str) -> None:
        self._buffer = ""
        self.set_status(description)

    def feed(self, chunk: str) -> None:
        self._buffer += chunk
        self._refresh()

    def show_text(self, text: str) -> None:
        self._buffer = text.strip()
        self._refresh()

    def advance(self) -> None:
        self.progress.advance(self.task)
        self._refresh()


def render_page_png(page: "pymupdf.Page", zoom: float = OCR_RENDER_ZOOM) -> bytes:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
    return pix.tobytes("png")


def check_ollama(host: str, model: str) -> bool:
    """Verifica que Ollama responda y que el tag del modelo esté instalado."""
    names = list_ollama_model_names(host)
    if not names:
        return False
    return model_is_installed(names, model)


def ocr_image_bytes_with_ollama(
    image_bytes: bytes,
    model: str,
    host: str,
    on_chunk: object | None = None,
) -> str:
    payload = {
        "model": model,
        "prompt": OCR_PROMPT,
        "images": [base64.b64encode(image_bytes).decode("ascii")],
        "stream": True,
        # Qwen 3.5/3.6 "piensan" por defecto; para OCR queremos solo la transcripción.
        "think": False,
        "options": {"temperature": 0},
    }
    pieces: list[str] = []
    with requests.post(
        f"{host}/api/generate",
        json=payload,
        stream=True,
        timeout=(15, 1800),
    ) as resp:
        resp.raise_for_status()
        for raw_line in resp.iter_lines(decode_unicode=True):
            if not raw_line:
                continue
            try:
                data = json.loads(raw_line)
            except json.JSONDecodeError:
                continue
            token = data.get("response") or ""
            if token:
                pieces.append(token)
                if callable(on_chunk):
                    on_chunk(token)
            if data.get("done"):
                break
    return "".join(pieces).strip()


def ocr_page_with_ollama(
    png_bytes: bytes,
    model: str,
    host: str,
    on_chunk: object | None = None,
) -> str:
    return ocr_image_bytes_with_ollama(png_bytes, model, host, on_chunk=on_chunk)


# --------------------------------------------------------------------------- #
# Markdown → Word (.docx)
# --------------------------------------------------------------------------- #
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_UL_RE = re.compile(r"^[-*+]\s+(.*)$")
_OL_RE = re.compile(r"^(\d+)[.)]\s+(.*)$")
_PAGE_MARK_RE = re.compile(r"^<!--\s*página\s+\d+\s*-->\s*$", re.I)
_HTML_COMMENT_RE = re.compile(r"^<!--.*?-->\s*$")
_TABLE_ROW_RE = re.compile(r"^\|(.+)\|\s*$")
_TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$")
_INLINE_RE = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")


def _strip_frontmatter(md: str) -> str:
    if not md.startswith("---"):
        return md
    end = md.find("\n---\n", 3)
    if end == -1:
        return md
    return md[end + 5 :]


def _add_inline_runs(paragraph, text: str) -> None:
    for part in _INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        elif part.startswith("`") and part.endswith("`") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            paragraph.add_run(part)


def _split_table_row(line: str) -> list[str]:
    inner = line.strip()
    if inner.startswith("|"):
        inner = inner[1:]
    if inner.endswith("|"):
        inner = inner[:-1]
    return [c.strip() for c in inner.split("|")]


def markdown_to_docx(md: str, path: Path) -> None:
    """Convierte el Markdown de la transcripción a un .docx editable en Word."""
    doc = Document()
    body = _strip_frontmatter(md)
    lines = body.splitlines()
    i = 0
    page_breaks = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if _PAGE_MARK_RE.match(stripped):
            if page_breaks:
                doc.add_page_break()
            page_breaks += 1
            i += 1
            continue

        if _HTML_COMMENT_RE.match(stripped):
            i += 1
            continue

        heading = _HEADING_RE.match(stripped)
        if heading:
            level = min(len(heading.group(1)), 9)
            text = heading.group(2).strip()
            if level == 1 and page_breaks <= 1 and i < 8:
                doc.add_heading(text, level=0)
            else:
                doc.add_heading(text, level=level)
            i += 1
            continue

        if (
            _TABLE_ROW_RE.match(stripped)
            and i + 1 < len(lines)
            and _TABLE_SEP_RE.match(lines[i + 1].strip())
        ):
            rows: list[list[str]] = [_split_table_row(stripped)]
            i += 2
            while i < len(lines) and _TABLE_ROW_RE.match(lines[i].strip()):
                rows.append(_split_table_row(lines[i].strip()))
                i += 1
            cols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=cols)
            table.style = "Table Grid"
            for r_idx, row in enumerate(rows):
                for c_idx in range(cols):
                    cell_text = row[c_idx] if c_idx < len(row) else ""
                    table.rows[r_idx].cells[c_idx].text = cell_text
            continue

        ul = _UL_RE.match(stripped)
        if ul:
            p = doc.add_paragraph(style="List Bullet")
            _add_inline_runs(p, ul.group(1))
            i += 1
            continue

        ol = _OL_RE.match(stripped)
        if ol:
            p = doc.add_paragraph(style="List Number")
            _add_inline_runs(p, ol.group(2))
            i += 1
            continue

        if stripped in ("---", "***", "___"):
            i += 1
            continue

        p = doc.add_paragraph()
        _add_inline_runs(p, stripped)
        i += 1

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def output_extension(fmt: str) -> str:
    return ".docx" if fmt == "docx" else ".md"


def resolve_output_format(explicit: str | None, output: str | None) -> str:
    if explicit in ("md", "docx"):
        return explicit
    if output:
        suffix = Path(os.path.expanduser(output)).suffix.lower()
        if suffix == ".docx":
            return "docx"
        if suffix == ".md":
            return "md"
    return default_output_format()


def resolve_output_path(source_stem: str, output: str | None, fmt: str) -> Path:
    ext = output_extension(fmt)
    if output:
        out = Path(os.path.expanduser(output))
        looks_like_dir = (
            out.is_dir()
            or output.endswith(("/", "\\"))
            or out.suffix == ""
        )
        if looks_like_dir:
            return out / f"{source_stem}{ext}"
        if out.suffix.lower() not in (".md", ".docx"):
            out = out.with_suffix(ext)
        return out
    dest = default_output_dir()
    dest.mkdir(parents=True, exist_ok=True)
    return dest / f"{source_stem}{ext}"


def write_output(markdown: str, out_path: Path, fmt: str) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if fmt == "docx" or out_path.suffix.lower() == ".docx":
        markdown_to_docx(markdown, out_path)
        return
    out_path.write_text(markdown, encoding="utf-8")


# --------------------------------------------------------------------------- #
# Construcción del Markdown (PDF)
# --------------------------------------------------------------------------- #
def build_markdown(
    pdf_path: Path,
    ocr_mode: str,
    model: str,
    host: str,
    scanned: list[bool],
) -> str:
    """Genera el Markdown completo del PDF.

    Las páginas con texto se extraen con pymupdf4llm (sin Tesseract). Las páginas
    escaneadas (o todas, si ocr_mode='always') se transcriben con Ollama, una a
    una: no hay tope de páginas (el límite es el tiempo de OCR).
    """
    doc = pymupdf.open(pdf_path)
    total = doc.page_count
    parts: list[str] = []

    ocr_available = True
    if ocr_mode != "never":
        ocr_available = check_ollama(host, model)
        if not ocr_available:
            msg = f"Ollama/modelo '{model}' no disponible en {host}."
            if ocr_mode == "always" or any(scanned):
                console.print(
                    f"[yellow]{msg} Las páginas escaneadas quedarán vacías "
                    f"(¿está corriendo 'ollama serve'?).[/yellow]"
                )

    try:
        need_text_extract = ocr_mode != "always"
        chunks: list[dict] = []
        if need_text_extract:
            # use_ocr=False -> pymupdf4llm NO llama a Tesseract; solo extrae el
            # texto existente. Las páginas escaneadas se cubren con Ollama.
            chunks = pymupdf4llm.to_markdown(
                doc, page_chunks=True, use_ocr=False, show_progress=False
            )
        with TranscriptionDisplay(total, "Procesando") as view:
            for i in range(total):
                do_ocr = ocr_mode == "always" or (ocr_mode == "auto" and scanned[i])
                if do_ocr and ocr_available:
                    view.start_unit(f"OCR (visión) pág. {i + 1}/{total}")
                    try:
                        png = render_page_png(doc.load_page(i))
                        text = ocr_page_with_ollama(
                            png, model, host, on_chunk=view.feed
                        )
                    except requests.RequestException as exc:
                        console.print(f"[red]Error de OCR pág. {i + 1}: {exc}[/red]")
                        text = ""
                else:
                    view.start_unit(f"Texto pág. {i + 1}/{total}")
                    text = (
                        chunks[i].get("text", "") if i < len(chunks) else ""
                    ).strip()
                    view.show_text(text)
                parts.append(f"<!-- página {i + 1} -->\n\n{text}".rstrip())
                view.advance()
    finally:
        doc.close()

    body = "\n\n".join(parts).strip()
    frontmatter = (
        "---\n"
        f"fuente: {pdf_path.name}\n"
        f"paginas: {total}\n"
        f"ocr: {ocr_mode} (ollama:{model})\n"
        "generado_por: pdf2md\n"
        "---\n\n"
        f"# {pdf_path.stem}\n\n"
    )
    return frontmatter + body + "\n"


# --------------------------------------------------------------------------- #
# Construcción del Markdown (imágenes)
# --------------------------------------------------------------------------- #
def default_images_stem(images: list[Path]) -> str:
    if len(images) == 1:
        return images[0].stem
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"ocr-imagenes-{stamp}"


def build_images_markdown(
    images: list[Path],
    model: str,
    host: str,
) -> str:
    """OCR de cada imagen con Ollama y ensambla un único Markdown."""
    if not check_ollama(host, model):
        raise RuntimeError(
            f"Ollama/modelo '{model}' no disponible en {host}. "
            "¿Está corriendo 'ollama serve' y el modelo descargado?"
        )

    total = len(images)
    parts: list[str] = []

    with TranscriptionDisplay(total, "OCR imágenes") as view:
        for i, img_path in enumerate(images, start=1):
            view.start_unit(f"OCR {i}/{total}: {img_path.name}")
            try:
                image_bytes = img_path.read_bytes()
                text = ocr_image_bytes_with_ollama(
                    image_bytes, model, host, on_chunk=view.feed
                )
            except requests.RequestException as exc:
                console.print(f"[red]Error de OCR ({img_path.name}): {exc}[/red]")
                text = ""
            except OSError as exc:
                console.print(f"[red]No se pudo leer {img_path.name}: {exc}[/red]")
                text = ""
            parts.append(
                f"<!-- imagen {i}: {img_path.name} -->\n\n"
                f"## {img_path.name}\n\n"
                f"{text}".rstrip()
            )
            view.advance()

    body = "\n\n".join(parts).strip()
    fuentes_yaml = "\n".join(f"  - {p.name}" for p in images)
    title = images[0].stem if len(images) == 1 else "Transcripción OCR de imágenes"
    frontmatter = (
        "---\n"
        "fuentes:\n"
        f"{fuentes_yaml}\n"
        f"imagenes: {total}\n"
        f"ocr: ollama:{model}\n"
        "generado_por: pdf2md\n"
        "---\n\n"
        f"# {title}\n\n"
    )
    return frontmatter + body + "\n"


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def normalize_argv(argv: list[str]) -> list[str]:
    """Inserta el subcomando si el usuario pasó solo una ruta o flags."""
    if not argv:
        return argv
    if argv[0] in ("pdf", "images", "img", "imgs", "-h", "--help"):
        if argv[0] in ("img", "imgs"):
            return ["images", *argv[1:]]
        return argv

    # Flags globales / de pdf sin subcomando → modo pdf (compatibilidad).
    if argv[0].startswith("-"):
        return ["pdf", *argv]

    first = Path(clean_dropped_path(argv[0]))
    if is_image_path(first):
        return ["images", *argv]
    return ["pdf", *argv]


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Transcribe un PDF o imágenes a Markdown o Word "
            "(OCR con Ollama / visión). Sin límite de páginas."
        ),
    )
    sub = parser.add_subparsers(dest="command")

    pdf_p = sub.add_parser("pdf", help="Transcribir un PDF a Markdown o Word.")
    pdf_p.add_argument("pdf", nargs="?", help="Ruta del PDF (si se omite, se pregunta).")
    pdf_p.add_argument(
        "-o", "--output",
        help=(
            "Ruta o carpeta de salida. Por defecto: "
            f"{default_output_dir()} ({default_output_format()})"
        ),
    )
    pdf_p.add_argument(
        "--format", choices=["md", "docx"], default=None,
        dest="out_format",
        help=(
            "Formato de salida. Windows: docx. macOS/Linux: md. "
            "También se infiere por la extensión de -o."
        ),
    )
    pdf_p.add_argument(
        "--ocr", choices=["auto", "always", "never"], default="auto",
        help="Cuándo aplicar OCR con Ollama. 'auto' = solo páginas escaneadas.",
    )
    pdf_p.add_argument(
        "--model", default=None,
        help=(
            "Modelo de visión en Ollama. Si se omite, se usa PDF2MD_VISION_MODEL "
            "o el mejor modelo instalado según la RAM (9b en 16 GB)."
        ),
    )
    pdf_p.add_argument(
        "--host", default=DEFAULT_OLLAMA_HOST,
        help=f"Host de Ollama. Por defecto: {DEFAULT_OLLAMA_HOST}",
    )

    img_p = sub.add_parser(
        "images",
        aliases=["img", "imgs"],
        help="OCR de una o más imágenes → Markdown o Word.",
    )
    img_p.add_argument(
        "images",
        nargs="*",
        help="Rutas de imágenes (también puedes arrastrarlas al prompt).",
    )
    img_p.add_argument(
        "-o", "--output",
        help=(
            "Ruta o carpeta de salida. Por defecto: "
            f"{default_output_dir()} ({default_output_format()})"
        ),
    )
    img_p.add_argument(
        "--format", choices=["md", "docx"], default=None,
        dest="out_format",
        help=(
            "Formato de salida. Windows: docx. macOS/Linux: md. "
            "También se infiere por la extensión de -o."
        ),
    )
    img_p.add_argument(
        "--model", default=None,
        help=(
            "Modelo de visión en Ollama. Si se omite, se usa PDF2MD_VISION_MODEL "
            "o el mejor modelo instalado según la RAM (9b en 16 GB)."
        ),
    )
    img_p.add_argument(
        "--host", default=DEFAULT_OLLAMA_HOST,
        help=f"Host de Ollama. Por defecto: {DEFAULT_OLLAMA_HOST}",
    )

    return parser.parse_args(argv)


def run_pdf(args: argparse.Namespace) -> int:
    pdf_path = Path(clean_dropped_path(args.pdf)) if args.pdf else prompt_for_pdf()
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        console.print(f"[red]Archivo PDF inválido:[/red] {pdf_path}")
        return 1

    fmt = resolve_output_format(getattr(args, "out_format", None), args.output)
    out_path = resolve_output_path(pdf_path.stem, args.output, fmt)

    try:
        scanned = detect_scanned_pages(pdf_path)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]No se pudo abrir el PDF:[/red] {exc}")
        return 1
    n_scanned = sum(scanned)

    console.print(f"[cyan]Entrada :[/cyan] {pdf_path}")
    console.print(f"[cyan]Salida  :[/cyan] {out_path}")
    console.print(
        f"[cyan]Páginas :[/cyan] {len(scanned)} "
        f"({n_scanned} escaneadas -> OCR; sin tope de n)"
    )
    console.print(f"[cyan]OCR     :[/cyan] {args.ocr} · ollama · {args.model}")
    console.print(f"[cyan]Formato :[/cyan] {fmt}")

    try:
        markdown = build_markdown(pdf_path, args.ocr, args.model, args.host, scanned)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]Error al procesar el PDF:[/red] {exc}")
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    write_output(markdown, out_path, fmt)

    kind = "Word" if fmt == "docx" else "Markdown"
    console.print(
        Panel.fit(
            f"[green]Listo[/green] -> {out_path}\n"
            f"{len(markdown):,} caracteres de transcripción ({kind}).",
            border_style="green",
        )
    )
    return 0


def run_images(args: argparse.Namespace) -> int:
    if args.images:
        # Permite tanto args separados como una sola línea con varias rutas.
        raw_joined = " ".join(args.images)
        # Si cada arg ya es una ruta limpia de argparse, preferimos eso.
        candidates = [Path(clean_dropped_path(p)) for p in args.images]
        # Si el usuario pegó todo en un solo token con escapes, re-parsear.
        if len(candidates) == 1 and ("\\" in args.images[0] or " " in args.images[0]):
            candidates = parse_dropped_paths(raw_joined)
        images, errors = validate_image_paths(candidates)
        for msg in errors:
            console.print(f"[red]{msg}[/red]")
        if not images:
            console.print("[red]No hay imágenes válidas para procesar.[/red]")
            return 1
    else:
        images = prompt_for_images()

    fmt = resolve_output_format(getattr(args, "out_format", None), args.output)
    out_path = resolve_output_path(default_images_stem(images), args.output, fmt)

    console.print(f"[cyan]Imágenes:[/cyan] {len(images)}")
    for img in images:
        console.print(f"  · {img}")
    console.print(f"[cyan]Salida  :[/cyan] {out_path}")
    console.print(f"[cyan]OCR     :[/cyan] ollama · {args.model}")
    console.print(f"[cyan]Formato :[/cyan] {fmt}")

    try:
        markdown = build_images_markdown(images, args.model, args.host)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]Error al procesar las imágenes:[/red] {exc}")
        return 1

    write_output(markdown, out_path, fmt)

    kind = "Word" if fmt == "docx" else "Markdown"
    console.print(
        Panel.fit(
            f"[green]Listo[/green] -> {out_path}\n"
            f"{len(images)} imagen(es) · {len(markdown):,} caracteres ({kind}).",
            border_style="green",
        )
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    raw = list(argv if argv is not None else sys.argv[1:])
    interactive = not raw
    if not raw:
        raw = [prompt_for_mode()]

    args = parse_args(normalize_argv(raw))
    if not args.command:
        args = parse_args([prompt_for_mode()])
        interactive = True

    if interactive and not getattr(args, "out_format", None) and os.name != "nt":
        args.out_format = Prompt.ask(
            "[cyan]Formato[/cyan]",
            choices=["md", "docx"],
            default="md",
        )

    args.model = resolve_vision_model(args.host, args.model)
    warn_if_model_too_heavy(args.model)

    if args.command == "pdf":
        return run_pdf(args)
    if args.command == "images":
        return run_images(args)

    console.print("[red]Comando no reconocido. Usa 'pdf' o 'images'.[/red]")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
