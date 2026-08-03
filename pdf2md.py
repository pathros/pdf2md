#!/usr/bin/env python3
"""pdf2md - Convierte PDF o imágenes a Markdown digerible para LLMs.

Modos:
  pdf     Extrae texto/tablas con pymupdf4llm. Páginas escaneadas → OCR con
          Ollama (visión, por defecto qwen3.6:27b).
  images  Transcribe N imágenes con el mismo OCR de visión y las junta en
          un solo .md.

Uso: ejecuta el comando, elige modo y ARRASTRA el/los archivo(s) a la
terminal. El .md se guarda en ~/Downloads por defecto.
"""

from __future__ import annotations

import argparse
import base64
import os
import shlex
import sys
from datetime import datetime
from pathlib import Path

import requests
from rich.console import Console
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

import pymupdf  # PyMuPDF
import pymupdf4llm

console = Console()

# --- Configuración por defecto (sobreescribible por CLI o variables de entorno) ---
DEFAULT_OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
DEFAULT_VISION_MODEL = os.environ.get("PDF2MD_VISION_MODEL", "qwen3.6:27b")
DEFAULT_OUTPUT_DIR = Path(os.environ.get("PDF2MD_OUTPUT_DIR", str(Path.home() / "Downloads")))

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
    "Markdown limpio y bien estructurado, en español (respeta el idioma original). "
    "Usa encabezados (#), listas y tablas de Markdown cuando corresponda y conserva "
    "el orden de lectura. NO traduzcas, NO resumas y NO agregues comentarios ni "
    "explicaciones: devuelve únicamente la transcripción en Markdown."
)

PROGRESS_COLUMNS = (
    SpinnerColumn(),
    TextColumn("[progress.description]{task.description}"),
    BarColumn(),
    TextColumn("{task.completed}/{task.total}"),
    TimeElapsedColumn(),
    TimeRemainingColumn(),
)


# --------------------------------------------------------------------------- #
# Utilidades de ruta / entrada
# --------------------------------------------------------------------------- #
def clean_dropped_path(raw: str) -> str:
    """Normaliza una ruta que el usuario pudo arrastrar a la terminal.

    Al arrastrar un archivo, la terminal suele rodearlo de comillas o escapar
    los espacios con '\\ '. Aquí lo dejamos como una ruta usable.
    """
    s = raw.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
        s = s[1:-1]
    for esc in ("\\ ", "\\(", "\\)", "\\&", "\\'", "\\\""):
        s = s.replace(esc, esc[1])
    return os.path.expanduser(s.strip())


def parse_dropped_paths(raw: str) -> list[Path]:
    """Separa varias rutas arrastradas desde Finder (espacios / escapes / comillas)."""
    raw = raw.strip()
    if not raw:
        return []
    try:
        tokens = shlex.split(raw, posix=True)
    except ValueError:
        tokens = raw.split()
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
            "[bold]1[/bold]  PDF → Markdown\n"
            "[bold]2[/bold]  Imágenes → Markdown (OCR con Ollama)",
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
    """Pide una o más rutas de imagen (arrastrar desde Finder, separadas por espacio)."""
    console.print(
        Panel.fit(
            "[bold]Arrastra una o más imágenes a esta terminal[/bold] y pulsa Enter.\n"
            "Las rutas quedan separadas por espacio (Finder lo hace solo).\n"
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
# OCR con Ollama (modelo de visión)
# --------------------------------------------------------------------------- #
def render_page_png(page: "pymupdf.Page", zoom: float = OCR_RENDER_ZOOM) -> bytes:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
    return pix.tobytes("png")


def check_ollama(host: str, model: str) -> bool:
    """Verifica que Ollama responda y que el modelo esté disponible."""
    try:
        resp = requests.get(f"{host}/api/tags", timeout=5)
        resp.raise_for_status()
    except requests.RequestException:
        return False
    names = {m.get("name", "") for m in resp.json().get("models", [])}
    base = model.split(":")[0]
    return model in names or any(n.split(":")[0] == base for n in names)


def ocr_image_bytes_with_ollama(image_bytes: bytes, model: str, host: str) -> str:
    payload = {
        "model": model,
        "prompt": OCR_PROMPT,
        "images": [base64.b64encode(image_bytes).decode("ascii")],
        "stream": False,
        "options": {"temperature": 0},
    }
    resp = requests.post(f"{host}/api/generate", json=payload, timeout=900)
    resp.raise_for_status()
    return resp.json().get("response", "").strip()


def ocr_page_with_ollama(png_bytes: bytes, model: str, host: str) -> str:
    return ocr_image_bytes_with_ollama(png_bytes, model, host)


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
    escaneadas (o todas, si ocr_mode='always') se transcriben con Ollama.
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
        # use_ocr=False -> pymupdf4llm NO llama a Tesseract; solo extrae el texto
        # existente. Las páginas escaneadas salen vacías y las cubrimos con Ollama.
        chunks = pymupdf4llm.to_markdown(
            doc, page_chunks=True, use_ocr=False, show_progress=False
        )
        with Progress(*PROGRESS_COLUMNS, console=console) as progress:
            task = progress.add_task("Procesando", total=total)
            for i in range(total):
                do_ocr = ocr_mode == "always" or (ocr_mode == "auto" and scanned[i])
                if do_ocr and ocr_available:
                    progress.update(task, description=f"OCR (visión) pág. {i + 1}/{total}")
                    try:
                        png = render_page_png(doc.load_page(i))
                        text = ocr_page_with_ollama(png, model, host)
                    except requests.RequestException as exc:
                        console.print(f"[red]Error de OCR pág. {i + 1}: {exc}[/red]")
                        text = ""
                else:
                    progress.update(task, description=f"Texto pág. {i + 1}/{total}")
                    text = (chunks[i].get("text", "") if i < len(chunks) else "").strip()
                parts.append(f"<!-- página {i + 1} -->\n\n{text}".rstrip())
                progress.advance(task)
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


def resolve_output_path(source_stem: str, output: str | None) -> Path:
    if output:
        out = Path(os.path.expanduser(output))
        if out.is_dir() or output.endswith(os.sep):
            return out / f"{source_stem}.md"
        if out.suffix.lower() != ".md":
            out = out.with_suffix(".md")
        return out
    DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return DEFAULT_OUTPUT_DIR / f"{source_stem}.md"


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

    with Progress(*PROGRESS_COLUMNS, console=console) as progress:
        task = progress.add_task("OCR imágenes", total=total)
        for i, img_path in enumerate(images, start=1):
            progress.update(
                task,
                description=f"OCR {i}/{total}: {img_path.name}",
            )
            try:
                image_bytes = img_path.read_bytes()
                text = ocr_image_bytes_with_ollama(image_bytes, model, host)
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
            progress.advance(task)

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
            "Convierte un PDF o varias imágenes a Markdown "
            "(OCR con Ollama / visión)."
        ),
    )
    sub = parser.add_subparsers(dest="command")

    pdf_p = sub.add_parser("pdf", help="Convertir un PDF a Markdown.")
    pdf_p.add_argument("pdf", nargs="?", help="Ruta del PDF (si se omite, se pregunta).")
    pdf_p.add_argument(
        "-o", "--output",
        help=f"Ruta o carpeta de salida (.md). Por defecto: {DEFAULT_OUTPUT_DIR}",
    )
    pdf_p.add_argument(
        "--ocr", choices=["auto", "always", "never"], default="auto",
        help="Cuándo aplicar OCR con Ollama. 'auto' = solo páginas escaneadas.",
    )
    pdf_p.add_argument(
        "--model", default=DEFAULT_VISION_MODEL,
        help=f"Modelo de visión en Ollama. Por defecto: {DEFAULT_VISION_MODEL}",
    )
    pdf_p.add_argument(
        "--host", default=DEFAULT_OLLAMA_HOST,
        help=f"Host de Ollama. Por defecto: {DEFAULT_OLLAMA_HOST}",
    )

    img_p = sub.add_parser(
        "images",
        aliases=["img", "imgs"],
        help="OCR de una o más imágenes → un solo Markdown.",
    )
    img_p.add_argument(
        "images",
        nargs="*",
        help="Rutas de imágenes (también puedes arrastrarlas al prompt).",
    )
    img_p.add_argument(
        "-o", "--output",
        help=f"Ruta o carpeta de salida (.md). Por defecto: {DEFAULT_OUTPUT_DIR}",
    )
    img_p.add_argument(
        "--model", default=DEFAULT_VISION_MODEL,
        help=f"Modelo de visión en Ollama. Por defecto: {DEFAULT_VISION_MODEL}",
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

    out_path = resolve_output_path(pdf_path.stem, args.output)

    try:
        scanned = detect_scanned_pages(pdf_path)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]No se pudo abrir el PDF:[/red] {exc}")
        return 1
    n_scanned = sum(scanned)

    console.print(f"[cyan]Entrada :[/cyan] {pdf_path}")
    console.print(f"[cyan]Salida  :[/cyan] {out_path}")
    console.print(
        f"[cyan]Páginas :[/cyan] {len(scanned)} ({n_scanned} escaneadas -> OCR)"
    )
    console.print(f"[cyan]OCR     :[/cyan] {args.ocr} · ollama · {args.model}")

    try:
        markdown = build_markdown(pdf_path, args.ocr, args.model, args.host, scanned)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]Error al procesar el PDF:[/red] {exc}")
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(markdown, encoding="utf-8")

    console.print(
        Panel.fit(
            f"[green]Listo[/green] -> {out_path}\n"
            f"{len(markdown):,} caracteres de Markdown.",
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

    out_path = resolve_output_path(default_images_stem(images), args.output)

    console.print(f"[cyan]Imágenes:[/cyan] {len(images)}")
    for img in images:
        console.print(f"  · {img}")
    console.print(f"[cyan]Salida  :[/cyan] {out_path}")
    console.print(f"[cyan]OCR     :[/cyan] ollama · {args.model}")

    try:
        markdown = build_images_markdown(images, args.model, args.host)
    except Exception as exc:  # noqa: BLE001
        console.print(f"[red]Error al procesar las imágenes:[/red] {exc}")
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(markdown, encoding="utf-8")

    console.print(
        Panel.fit(
            f"[green]Listo[/green] -> {out_path}\n"
            f"{len(images)} imagen(es) · {len(markdown):,} caracteres de Markdown.",
            border_style="green",
        )
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    raw = list(argv if argv is not None else sys.argv[1:])
    if not raw:
        raw = [prompt_for_mode()]

    args = parse_args(normalize_argv(raw))
    if not args.command:
        args = parse_args([prompt_for_mode()])

    if args.command == "pdf":
        return run_pdf(args)
    if args.command == "images":
        return run_images(args)

    console.print("[red]Comando no reconocido. Usa 'pdf' o 'images'.[/red]")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
