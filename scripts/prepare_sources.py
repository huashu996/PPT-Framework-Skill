#!/usr/bin/env python3
"""Render PDFs and normalize image inputs into traceable PNG page images."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageOps

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}


def render_image(source: Path, destination: Path) -> tuple[int, int]:
    with Image.open(source) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        image.save(destination, "PNG")
        return image.size


def render_pdf(source: Path, pages_dir: Path, dpi: int, start_index: int):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return render_pdf_with_poppler(source, pages_dir, dpi, start_index)

    document = fitz.open(source)
    records = []
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    for page_number, page in enumerate(document, start=1):
        output_name = f"{start_index:04d}-{source.stem}-page-{page_number:03d}.png"
        output_path = pages_dir / output_name
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        pixmap.save(output_path)
        records.append(
            {
                "source": str(source.resolve()),
                "kind": "pdf_page",
                "page": page_number,
                "image": str(output_path.resolve()),
                "width_px": pixmap.width,
                "height_px": pixmap.height,
                "dpi": dpi,
            }
        )
        start_index += 1
    document.close()
    return records, start_index


def render_pdf_with_poppler(source: Path, pages_dir: Path, dpi: int, start_index: int):
    """Fallback renderer for bundled runtimes that expose pdftoppm but not PyMuPDF."""
    renderer = shutil.which("pdftoppm.exe") or shutil.which("pdftoppm")
    if not renderer:
        raise RuntimeError("PDF rendering requires PyMuPDF or pdftoppm (Poppler)")

    records = []
    with tempfile.TemporaryDirectory(prefix="paper-fig-formula-pdf-") as temp_dir:
        prefix = Path(temp_dir) / "page"
        arguments = [renderer, "-png", "-r", str(dpi), str(source), str(prefix)]
        if Path(renderer).suffix.lower() in {".cmd", ".bat"}:
            command_shell = os.environ.get("COMSPEC", "cmd.exe")
            arguments = [command_shell, "/d", "/c", "call", *arguments]
        subprocess.run(arguments, check=True, capture_output=True, text=True)

        rendered_pages = sorted(
            Path(temp_dir).glob("page-*.png"),
            key=lambda path: int(path.stem.rsplit("-", 1)[1]),
        )
        if not rendered_pages:
            raise RuntimeError(f"pdftoppm produced no pages for {source}")

        for page_number, rendered in enumerate(rendered_pages, start=1):
            output_name = f"{start_index:04d}-{source.stem}-page-{page_number:03d}.png"
            output_path = pages_dir / output_name
            shutil.move(str(rendered), output_path)
            with Image.open(output_path) as image:
                width, height = image.size
            records.append(
                {
                    "source": str(source.resolve()),
                    "kind": "pdf_page",
                    "page": page_number,
                    "image": str(output_path.resolve()),
                    "width_px": width,
                    "height_px": height,
                    "dpi": dpi,
                }
            )
            start_index += 1
    return records, start_index


def collect_sources(input_path: Path):
    if input_path.is_file():
        return [input_path]
    return sorted(
        path
        for path in input_path.rglob("*")
        if path.is_file() and (path.suffix.lower() in IMAGE_EXTS or path.suffix.lower() == ".pdf")
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=240)
    args = parser.parse_args()

    source_path = args.input.expanduser().resolve()
    if not source_path.exists():
        raise SystemExit(f"Input does not exist: {source_path}")
    if args.dpi < 120 or args.dpi > 600:
        raise SystemExit("--dpi must be between 120 and 600")

    output_dir = args.output.expanduser().resolve()
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    records = []
    index = 1
    for source in collect_sources(source_path):
        suffix = source.suffix.lower()
        if suffix == ".pdf":
            pdf_records, index = render_pdf(source, pages_dir, args.dpi, index)
            records.extend(pdf_records)
        elif suffix in IMAGE_EXTS:
            output_name = f"{index:04d}-{source.stem}.png"
            output_path = pages_dir / output_name
            width, height = render_image(source, output_path)
            records.append(
                {
                    "source": str(source.resolve()),
                    "kind": "image",
                    "page": None,
                    "image": str(output_path.resolve()),
                    "width_px": width,
                    "height_px": height,
                    "dpi": None,
                }
            )
            index += 1

    manifest = {
        "input": str(source_path),
        "output": str(output_dir),
        "page_count": len(records),
        "pages": records,
    }
    manifest_path = output_dir / "source-manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(manifest_path)


if __name__ == "__main__":
    main()
