#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["typer>=0.12,<1", "pillow>=10,<13"]
# ///
"""Export a .pptx to PDF, one slide per page, via LibreOffice — the whole deck or
a page range — with an optional size cap that rasterizes the PDF (poppler +
Pillow) when LibreOffice's font-embedded output is too heavy to share.

    uv run export_pdf.py deck.pptx                          # -> deck.pdf beside it
    uv run export_pdf.py deck.pptx --pages 4-10 --out section.pdf
    uv run export_pdf.py deck.pptx --max-mb 1               # rasterize if over 1 MB

Needs `soffice` (LibreOffice); page ranges and rasterizing also need poppler
(`pdfseparate`, `pdfunite`, `pdftoppm`).
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import typer
    from PIL import Image
except ImportError as exc:  # a bare `python3` without the deps installed
    sys.exit(
        f"[pdf] missing dependency: {exc.name}. Run via `uv run {Path(__file__).name}` "
        "(deps are declared inline) or `pip install typer pillow`."
    )

_MAC_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def _tool(name: str, *alternatives: str) -> str:
    for candidate in (name, *alternatives):
        found = shutil.which(candidate)
        if found:
            return found
    if name == "soffice" and Path(_MAC_SOFFICE).exists():
        return _MAC_SOFFICE
    typer.echo(f"[pdf] {name} not found on PATH", err=True)
    raise typer.Exit(2)


def _soffice_pdf(pptx: Path, work: Path) -> Path:
    """LibreOffice's Impress export is already one slide per page. A private
    profile dir avoids clashing with an interactive LibreOffice the user has open."""
    soffice = _tool("soffice", "libreoffice")
    profile = work / "_lo_profile"
    subprocess.run(
        [soffice, "--headless", f"-env:UserInstallation=file://{profile}",
         "--convert-to", "pdf", "--outdir", str(work), str(pptx)],
        check=True, capture_output=True, timeout=300,
    )
    pdf = work / (pptx.stem + ".pdf")
    if not pdf.is_file():
        typer.echo("[pdf] LibreOffice produced no PDF", err=True)
        raise typer.Exit(1)
    return pdf


def parse_pages(spec: str) -> list[int]:
    """'4-10,12' -> [4, 5, ..., 10, 12]; 1-indexed, matching slide order."""
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = (int(p) for p in part.split("-", 1))
            pages.extend(range(lo, hi + 1))
        elif part:
            pages.append(int(part))
    if not pages or min(pages) < 1:
        raise typer.BadParameter(f"bad page spec: {spec!r}")
    return pages


def _slice_pages(pdf: Path, pages: list[int], work: Path) -> Path:
    pdfseparate, pdfunite = _tool("pdfseparate"), _tool("pdfunite")
    parts: list[str] = []
    for page in pages:
        part = work / f"pg-{page}.pdf"
        subprocess.run(
            [pdfseparate, "-f", str(page), "-l", str(page), str(pdf), str(part)],
            check=True, capture_output=True,
        )
        parts.append(str(part))
    out = work / "sliced.pdf"
    subprocess.run([pdfunite, *parts, str(out)], check=True, capture_output=True)
    return out


def _rasterize(pdf: Path, work: Path, *, dpi: int, quality: int) -> Path:
    """pdftoppm -> JPEG pages -> one PDF via Pillow. 150 DPI stays legible."""
    pdftoppm = _tool("pdftoppm")
    prefix = work / "raster"
    subprocess.run(
        [pdftoppm, "-jpeg", "-r", str(dpi), "-jpegopt", f"quality={quality}",
         str(pdf), str(prefix)],
        check=True, capture_output=True,
    )
    pages = sorted(work.glob("raster-*.jpg"))
    images = [Image.open(p).convert("RGB") for p in pages]
    out = work / "raster.pdf"
    images[0].save(out, "PDF", save_all=True, append_images=images[1:], resolution=dpi)
    return out


def main(
    pptx: Path = typer.Argument(..., help="the .pptx to export"),
    out: Path = typer.Option(None, "--out", "-o", help="output .pdf (default: beside the pptx)"),
    pages: str = typer.Option(None, "--pages", help="1-indexed page range, e.g. '4-10' or '2,5-7'"),
    max_mb: float = typer.Option(None, "--max-mb", help="rasterize if the PDF is larger than this"),
    dpi: int = typer.Option(150, help="rasterize resolution"),
    quality: int = typer.Option(85, help="rasterize JPEG quality"),
) -> None:
    """Export a pptx to a one-slide-per-page PDF (whole deck or a page range)."""
    if not pptx.is_file():
        typer.echo(f"[pdf] not found: {pptx}", err=True)
        raise typer.Exit(1)
    out_path = out or pptx.with_suffix(".pdf")
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        pdf = _soffice_pdf(pptx, work)
        if pages:
            pdf = _slice_pages(pdf, parse_pages(pages), work)
        size_mb = pdf.stat().st_size / 1e6
        if max_mb is not None and size_mb > max_mb:
            pdf = _rasterize(pdf, work, dpi=dpi, quality=quality)
            typer.echo(f"[pdf] {size_mb:.1f} MB > {max_mb} MB cap — rasterized at {dpi} DPI")
        shutil.copyfile(pdf, out_path)
    typer.echo(f"[pdf] wrote {out_path} ({out_path.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    typer.run(main)
