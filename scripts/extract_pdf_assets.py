#!/usr/bin/env python3
"""Extract reusable figure assets from a literature PDF.

Usage:
    python extract_pdf_assets.py <input.pdf> [more.pdf ...] --out <output_dir> [--dpi 200]
        [--min-size 200] [--no-pages]

What it does:
    1. Extracts embedded raster images from every page (filters out tiny
       icons / decorative elements below --min-size pixels).
    2. Renders every full page as a PNG at the requested DPI (useful for
       vector figures and for cropping subpanels manually afterwards).
    3. Writes a manifest.json mapping each output file to its source page
       and size, so slides can reference figures deterministically.

Dependencies:
    PyMuPDF (fitz). Install into the managed venv:
        pip install PyMuPDF
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    try:
        import fitz  # legacy import name
    except ImportError:
        sys.stderr.write(
            "[ERROR] PyMuPDF is not installed. Run:\n"
            "  pip install PyMuPDF\n"
            "then retry.\n"
        )
        sys.exit(2)


def pdf_slug(path: Path) -> str:
    """Short, filesystem-safe prefix derived from the PDF filename."""
    return path.stem[:30].replace(" ", "_")


def extract_pdf(pdf_path: Path, out_dir: Path, dpi: int, min_size: int, render_pages: bool):
    entries = []
    prefix = pdf_slug(pdf_path)
    doc = fitz.open(pdf_path)

    for page_index, page in enumerate(doc, start=1):
        # --- embedded raster images ---
        seen_xrefs = set()
        for img_index, img in enumerate(page.get_images(full=True), start=1):
            xref = img[0]
            if xref in seen_xrefs:  # same image placed multiple times
                continue
            seen_xrefs.add(xref)
            try:
                pix = fitz.Pixmap(doc, xref)
            except Exception:
                continue
            if pix.width < min_size or pix.height < min_size:
                continue  # skip icons, bullets, decorations
            if pix.colorspace and pix.colorspace.n > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            name = f"{prefix}_p{page_index:03d}_img{img_index:02d}.png"
            pix.save(out_dir / name)
            entries.append({
                "file": name,
                "source": pdf_path.name,
                "page": page_index,
                "width": pix.width,
                "height": pix.height,
                "kind": "embedded-image",
            })
            pix = None

        # --- full page render (captures vector figures) ---
        if render_pages:
            zoom = dpi / 72.0
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            name = f"{prefix}_page{page_index:03d}.png"
            pix.save(out_dir / name)
            entries.append({
                "file": name,
                "source": pdf_path.name,
                "page": page_index,
                "width": pix.width,
                "height": pix.height,
                "kind": "page-render",
                "dpi": dpi,
            })

    doc.close()
    return entries


def main():
    ap = argparse.ArgumentParser(description="Extract figure assets from PDF(s)")
    ap.add_argument("pdfs", nargs="+", type=Path, help="input PDF file(s), e.g. main paper + SI")
    ap.add_argument("--out", type=Path, default=Path("figures"), help="output directory")
    ap.add_argument("--dpi", type=int, default=200, help="DPI for full-page renders")
    ap.add_argument("--min-size", type=int, default=200,
                    help="skip embedded images smaller than this (pixels), filters icons")
    ap.add_argument("--no-pages", action="store_true", help="skip full-page rendering")
    args = ap.parse_args()

    args.out.mkdir(parents=True, exist_ok=True)

    all_entries = []
    for pdf in args.pdfs:
        if not pdf.exists():
            sys.stderr.write(f"[WARN] not found, skipped: {pdf}\n")
            continue
        entries = extract_pdf(pdf, args.out, args.dpi, args.min_size,
                              render_pages=not args.no_pages)
        print(f"[OK] {pdf.name}: {len(entries)} assets")
        all_entries.extend(entries)

    manifest = args.out / "manifest.json"
    manifest.write_text(json.dumps(all_entries, indent=2, ensure_ascii=False),
                        encoding="utf-8")
    embedded = sum(1 for e in all_entries if e["kind"] == "embedded-image")
    pages = sum(1 for e in all_entries if e["kind"] == "page-render")
    print(f"\nDone. {embedded} embedded images, {pages} page renders -> {args.out}")
    print(f"Manifest: {manifest}")
    print("Next: view page renders, crop subpanels (a/b/c) from them, and map each "
          "subpanel to a slide via the figure inventory.")


if __name__ == "__main__":
    main()
