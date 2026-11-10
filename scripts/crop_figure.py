#!/usr/bin/env python3
"""Crop figure subpanels from a PDF page at high resolution.

Usage:
    python crop_figure.py <input.pdf> --page 5 \
        --panels "1a:10,15,45,60" "1b:50,15,85,60" \
        [--out figures] [--dpi 300] [--pad 1.5]

Panel spec format: "<label>:x0,y0,x1,y1"
    - label: output filename, e.g. "fig2a" -> fig2a.png
    - x0,y0,x1,y1: bounding box as PERCENTAGES (0-100) of page width/height,
      top-left origin (x0,y0 = top-left corner, x1,y1 = bottom-right corner)
    --pad: extra padding around the box, in percent of page size (default 1.5)

Why percentages: the agent reads the low-DPI page render to locate panels;
percent coords are DPI-independent, so they transfer cleanly to any render.

The crop is re-rendered from the PDF vector content at --dpi (default 300),
NOT cut from a screenshot, so panel text and lines stay sharp.

Dependencies: PyMuPDF (pymupdf). Install: pip install PyMuPDF
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    try:
        import fitz  # legacy import name
    except ImportError:
        sys.stderr.write("[ERROR] PyMuPDF is not installed. Run: pip install PyMuPDF\n")
        sys.exit(2)

PANEL_RE = re.compile(r"^(?P<label>[\w\-\.]+):(?P<x0>[\d.]+),(?P<y0>[\d.]+),(?P<x1>[\d.]+),(?P<y1>[\d.]+)$")


def safe_name(label: str) -> str:
    return re.sub(r"[^\w\-.]", "_", label)


def main():
    ap = argparse.ArgumentParser(description="Crop subpanels from a PDF page")
    ap.add_argument("pdf", type=Path, help="input PDF (e.g. the paper or its SI)")
    ap.add_argument("--page", type=int, required=True, help="1-based page number")
    ap.add_argument("--panels", nargs="+", required=True,
                    help='panel specs: "label:x0,y0,x1,y1" in percent, top-left origin')
    ap.add_argument("--out", type=Path, default=Path("figures"), help="output directory")
    ap.add_argument("--dpi", type=int, default=300, help="render DPI for the crop")
    ap.add_argument("--pad", type=float, default=1.5,
                    help="padding around box in percent of page size")
    args = ap.parse_args()

    if not args.pdf.exists():
        sys.stderr.write(f"[ERROR] PDF not found: {args.pdf}\n")
        sys.exit(2)

    args.out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(args.pdf)
    if not (1 <= args.page <= len(doc)):
        sys.stderr.write(f"[ERROR] page {args.page} out of range (1-{len(doc)})\n")
        sys.exit(2)
    page = doc[args.page - 1]
    pw, ph = page.rect.width, page.rect.height

    ok = 0
    for spec in args.panels:
        m = PANEL_RE.match(spec.strip())
        if not m:
            sys.stderr.write(f"[WARN] bad panel spec (need label:x0,y0,x1,y1): {spec}\n")
            continue
        g = {k: float(v) if k != "label" else v for k, v in m.groupdict().items()}
        if not (g["x0"] < g["x1"] and g["y0"] < g["y1"]):
            sys.stderr.write(f'[WARN] bad box (x0 must be < x1, y0 < y1): {spec}\n')
            continue
        # percent -> page points, with padding, clamped to page
        x0 = max(0, g["x0"] / 100 * pw - args.pad / 100 * pw)
        y0 = max(0, g["y0"] / 100 * ph - args.pad / 100 * ph)
        x1 = min(pw, g["x1"] / 100 * pw + args.pad / 100 * pw)
        y1 = min(ph, g["y1"] / 100 * ph + args.pad / 100 * ph)
        clip = fitz.Rect(x0, y0, x1, y1)
        zoom = args.dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
        out_file = args.out / f"{safe_name(g['label'])}.png"
        pix.save(out_file)
        print(f"[OK] {out_file}  (page {args.page}, box {g['x0']},{g['y0']},{g['x1']},{g['y1']}%, "
              f"{pix.width}x{pix.height}px)")
        ok += 1

    doc.close()
    print(f"\nDone. {ok} panel(s) cropped -> {args.out}")
    print("Next: view each crop to verify boundaries; re-run with adjusted coords if content "
          "is cut off or too much whitespace remains.")


if __name__ == "__main__":
    main()
