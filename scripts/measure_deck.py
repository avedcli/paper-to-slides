#!/usr/bin/env python3
"""Post-export text-fit audit for a generated deck.

Usage:
    python measure_deck.py output/deck.pptx [--font-dir C:/Windows/Fonts]
                                           [--font Calibri]

Why this exists
---------------
Stage 4 step 7 of the skill only counts slide parts / shapes. That catches a
dropped page, but NOT the two failure modes that actually ruin a deck:

  1. a text box whose wrapped content is taller than the box (text spills into
     whatever sits below it), and
  2. two text blocks whose *rendered* extents overlap each other.

python-pptx never computes wrapped height, and PowerPoint does not clip
overflow, so both defects are invisible to a structural check and invisible in
the PPTX XML. This script re-measures every text shape with the real font
metrics (greedy word wrap on the actual TTF) and with the same line-height
model PowerPoint uses (size * 1.22 * line_spacing), then reports:

  - per shape: box, required height, absolute bottom edge, flag if it overflows
  - per slide: pairwise overlaps between text blocks

Exit code is 1 when any issue is found, so it can gate delivery.

Notes on the model
------------------
* Sizes come from each run's ``sz`` (explicit) *before* falling back to the
  theme default — reading the max explicit size of the paragraph. A common bug
  in home-made checkers is seeding the running maximum with a default (e.g.
  18.0), which then silently wins over every smaller real size.
* Bold/italic runs use the matching TTF variant, since bold Calibri is ~4%
  wider and that is often exactly the margin that matters.
* ``marL`` (paragraph left margin, used for hanging bullets) is subtracted from
  the available width.
* The shape's text-frame insets (``margin_left/right/top/bottom``) are also
  subtracted — they are ~0.1/0.05 in by default and 0.14/0.08 in on a filled
  module block, so ignoring them under-estimates both the line count and the
  height. ``--ignore-insets`` restores the old (looser) model.
* A tolerance of 0.08 in is applied to the overflow test, and 0.05 in to the
  overlap test, to ignore rounding.
"""

import argparse
import sys
from pathlib import Path

from pptx import Presentation

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

VARIANTS = {
    ("", ""): ("calibri.ttf",),
    ("b", ""): ("calibrib.ttf", "calibri.ttf"),
    ("", "i"): ("calibrii.ttf", "calibri.ttf"),
    ("b", "i"): ("calibriz.ttf", "calibri.ttf"),
}

_cache = {}


def load_font(font_dir, family, size_pt, bold, italic):
    from PIL import ImageFont
    key = (family, round(size_pt, 1), bold, italic)
    if key in _cache:
        return _cache[key]
    px = max(1, int(round(size_pt * 96 / 72)))
    stem = family.replace(" ", "") + ("b" if bold else "") + ("i" if italic else "")
    cands = [f"{stem}.ttf", f"{stem}.otf", "calibri.ttf"]
    font = None
    for cand in cands:
        p = Path(font_dir) / cand
        if p.exists():
            font = ImageFont.truetype(str(p), px)
            break
    if font is None:                       # last resort: Pillow's bitmap font
        font = ImageFont.load_default()
    _cache[key] = font
    return font


def text_width_pt(font, text):
    return font.getlength(text) * 72 / 96   # px @96dpi -> pt


def wrap_lines(chunks, avail_pt, font):
    """chunks: strings already split on explicit line breaks."""
    lines = 0
    for text in chunks:
        if not text:
            lines += 1
            continue
        n, cur = 1, ""
        for word in text.split(" "):
            trial = word if not cur else cur + " " + word
            if text_width_pt(font, trial) <= avail_pt or not cur:
                cur = trial
            else:
                n += 1
                cur = word
        lines += n
    return lines


def marL_in(pPr):
    if pPr is None:
        return 0.0
    return int(pPr.get("marL") or 0) / 914400


def paragraphs(tf):
    """-> [(chunks, size_pt, bold, italic, line_spacing, space_after, space_before, pPr)]"""
    out = []
    for p in tf.paragraphs:
        chunks, cur, size, bold, italic = [], "", 0.0, False, False
        for child in p._p:
            tag = child.tag.replace(A, "")
            if tag == "br":
                chunks.append(cur)
                cur = ""
            elif tag == "r":
                rPr = child.find(f"{A}rPr")
                sz, b, i = 0.0, False, False
                if rPr is not None:
                    if rPr.get("sz"):
                        sz = int(rPr.get("sz")) / 100
                    b = rPr.get("b") == "1"
                    i = rPr.get("i") == "1"
                size = max(size, sz)
                bold = bold or b
                italic = italic or i
                cur += "".join(e.text or "" for e in child.findall(f"{A}t"))
        if cur:
            chunks.append(cur)
        if not chunks:
            continue
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        sa = p.space_after.pt if p.space_after else 0.0
        sb = p.space_before.pt if p.space_before else 0.0
        out.append((chunks, size or 18.0, bold, italic, ls, sa, sb,
                    p._p.find(f"{A}pPr")))
    return out


def audit(path, font_dir, family, overflow_tol=0.08, overlap_tol=0.05,
          insets=True):
    prs = Presentation(str(path))
    issues = 0
    for si, slide in enumerate(prs.slides, 1):
        print(f"--- slide {si}")
        blocks = []
        for sh in slide.shapes:
            x, y = sh.left / 914400, sh.top / 914400
            w, h = sh.width / 914400, sh.height / 914400
            has_text = sh.has_text_frame and sh.text_frame.text.strip()
            if not has_text:
                blocks.append((sh.name, x, y, w, h, None, None))
                continue
            need_pt = 0.0
            tf = sh.text_frame
            # Text-frame insets are real space: a "module block" (filled shape
            # with reverse-white text, see typography-spec.md) carries 0.14 in
            # left/right and 0.08 in top/bottom padding. Ignoring them under-
            # estimates both the needed width and height, i.e. it hides exactly
            # the overflow this script exists to catch.
            ins_l = (tf.margin_left or 0) / 914400 if insets else 0
            ins_r = (tf.margin_right or 0) / 914400 if insets else 0
            ins_t = (tf.margin_top or 0) / 914400 if insets else 0
            ins_b = (tf.margin_bottom or 0) / 914400 if insets else 0
            inner_w = max(w - ins_l - ins_r, 0.01)
            for chunks, size, bold, italic, ls, sa, sb, pPr in paragraphs(sh.text_frame):
                font = load_font(font_dir, family, size, bold, italic)
                avail = inner_w * 72 - marL_in(pPr) * 72
                n = wrap_lines(chunks, avail, font)
                need_pt += sb + n * size * 1.22 * ls + sa
            need = need_pt / 72 + ins_t + ins_b
            note = ""
            if need > h + overflow_tol:
                note = f"  <== text needs {need:.2f} in, box is {h:.2f} in"
                issues += 1
            print(f"   {sh.name[:20]:22s} x={x:5.2f} y={y:5.2f} w={w:5.2f} "
                  f"h={h:4.2f} need={need:5.2f} bottom={y + need:5.2f}{note}")
            blocks.append((sh.name, x, y, w, h, need, sh))
        for i in range(len(blocks)):
            for j in range(i + 1, len(blocks)):
                a, b = blocks[i], blocks[j]
                if a[5] is None or b[5] is None or a[6] is b[6]:
                    continue
                vo = min(a[2] + a[5], b[2] + b[5]) - max(a[2], b[2])
                ho = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
                if vo > overlap_tol and ho > overlap_tol:
                    print(f"   !! OVERLAP {a[0][:16]} vs {b[0][:16]} "
                          f"(v={vo:.2f} h={ho:.2f})")
                    issues += 1
    print(f"\nissues: {issues}")
    return issues


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--font-dir", default=r"C:/Windows/Fonts",
                    help="directory holding the TTF used by the deck's theme")
    ap.add_argument("--font", default="Calibri", help="latin typeface family name")
    ap.add_argument("--ignore-insets", action="store_true",
                    help="ignore text-frame margins (legacy behaviour; less accurate "
                         "for filled module blocks, whose 0.14/0.08 in padding is real)")
    args = ap.parse_args()
    sys.exit(1 if audit(args.deck, args.font_dir, args.font,
                        insets=not args.ignore_insets) else 0)


if __name__ == "__main__":
    main()
