#!/usr/bin/env python3
"""Plan subpanel layout on a presentation slide.

Deterministic layout planner used by the paper-to-slides skill. Given a set of
cropped subpanel images (with per-panel estimates of the smallest text height
in source pixels), it enumerates candidate grid layouts inside the slide's
content area, scores them, and outputs exact placement rectangles plus
warnings. PPT generation then applies the placements verbatim.

Usage:
    python plan_layout.py plan.json [-o layout.json]

plan.json schema (all size fields in inches; slide assumed 96 DPI for the
text-size check, so the 14 pt content floor = 14/72 in = 18.67 px):

    {
      "slide_w_in": 13.333, "slide_h_in": 7.5,
      "title_band_in": 0.85,      // reserved title strip on top
      "caption_band_in": 0,       // caption strip at bottom (L4 = 14pt
                                  //   cap-band). Two kinds of rows live here
                                  //   (SKILL.md 3.6): A) name annotations
                                  //   (`Lgr5: marker`) — written whenever the
                                  //   page introduces a term, so a one-line
                                  //   band is a NORMAL page, not a surplus;
                                  //   B) supplementary rows (bottleneck /
                                  //   reagent principle) — rare, only when
                                  //   omitting them makes the audience misread
                                  //   the figure. 0 = no band at all, figure
                                  //   area takes the space back (pages whose
                                  //   terms were all established earlier).
                                  //   SIZE IT BY LINE COUNT, never by taste:
                                  //   one 14 pt line = 0.237 in + 0.18 in
                                  //   padding -> 1 line 0.42 / 2 lines 0.66 /
                                  //   3 lines 0.90 (SKILL.md 3.6). Under-reserving
                                  //   a line is a silent overflow: PowerPoint does
                                  //   not clip, the extra line lands on the
                                  //   figures or the footer.
      "margin_in": 0.3,
      "gap_in": 0.15,             // min gap between panels (no-overlap rule)
                                  // content area = slide minus bands minus
                                  // margins; that area is what "usable figure
                                  // area" means for the coverage metric below
      "target_text_px": 20,       // comfort target for on-slide figure text.
                                  //   Ignored when min_text_px <= 0.
      "min_text_px": 18.67,       // comfort floor = the skill-wide 14 pt
                                  //   content minimum (L4 in typography-spec.md),
                                  //   expressed in px at 96 dpi. Below it the
                                  //   number is reported as a note only —
                                  //   never blocks, never escalates.
                                  //   PASS 0 IN THIS SKILL: it disables the whole
                                  //   text-size term (floor AND target), leaving
                                  //   only "no overlap + area + whitespace" as
                                  //   the objective. On-slide text is a step-4
                                  //   record and must not steer layout choice.
      "panel_caption_in": 0.30,   // OPTIONAL. Height reserved UNDER EACH panel
                                  //   for its own "Fig. 5a  <one-line title>"
                                  //   row (14 pt). 1 line ~= 0.30, 2 lines ~= 0.53.
                                  //   Omitted / 0 -> legacy behaviour (no reserve).
                                  //   Alternatively give per-panel
                                  //   "caption_lines": 2 and let the script derive
                                  //   it from caption_gap_in + lines*caption_line_in.
      "caption_gap_in": 0.06,     // OPTIONAL. Gap between a panel's bottom edge
      "caption_line_in": 0.237,   // OPTIONAL. One 14 pt line = 14*1.22/72 in
      "body_zone": {              // OPTIONAL. The page's BODY text zone (bullets
        "pos": "left",            //   / module blocks) as a first-class module,
        "in": 3.0                 //   distinct from the caption band pinned to
      },                          //   the page bottom (still reserved by
                                  //   caption_band_in).
                                  //   pos: "top" | "bottom" | "left" | "right";
                                  //   "in" = height (top/bottom) or width
                                  //   (left/right) in inches. Omitted -> the
                                  //   generator falls back to the legacy
                                  //   convention (body strip right above the
                                  //   caption band). Output gains "body_box"
                                  //   — place bullets / module blocks there,
                                  //   verbatim, laid out freely inside it.
                                  //   Legacy alias: "text_col_in" (+
                                  //   "text_col_side") still works and maps
                                  //   to pos left/right.
      "structure": "hierarchy",   // OPTIONAL. "hierarchy" -> big-cell layouts
                                  //   (1+2 / 1+3 / ...) get a -4 bonus: one
                                  //   key panel should dominate.
                                  //   "parallel" -> uniform grids (2x2 etc.)
                                  //   get a -4 bonus: panels are equals.
                                  //   Omitted -> purely geometric scoring.
      "panels": [                 // order = reading order; first panel gets the
        {                         //   "big" cell in 1+2 / 2+1 layouts
          "label": "fig2a",
          "file": "figures/fig2a.png",
          "text_px": 9,           // agent's estimate of smallest text height
          "group": "fig2",        // same original figure -> forced equal scale
          "role": "key",          // OPTIONAL. "key" = this page's decisive
                                  //   evidence (at most ONE per page). In
                                  //   big-cell candidates the key panel is
                                  //   moved into the big cell regardless of
                                  //   list order. Omit for parallel evidence.
          "caption_lines": 1      // OPTIONAL, 1 or 2 (never 3 — the title is
        }                         //   too long, shorten it instead)
      ]
    }

Output layout.json:
    {
      "ok": true|false,           // false only when a panel cannot be placed
                                  //   at all (non-positive extent); text size
                                  //   alone never makes this false
      "layout": "2x2",
      "placements": [ {"label","x","y","w","h","text_display_px",...} ],  // inches
      "body_box": {"x","y","w","h"},   // only when body_zone is declared: the
                                       //   rect reserved for this page's body
                                       //   text (L1-L3 bullets / module blocks)
                                       //   — inside the content area, never
                                       //   touching the title or caption bands
      "figure_coverage": 0.62,    // panel area / usable figure area (with a
                                  //   text column: panel area / REDUCED figure
                                  //   area, so the number stays honest)
      "warnings": [ "..." ],
      "alternatives": [ {"layout","penalty"} ]   // runner-up candidates
    }

Notes:
    - SCALE RULE: every panel is scaled to the LARGEST size the no-overlap
      constraint permits — that is the whole rule. Concretely, scale =
      min(cw/w, ch/h) per cell, and cells are carved out of the content area
      *after* the title/caption bands are reserved, so a panel can never
      encroach on the title, the caption, or a neighbour. Maximising within
      that bound is the target; there is no reserved headroom to spend on
      making text bigger.
    - CONSTRAINT PRIORITY (do not reorder — a later constraint never justifies
      breaking an earlier one):
        1. No overlap (hard, structural).
        2. Equal scale within a group (no aspect distortion).
        3. Whitespace balance, reading order.
      Text legibility is NOT a constraint here — it is a *reported metric*.
      When the resulting on-slide text falls under the comfort floor, the
      number is printed as a note so the agent knows what it is shipping, but
      it neither blocks the run (`ok` stays true) nor gets escalated to the
      user. Splitting a page is the agent's own call, not a forced hand-off.
      Never trade overlap for bigger text.
    - "text_px" is estimated by viewing the crop (agent is multimodal); if
      omitted, the text-size check is skipped for that panel and a warning is
      emitted.
    - The comfort floor is the same number as the minimum content tier
      (L4 = 14 pt, i.e. 18.67 px at 96 dpi; see references/typography-spec.md),
      but the two mean different things: content text is authored by the agent
      and MUST be >= 14 pt, whereas figure-internal text is a *consequence* of
      the scale above and is only reported. Do not "fix" a small figure text by
      touching the type scale.
    - Panels sharing a "group" are rendered at the same scale so relative
      proportions from the original figure are preserved.
    - PANEL CAPTIONS: when `panel_caption_in` > 0, that strip is carved out of
      the BOTTOM of every cell, the panel is bottom-aligned inside what is
      left, and each placement gains `cap_x/cap_y/cap_w/cap_h` — the exact box
      for its own "Fig. 5a  <one-line title>" row. Panels therefore sit on a
      shared baseline per row, so the caption rows line up. Apply those four
      numbers verbatim; do not re-derive them.
      The reserve slightly lowers `figure_coverage` (the metric counts panel
      area only). That is expected: a titled panel beats an untitled one, so
      never drop the caption to protect the number.
    - No external dependencies for PNG inputs (size read from IHDR);
      Pillow is used as fallback for other formats.
"""

import argparse
import json
import math
import sys
from pathlib import Path


def image_size(path: str):
    """Return (width, height) in pixels; PNG via header, else Pillow."""
    with open(path, "rb") as f:
        header = f.read(26)
    if header[:8] == b"\x89PNG\r\n\x1a\n" and header[12:16] == b"IHDR":
        return (int.from_bytes(header[16:20], "big"),
                int.from_bytes(header[20:24], "big"))
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        raise SystemExit(f"[ERROR] cannot read image size: {path}")


def content_area(p):
    x0 = p["margin_in"]
    x1 = p["slide_w_in"] - p["margin_in"]
    y0 = p["title_band_in"] + 0.1
    # caption_band_in is optional: 0 (or absent) = no caption band, so the
    # bottom simply keeps the page margin instead of a band + 0.1 gutter.
    # Note 0 is NOT the default posture: a page that introduces a new term
    # needs one 14pt line (0.42) for its annotation (SKILL.md 3.6).
    band = p.get("caption_band_in") or 0
    y1 = p["slide_h_in"] - (band + 0.1 if band > 0 else p["margin_in"])
    return (x0, y0, x1, y1)


def cells_grid(rows, cols, area, gap):
    x0, y0, x1, y1 = area
    W, H = x1 - x0, y1 - y0
    cw = (W - gap * (cols - 1)) / cols
    ch = (H - gap * (rows - 1)) / rows
    cells = []
    for r in range(rows):
        for c in range(cols):
            cx = x0 + c * (cw + gap)
            cy = y0 + r * (ch + gap)
            cells.append((cx, cy, cx + cw, cy + ch))
    return cells


def cells_big_side(area, gap, side, big_frac=0.55, n_small=2):
    """One big full-height panel on `side`, `n_small` stacked on the other."""
    x0, y0, x1, y1 = area
    W, H = x1 - x0, y1 - y0
    bigw = W * big_frac - gap / 2
    smallw = W - bigw - gap
    sh = (H - gap * (n_small - 1)) / n_small
    if side == "left":
        big = (x0, y0, x0 + bigw, y1)
        sx = x0 + bigw + gap
    else:
        big = (x1 - bigw, y0, x1, y1)
        sx = x0
    small = [(sx, y0 + i * (sh + gap), sx + smallw, y0 + i * (sh + gap) + sh)
             for i in range(n_small)]
    return ([big] + small) if side == "left" else (small + [big])


def cells_full_row(area, gap, wide_top):
    """n=3 variant: one full-width row on one side, two half-width cells on
    the other. Suits one wide panel + two narrower ones."""
    x0, y0, x1, y1 = area
    W, H = x1 - x0, y1 - y0
    rowH = H * 0.5 - gap / 2
    restH = H - rowH - gap
    halfW = (W - gap) / 2
    if wide_top:
        return [(x0, y0, x1, y0 + rowH),
                (x0, y1 - restH, x0 + halfW, y1),
                (x0 + halfW + gap, y1 - restH, x1, y1)]
    return [(x0, y0, x0 + halfW, y0 + rowH),
            (x0 + halfW + gap, y0, x1, y0 + rowH),
            (x0, y1 - rowH, x1, y1)]


def candidates(n, area, gap):
    """Enumerate (name, cells, kind); kind is "grid" (uniform cells) or
    "big" (one dominant cell). The set grows with panel count AND carries
    emphasis variants, so the per-page choice can follow the page's content
    amount and its key/supporting relationships — not a fixed grid."""
    C = []
    if n == 1:
        C.append(("1x1", cells_grid(1, 1, area, gap), "grid"))
    elif n == 2:
        C.append(("1x2-row", cells_grid(1, 2, area, gap), "grid"))
        C.append(("2x1-col", cells_grid(2, 1, area, gap), "grid"))
        C.append(("1+1-big-left", cells_big_side(area, gap, "left", n_small=1), "big"))
        C.append(("1+1-big-right", cells_big_side(area, gap, "right", n_small=1), "big"))
    elif n == 3:
        C.append(("1x3-row", cells_grid(1, 3, area, gap), "grid"))
        C.append(("3x1-col", cells_grid(3, 1, area, gap), "grid"))
        C.append(("1+2-big-left", cells_big_side(area, gap, "left", n_small=2), "big"))
        C.append(("2+1-big-right", cells_big_side(area, gap, "right", n_small=2), "big"))
        C.append(("wide-top+2", cells_full_row(area, gap, True), "big"))
        C.append(("2+wide-bottom", cells_full_row(area, gap, False), "big"))
    elif n == 4:
        C.append(("2x2", cells_grid(2, 2, area, gap), "grid"))
        C.append(("1x4-row", cells_grid(1, 4, area, gap), "grid"))
        C.append(("4x1-col", cells_grid(4, 1, area, gap), "grid"))
        C.append(("1+3-big-left", cells_big_side(area, gap, "left", 0.5, n_small=3), "big"))
        C.append(("3+1-big-right", cells_big_side(area, gap, "right", 0.5, n_small=3), "big"))
    else:
        # Every sensible factorization, plus a big-cell variant for 5-6
        # panels. Panel count stays a consequence of the evidence, never a
        # quota; the wider enumeration just stops the optimiser from having
        # exactly one candidate for over-4-panel pages.
        seen = set()
        for rows in range(1, n + 1):
            cols = math.ceil(n / rows)
            if (rows, cols) in seen:
                continue
            seen.add((rows, cols))
            C.append((f"{rows}x{cols}", cells_grid(rows, cols, area, gap), "grid"))
        if n <= 6:
            C.append((f"1+{n-1}-big-left",
                      cells_big_side(area, gap, "left", 0.45, n_small=n - 1), "big"))
            C.append((f"{n-1}+1-big-right",
                      cells_big_side(area, gap, "right", 0.45, n_small=n - 1), "big"))
    return C


def assign_key_to_big(panels, cells, side):
    """Return a cell->panel mapping that puts the page's key panel
    (role="key", first one wins) into the big cell, the rest in reading
    order. Returns None (identity) when no key panel is declared."""
    key_idx = [i for i, p in enumerate(panels) if p.get("role") == "key"]
    if not key_idx:
        return None
    n = len(panels)
    big_cell = 0 if side == "left" else len(cells) - 1
    others = [c for c in range(len(cells)) if c != big_cell]
    rest = [i for i in range(n) if i != key_idx[0]]
    assign = [None] * len(cells)
    assign[big_cell] = key_idx[0]
    for c, pi in zip(others, rest):
        assign[c] = pi
    return assign


def evaluate(panels, cells, area, gap, target, min_text,
             cap_in=0.0, cap_gap=0.06, assign=None):
    x0, y0, x1, y1 = area
    area_val = (x1 - x0) * (y1 - y0)

    # cell->panel mapping: identity unless a key panel (role="key") was moved
    # into the big cell. Surplus cells (rows*cols > n) stay empty, as before.
    np_ = len(panels)
    if assign is None:
        assign = list(range(np_))
    pairs = [(i, assign[i]) for i in range(len(cells))
             if i < len(assign) and assign[i] is not None]

    # 1) per-cell fit (aspect preserved), 2) same-group -> same scale (min).
    # `cap_in` is carved out of the cell BOTTOM before fitting, so a panel can
    # never grow into its own caption row.
    img_h = []                       # per-cell height actually available for art
    for cell in cells:
        ch = cell[3] - cell[1]
        img_h.append(max(ch - cap_in, 0.05))
    scales = []
    for i, pi in pairs:
        p = panels[pi]
        cell = cells[i]
        cw = cell[2] - cell[0]
        scales.append(min(cw / p["w"], img_h[i] / p["h"]))
    groups = {}
    for i, pi in pairs:
        if panels[pi].get("group"):
            groups.setdefault(panels[pi]["group"], []).append(i)
    for idxs in groups.values():
        s = min(scales[i] for i in idxs)
        for i in idxs:
            scales[i] = s

    penalty = 0.0
    warnings = []
    placements = []
    for (i, pi), s in zip(pairs, scales):
        p = panels[pi]
        cell = cells[i]
        ih = img_h[i]
        w, h = p["w"] * s, p["h"] * s
        cw, ch = cell[2] - cell[0], cell[3] - cell[1]
        px = cell[0] + (cw - w) / 2
        if cap_in > 0:
            # bottom-align inside the art slot -> caption baselines line up
            py = cell[1] + ih - h
        else:
            py = cell[1] + (ch - h) / 2
        pl = {
            "label": p["label"],
            "file": p["file"],
            "x": round(px, 3), "y": round(py, 3),
            "w": round(w, 3), "h": round(h, 3),
        }
        if cap_in > 0:
            pl["cap_x"] = round(px, 3)
            pl["cap_y"] = round(py + h + cap_gap, 3)
            pl["cap_w"] = round(w, 3)
            pl["cap_h"] = round(max(cap_in - cap_gap, 0.12), 3)
        t = p.get("text_px")
        if t:
            disp = t * s * 96  # inches per px * 96 px/in
            pl["text_display_px"] = round(disp, 1)
            if min_text <= 0:
                # Step 4 is a record, never a layout input. Passing
                # min_text_px <= 0 switches the whole text-size term off, so
                # the optimiser sees only "no overlap + area + whitespace"
                # (SKILL.md step 3; layout-spec.md section four step 3).
                pass
            elif disp < min_text:
                # Informational only — never a blocker and never escalated.
                # The scale below is already the largest the no-overlap
                # constraint permits, so there is nothing left to trade: this
                # number just tells the agent what it is shipping.
                penalty += 40
                warnings.append(
                    f"{p['label']}: text renders at {disp:.0f}px, under the "
                    f"{min_text:g}px comfort floor (= 14 pt; source text {t}px, "
                    f"displayed at {s * 96 * 100:.0f}% of source size). This is "
                    f"already the largest scale the no-overlap limit allows — "
                    f"no action needed.")
            elif disp < target:
                penalty += (target - disp) * 1.5
            elif disp > target * 2.2:
                penalty += (disp - target * 2.2) * 0.5
        else:
            pl["text_display_px"] = None
            warnings.append(f"{p['label']}: no text_px estimate, text-size check skipped")
        placements.append(pl)

    # ---- figure coverage ----
    # Coverage = placed panel area / USABLE FIGURE AREA (the content area left
    # after the title and caption bands are reserved), NOT the whole slide.
    # This is the metric the 50% *reference line* is expressed in (a diagnostic,
    # not a gate: verify_deck.py reports it as a WARN and it never gates exit).
    used = sum(pl["w"] * pl["h"] for pl in placements)
    coverage = (used / area_val) if area_val else 0.0
    if coverage < 0.5:
        penalty += (0.5 - coverage) * 60
        warnings.append(
            f"figure coverage is {coverage*100:.0f}% of the usable figure area, "
            f"under the 50% reference line. Not a failure: if every panel is "
            f"necessary and this is already the layout maximum, accept it and "
            f"note the number. Never drop or split evidence to raise it.")
    else:
        # above the target, only gently discourage very sparse pages
        penalty += (1 - coverage) * 15

    # NOTE: the retired "text-size disparity across panels > 3x is out of
    # proportion" heuristic is deliberately NOT applied. It pulled the step-4
    # record back into the layout decision (see layout-spec.md section three).

    return penalty, placements, warnings, coverage


def main():
    ap = argparse.ArgumentParser(description="Plan subpanel layout on a slide")
    ap.add_argument("plan", type=Path, help="plan.json (schema in module docstring)")
    ap.add_argument("-o", "--out", type=Path, default=Path("layout.json"))
    args = ap.parse_args()

    p = json.loads(args.plan.read_text(encoding="utf-8"))
    gap = p.get("gap_in", 0.15)
    target = p.get("target_text_px", 20)
    min_text = p.get("min_text_px", 18.67)   # 14 pt content floor, in px @96dpi
    cap_gap = float(p.get("caption_gap_in", 0.06))
    cap_line = float(p.get("caption_line_in", 0.237))   # 14 pt * 1.22 / 72

    # Caption reserve under every panel: explicit height wins; otherwise derive
    # it from the longest `caption_lines` requested; otherwise no reserve.
    if "panel_caption_in" in p:
        cap_in = float(p["panel_caption_in"])
    elif any("caption_lines" in s for s in p["panels"]):
        lines = max(int(s.get("caption_lines", 1)) for s in p["panels"])
        cap_in = 0.0 if lines <= 0 else cap_gap + lines * cap_line
    else:
        cap_in = 0.0

    panels = []
    for spec in p["panels"]:
        w, h = image_size(spec["file"])
        panels.append({"w": w, "h": h, **spec})

    n = len(panels)
    if n < 1:
        sys.exit("[ERROR] no panels in plan")
    if n > 4:
        # Panel count is a *consequence* of how much evidence the page's
        # question needs, never a quota (SKILL.md step 2, layout-spec.md
        # section three rule 0). Do not tell the caller to split: ask the
        # only question that can justify splitting.
        print(f"[NOTE] {n} panels on one slide. Panel count is not a quota — "
              f"keep all of them if they answer one question and removing any "
              f"would leave the conclusion unsupported. Split only if some of "
              f"them could stand as a question of their own "
              f"(layout-spec.md section three rule 0).")

    # ---- per-page adaptive inputs -------------------------------------
    # The layout is chosen FROM THIS PAGE'S CONTENT: panel count (evidence),
    # text volume (body zone or not), and the key/supporting relationship
    # (role / structure). The fixed page furniture — title band, page number,
    # 14pt caption band, per-panel caption rows — is reserved before any
    # candidate is scored and can never be encroached on.
    structure = p.get("structure")            # "hierarchy" | "parallel" | None
    n_key = sum(1 for sp in panels if sp.get("role") == "key")
    if n_key > 1:
        print("[NOTE] more than one role='key' panel: only the first is "
              "treated as decisive; a page has at most one key evidence.")

    # ---- body-text zone (正文区) vs caption band (图注带) ---------------
    # Two different modules: the 14pt caption band (cap-band, L4) is the
    # bottom strip, reserved via caption_band_in — and it is OPTIONAL:
    # caption_band_in = 0 means the page has no band and the figure area
    # takes that height back. When the page does have one, size it by LINE
    # COUNT (see the schema note at the top: 1 line 0.42 / 2 lines 0.66 /
    # 3 lines 0.90). The body text (L1-L3
    # bullets / module blocks) is a separate, freely placeable module —
    # declare body_zone {pos, in} to reserve it top/bottom/left/right of the
    # figure area. No body_zone -> legacy behaviour (body strip right above
    # the caption band, hand-placed by the generator).
    body = p.get("body_zone") or {}
    body_pos = body.get("pos")
    body_in = body.get("in", 0)
    if p.get("text_col_in"):                  # legacy alias -> left/right
        body_pos = body_pos or ("right" if p.get("text_col_side") == "right"
                                else "left")
        body_in = body_in or float(p["text_col_in"])

    area = content_area(p)
    body_box = None
    body_prefix = ""
    if body_pos:
        if body_pos not in ("top", "bottom", "left", "right"):
            sys.exit("[ERROR] body_zone.pos must be top / bottom / left / right")
        if body_in <= 0:
            sys.exit("[ERROR] body_zone.in (zone height or width, inches) is "
                     "required when body_zone.pos is set")
        x0, y0, x1, y1 = area
        gutter = gap * 2                      # breathing room text <-> figures
        if body_pos in ("left", "right"):
            if body_in >= (x1 - x0) * 0.6:
                sys.exit("[ERROR] body zone leaves under 40% of the content "
                         "width for figures — shrink the zone or move body "
                         "text back to the caption band")
            if body_pos == "right":
                bb = (x1 - body_in, y0, x1, y1)
                area = (x0, y0, x1 - body_in - gutter, y1)
                body_prefix = "bodyR|"
            else:
                bb = (x0, y0, x0 + body_in, y1)
                area = (x0 + body_in + gutter, y0, x1, y1)
                body_prefix = "bodyL|"
        else:
            if body_in >= (y1 - y0) * 0.5:
                sys.exit("[ERROR] body zone leaves under 50% of the content "
                         "height for figures — shrink the zone or move body "
                         "text back to the caption band")
            if body_pos == "top":
                bb = (x0, y0, x1, y0 + body_in)
                area = (x0, y0 + body_in + gutter, x1, y1)
                body_prefix = "bodyT|"
            else:                             # bottom: right above cap-band
                bb = (x0, y1 - body_in, x1, y1)
                area = (x0, y0, x1, y1 - body_in - gutter)
                body_prefix = "bodyB|"
        body_box = {"x": round(bb[0], 3), "y": round(bb[1], 3),
                    "w": round(bb[2] - bb[0], 3), "h": round(bb[3] - bb[1], 3)}

    results = []
    for name, cells, kind in candidates(n, area, gap):
        assign = None
        if kind == "big" and n_key:
            side = "right" if name.endswith("big-right") else "left"
            assign = assign_key_to_big(panels, cells, side)
        penalty, placements, warnings, coverage = evaluate(
            panels, cells, area, gap, target, min_text, cap_in, cap_gap,
            assign)
        # structure bias: small, deterministic nudge — geometry still wins
        # whenever the difference is real (the nudge never overrides fit).
        if structure == "hierarchy" and kind == "big":
            penalty -= 4
        elif structure == "parallel" and kind == "grid":
            penalty -= 4
        if cap_in > 0:
            # A reserve that eats most of a cell means the page is over-packed
            # for the number of panels it carries — report, do not block.
            min_cell_h = min(c[3] - c[1] for c in cells)
            if cap_in > 0.45 * min_cell_h:
                warnings.append(
                    f"{name}: caption reserve {cap_in:.2f}in takes "
                    f"{cap_in / min_cell_h * 100:.0f}% of the shortest cell "
                    f"({min_cell_h:.2f}in) — shorten the titles or carry fewer "
                    f"panels on this page")
        results.append({"layout": body_prefix + name,
                        "penalty": round(penalty, 2),
                        "placements": placements, "warnings": warnings,
                        "coverage": coverage})

    results.sort(key=lambda r: r["penalty"])
    best = results[0]
    # `ok` is only false on a genuine hard failure: a panel that cannot be
    # placed at all (non-positive extent). Text size is deliberately NOT a
    # blocker — every panel is already scaled to the largest size the
    # no-overlap constraint allows, so there is no decision left to escalate.
    unplaceable = [pl["label"] for pl in best["placements"]
                   if pl["w"] <= 0 or pl["h"] <= 0]
    out = {
        "ok": not unplaceable,
        "layout": best["layout"],
        "placements": best["placements"],
        # panel area / usable figure area (excludes title & caption bands)
        "figure_coverage": round(best["coverage"], 3),
        "panel_caption_in": round(cap_in, 3),
        "warnings": best["warnings"],
        "alternatives": [{"layout": r["layout"], "penalty": r["penalty"]}
                         for r in results[1:]],
    }
    if unplaceable:
        out["error"] = f"panel(s) {unplaceable} cannot fit any candidate cell"
    if body_box:
        out["body_box"] = body_box
    args.out.write_text(json.dumps(out, indent=2, ensure_ascii=False),
                        encoding="utf-8")

    print(f"Slide {p['slide_w_in']}x{p['slide_h_in']}in, {n} panel(s) -> "
          f"best layout: {best['layout']} (penalty {best['penalty']})")
    if body_box:
        print(f"body-text zone reserved ({body_pos}): "
              f"x={body_box['x']:.2f} y={body_box['y']:.2f} "
              f"w={body_box['w']:.2f} h={body_box['h']:.2f} "
              f"(place this page's bullets / module blocks there, verbatim; "
              f"the 14pt caption band below is a separate module — one line "
              f"(0.42 in) if the page introduces a term, 0.66 / 0.90 if it "
              f"also carries a bottleneck or reagent line, 0 = no band)")
    if not (p.get("caption_band_in") or 0):
        print("[NOTE] caption_band_in = 0, so no bottom band is reserved. "
              "That is only right if every term on this page was already "
              "established earlier — a first-mention term needs one 14pt "
              "annotation line (0.42 in) at the bottom of the page.")
    if cap_in > 0:
        print(f"panel caption reserve: {cap_in:.2f}in under every panel "
              f"(apply cap_x/cap_y/cap_w/cap_h verbatim)")
    print(f"figure coverage: {best['coverage']*100:.0f}% of the usable figure area"
          f"{'  [under the 50% reference line — diagnostic, not a failure]' if best['coverage'] < 0.5 else ''}")
    for pl in best["placements"]:
        t = pl["text_display_px"]
        line = (f"  {pl['label']:10s} x={pl['x']:.2f} y={pl['y']:.2f} "
                f"w={pl['w']:.2f} h={pl['h']:.2f}  text={t if t else 'n/a'}px")
        if "cap_y" in pl:
            line += (f"  | cap x={pl['cap_x']:.2f} y={pl['cap_y']:.2f} "
                     f"w={pl['cap_w']:.2f} h={pl['cap_h']:.2f}")
        print(line)
    for r in results[1:]:
        print(f"  alternative: {r['layout']:14s} penalty {r['penalty']}")
    if best["warnings"]:
        print("\nNOTES:")
        for w in best["warnings"]:
            print(f"  - {w}")
    if unplaceable:
        print(f"\n[ERROR] panel(s) {unplaceable} cannot fit any candidate cell — "
              f"check the content area / band heights in plan.json.")
    print(f"\nLayout written to {args.out}")


if __name__ == "__main__":
    main()
