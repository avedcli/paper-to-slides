#!/usr/bin/env python3
"""Inspect a .pptx template and emit a template profile.

Used by the paper-to-slides skill before generating slides: the profile gives
the exact facts the generator needs — slide size, available layouts and their
placeholder geometry, theme fonts/colors, and any decorative shapes that eat
into the usable content area.

Usage:
    python inspect_template.py template.pptx [-o profile.json]

Stdlib only (zipfile + ElementTree) — no dependencies.

Output profile.json:
    {
      "slide_size": {"w_in": 13.333, "h_in": 7.5, "ratio": "16:9"},
      "theme": {"major_latin": "...", "major_ea": "...", ...,
                "colors": {"accent1": "4472C4", ...}},
      "layouts": [
        {"index": 0, "file": "slideLayout1.xml", "name": "Title Slide",
         "kind_guess": "cover",
         "custom_design": true,
         "design_signals": ["userDrawn=1", "1 hand-drawn shape(s)"],
         "scheme_refs": ["accent1"],
         "placeholders": [{"type": "ctrTitle", "idx": 0,
                           "x": 0.9, "y": 2.2, "w": 11.5, "h": 1.6,
                           "size_pt_override": 26.0, "source": "layout"}],
         "inherited_size_risks": [{"type": "body", "idx": 12, "box_h_in": 0.283,
                                   "inherits_pt": 32.0, "line_h_in": 0.533,
                                   "overflow_by_in": 0.25}],
         "decorations": [{"kind": "picture", "x": 0, "y": 0, "w": 13.33, "h": 1.2,
                          "area_pct": 9.0}]}
      ],
      "master_text_styles": {"slideMaster1.xml": {"title_pt": 44.0, "body_pt": 32.0}},
      "master_hf": {"slideMaster1.xml": {"sldNum": "0", "hdr": "0", "ftr": "0", "dt": "0"}},
      "sample_slides": [{"file": "slide1.xml", "layout": "Title Slide"}],
      "sample_slides_detail": [{"file": "slide1.xml", "layout": "Title Slide",
                                "shapes": 9, "with_geometry": 9, "with_text": 8,
                                "empty_placeholders": 0, "pictures": 1,
                                "max_descr_len": 0, "texts": ["..."]}],
      "content_area": {"x": 0.5, "y": 1.35, "w": 12.3, "h": 5.4},
      "warnings": ["..."]
    }

Beyond geometry, the profile reports whether the template really carries its own
design. What is flagged in `warnings`:

- layouts that carry nothing of their own — no `userDrawn` flag, no hand-drawn
  shape, no placeholder geometry/style override (`custom_design: false`).
  Placeholder prompt text is NOT such a signal: every hand-built Office template
  keeps "Click to edit master title style" inside its layout placeholders.
- placeholders with no font size of their own that fall back to the master default
  (Office ships 44pt title / 32pt body) and end up taller than their own box
- a slide-number placeholder that exists but is switched off by the master's
  `<p:hf sldNum="0"/>`
- layouts painting with theme colour slots (`accent1` …) instead of literal hex,
  so the rendered colour depends on the theme palette
- slides that are empty shells (placeholders with no geometry/text), embedded
  source code in cNvPr/@descr, and hardcoded page numbers

A template that trips those checks must be fixed (or rebuilt from source) before
slides are generated.
"""

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EMU_PER_IN = 914400.0

# placeholder types that carry the page's main body text
BODY_TYPES = {"body", "subTitle", "obj"}
TITLE_TYPES = {"title", "ctrTitle"}
# placeholders that do not carry page content (and never count as "content")
INERT_TYPES = {"sldNum", "dt", "ftr", "hdr"}

# fonts shipped with Windows/Office (present on most machines)
SYSTEM_FONTS = {
    "calibri", "arial", "helvetica", "times new roman", "segoe ui", "verdana",
    "tahoma", "georgia", "cambria", "consolas", "courier new", "trebuchet ms",
    "aptos", "aptos display", "arial black", "impact", "franklin gothic book",
    "microsoft yahei", "微软雅黑", "black", "simhei", "黑体", "simsun", "宋体",
    "nsimsun", "新宋体", "dengxian", "等线", "kaiti", "楷体", "fangsong", "仿宋",
    "ms pgothic", "meiryo", "yu gothic", "noto sans", "source han sans",
}

# stock Office layouts still carry these prompt strings (zh + en)
STOCK_PROMPT_RE = re.compile(
    r"单击此处编辑|单击此处添加|Click to edit|Click to add", re.IGNORECASE)
# hardcoded page numbers such as "01 / 09" that are plain text, not a field
PAGE_NUM_RE = re.compile(r"(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})(?!\d)")
# a descr longer than this is embedded source code, not alt text
DESCR_CODE_MIN = 200


def emu2in(v):
    return round(v / EMU_PER_IN, 3) if v is not None else None


def bbox(el):
    """Return (x,y,w,h) in inches from the first a:xfrm found, else None."""
    xfrm = el.find(".//a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return (emu2in(int(off.get("x"))), emu2in(int(off.get("y"))),
            emu2in(int(ext.get("cx"))), emu2in(int(ext.get("cy"))))


def placeholders(tree):
    out = []
    for sp in tree.findall(".//p:cSld/p:spTree/p:sp", NS):
        ph = sp.find(".//p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None:
            continue
        bb = bbox(sp)
        item = {
            "type": ph.get("type", "body"),
            "idx": int(ph.get("idx", "0")),
            "source": "layout",
        }
        if bb:
            item.update(zip(("x", "y", "w", "h"), bb))
        else:
            item["source"] = "inherits-master"
        # per-placeholder font / size overrides.
        # NOTE: match on the tree, not on a serialized string — ElementTree writes
        # namespaced tags as <ns0:defRPr ...>, so a regex for '<a:defRPr' silently
        # misses every override.
        dpr = sp.find(".//a:lstStyle/a:lvl1pPr/a:defRPr", NS)
        if dpr is None:
            dpr = sp.find(".//a:lstStyle//a:defRPr", NS)
        if dpr is not None and dpr.get("sz"):
            item["size_pt_override"] = round(int(dpr.get("sz")) / 100.0, 1)
        fonts = {el.get("typeface") for el in sp.iter()
                 if el.tag in (f"{{{NS['a']}}}latin", f"{{{NS['a']}}}ea") and el.get("typeface")}
        if fonts:
            item["fonts"] = sorted(fonts)
        out.append(item)
    return out


def layout_fonts(tree):
    """Fonts actually referenced inside a layout (any element)."""
    fonts = set()
    for el in tree.iter():
        if el.tag in (f"{{{NS['a']}}}latin", f"{{{NS['a']}}}ea") and el.get("typeface"):
            fonts.add(el.get("typeface"))
    return sorted(fonts)


def part_texts(tree):
    """Every non-empty <a:t> string inside a part."""
    return [t.text.strip() for t in tree.iter(f"{{{NS['a']}}}t")
            if t.text and t.text.strip()]


def slide_census(tree):
    """How much real content a slide actually carries.

    A slide made only of placeholder <p:sp> elements with an empty <p:spPr/> and an
    empty <p:txBody/> renders as a blank page: nothing was baked in. This is the
    failure mode of HTML->pptx exporters that ship an empty shell plus the source
    markup stashed in cNvPr/@descr.
    """
    stats = {"shapes": 0, "with_geometry": 0, "with_text": 0,
             "empty_placeholders": 0, "pictures": 0, "max_descr_len": 0}
    for sp in tree.findall(".//p:cSld/p:spTree/p:sp", NS):
        stats["shapes"] += 1
        if bbox(sp) is not None:
            stats["with_geometry"] += 1
        if any(t.text and t.text.strip() for t in sp.iter(f"{{{NS['a']}}}t")):
            stats["with_text"] += 1
        elif sp.find(".//p:nvSpPr/p:nvPr/p:ph", NS) is not None:
            stats["empty_placeholders"] += 1
    stats["pictures"] = len(list(tree.iter(f"{{{NS['p']}}}pic")))
    descrs = [len(el.get("descr") or "") for el in tree.iter(f"{{{NS['p']}}}cNvPr")]
    stats["max_descr_len"] = max(descrs) if descrs else 0
    return stats


def master_text_styles(m_root):
    """Title / body default sizes from the master's txStyles (points)."""
    out = {}
    for style, label in (("titleStyle", "title"), ("bodyStyle", "body"),
                         ("otherStyle", "other")):
        el = m_root.find(f".//p:txStyles/p:{style}/a:lvl1pPr/a:defRPr", NS)
        if el is not None and el.get("sz"):
            out[f"{label}_pt"] = round(int(el.get("sz")) / 100.0, 1)
    return out


def hf_flags(m_root):
    """Master header/footer toggles: <p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>.

    A slide-number / footer / date placeholder can exist in the master and every
    layout and still never render, because these flags switch it off. Value "0"
    means off.
    """
    hf = m_root.find(".//p:hf", NS)
    if hf is None:
        return {}
    return {k: hf.get(k) for k in ("sldNum", "hdr", "ftr", "dt") if hf.get(k) is not None}


def layout_authenticity(tree, phs):
    """Judge whether a layout carries the deck's own page design.

    Placeholder prompt text ("Click to edit master title style" / "单击此处编辑母版
    标题样式") is NOT evidence of a stock layout: every hand-built Office template
    keeps it inside its layout placeholders — that is what the placeholder shows
    while you edit the layout. What marks a real layout is having something of its
    own. Returns the list of such signals (empty list = nothing of its own).
    """
    signals = []
    if tree.get("userDrawn") == "1":
        signals.append("userDrawn=1")
    extras = sum(
        1 for sp in tree.findall(".//p:cSld/p:spTree/p:sp", NS)
        if sp.find(".//p:nvSpPr/p:nvPr/p:ph", NS) is None)
    if extras:
        signals.append(f"{extras} hand-drawn shape(s)")
    if any(ph.get("source") == "layout" and "x" in ph for ph in phs):
        signals.append("explicit placeholder geometry")
    styled = sum(1 for ph in phs if "size_pt_override" in ph or "fonts" in ph)
    if styled:
        signals.append(f"{styled} placeholder style override(s)")
    return signals


def inherited_size_risks(phs, master_styles):
    """Placeholders that define no font size fall back to the master default
    (Office ships 44pt title / 32pt body). If that line is taller than the
    placeholder box, the text overflows its own box and can overlap its
    neighbours — a template looks "designed" but is not WYSIWYG."""
    out = []
    for ph in phs:
        if "size_pt_override" in ph or "h" not in ph or ph["type"] in INERT_TYPES:
            continue
        default_pt = master_styles.get(
            "title_pt" if ph["type"] in TITLE_TYPES else "body_pt")
        if not default_pt:
            continue
        line_in = round(default_pt * 1.2 / 72.0, 3)
        if line_in > ph["h"]:
            out.append({
                "type": ph["type"], "idx": ph["idx"], "box_h_in": ph["h"],
                "inherits_pt": default_pt, "line_h_in": line_in,
                "overflow_by_in": round(line_in - ph["h"], 3),
            })
    return out


def scheme_refs(tree):
    """Theme colour slots a layout leans on (accent1, accent5, ...).

    A layout that paints with <a:schemeClr val="accent1"/> does not show the
    hex in the file — the rendered colour is whatever the theme defines, which
    may differ from the design's own brand colour.
    """
    slots = {el.get("val") for el in tree.iter(f"{{{NS['a']}}}schemeClr") if el.get("val")}
    return sorted(slots - {"tx1", "tx2", "bg1", "bg2", "phClr"})


def figure_area(layout, slide_w, slide_h, margin=0.44, bottom=0.3):
    """Usable area left for figures once the layout's own furniture is respected."""
    phs = [p for p in layout["placeholders"] if p["type"] not in INERT_TYPES and "y" in p]
    title_bottom = 0.0
    for p in phs:
        if p["type"] in TITLE_TYPES:
            title_bottom = max(title_bottom, p["y"] + p["h"])
    blocked_bottom = max(title_bottom, max((p["y"] + p["h"] for p in phs), default=0.0)
                         if all(p["type"] in TITLE_TYPES for p in phs) else title_bottom)
    y0 = round(max(title_bottom, 0.0) + 0.1, 3) if title_bottom else round(margin, 3)
    return {
        "x": margin,
        "y": y0,
        "w": round(slide_w - 2 * margin, 3),
        "h": round(slide_h - y0 - bottom, 3),
        "title_band_bottom": round(title_bottom, 3) if title_bottom else None,
    }


def decorations(tree, slide_w, slide_h):
    """Non-placeholder shapes / pics — they may eat into the content area."""
    out = []
    tags = [("p:sp", "shape"), ("p:pic", "picture"),
            ("p:graphicFrame", "graphicFrame"), ("p:grpSp", "group")]
    for tag, kind in tags:
        for el in tree.findall(f".//p:cSld/p:spTree/{tag}", NS):
            if el.find(".//p:nvSpPr/p:nvPr/p:ph", NS) is not None:
                continue  # placeholder, not decoration
            bb = bbox(el)
            if not bb:
                continue
            x, y, w, h = bb
            if slide_w and slide_h:
                area_pct = round(100.0 * (w * h) / (slide_w * slide_h), 1)
            else:
                area_pct = None
            out.append({"kind": kind, "x": x, "y": y, "w": w, "h": h,
                        "area_pct": area_pct})
    return out


def guess_kind(name, phs):
    """Decide a layout's role from its name + placeholder geometry.

    Heuristics, in priority order:
      1. ctrTitle placeholder → cover (definitive)
      2. name contains "封面" / "cover" / "title slide" / "title page" → cover
      3. name contains "分节" / "section" / "divider" / "过渡" / "question" → section-divider
      4. name is exactly "标题" (single-token Chinese title) → cover (common naming)
      5. title placeholder height ≥ 1.5in and no body placeholder → cover
      6. title + body → content
      7. only title → title-only
      8. no placeholders → blank
    """
    n = (name or "").strip()
    real = [p for p in phs if p["type"] not in INERT_TYPES]
    types = {p["type"] for p in real}

    if "ctrtitle" in types:
        return "cover"
    if any(k in n for k in ("封面",)) or "title slide" in n.lower() or "title page" in n.lower() or "cover" in n.lower():
        return "cover"
    # Common Chinese single-token layout names
    if n in ("标题", "标题页"):
        return "cover"
    if any(k in n for k in ("分节", "section", "divider", "过渡", "question")):
        return "section-divider"

    if not real:
        return "blank"
    if types & {"title"}:
        title = next((p for p in real if p["type"] == "title"), None)
        if title and not (types & BODY_TYPES) and title.get("h", 0) >= 1.5:
            return "cover"
        if not (types & BODY_TYPES):
            return "title-only"
        return "content"
    return "other"


def read_theme(z):
    """读主题：优先从 slideMaster1 的 rels 拿到它引用的 theme（=primary theme），
    因为多套主题共存时 primary 才是真正被 deck 渲染时引用的；否则回退到
    theme1.xml；最后回退到 namelist 第一个 theme*.xml（兼容怪导出器）。
    """
    # 1) 跟随 slideMaster1.xml.rels 找到 primary theme
    try:
        sm_rels = rels_map(z, "ppt/slideMasters/slideMaster1.xml")
        # 找一个 Target 以 theme 结尾的
        for rid, target in sm_rels.items():
            if "theme" in target and target.endswith(".xml"):
                primary = norm(str(PurePosixPath("ppt/slideMasters/slideMaster1.xml").parent), target)
                if primary in z.namelist():
                    return _parse_theme(z.read(primary))
    except Exception:
        pass
    # 2) 显式找 theme1.xml
    for n in z.namelist():
        if re.match(r"ppt/theme/theme1\.xml$", n):
            return _parse_theme(z.read(n))
    # 3) 回退：namelist 第一个 theme*.xml
    theme_path = next((n for n in z.namelist()
                       if re.match(r"ppt/theme/theme\d+\.xml$", n)), None)
    if not theme_path:
        return {}
    return _parse_theme(z.read(theme_path))


def _parse_theme(blob):
    root = ET.fromstring(blob)
    out = {}
    fs = root.find(".//a:fontScheme", NS)
    if fs is not None:
        for tag, prefix in (("a:majorFont", "major"), ("a:minorFont", "minor")):
            f = fs.find(tag, NS)
            if f is None:
                continue
            latin = f.find("a:latin", NS)
            ea = f.find("a:ea", NS)
            out[f"{prefix}_latin"] = latin.get("typeface") if latin is not None else None
            out[f"{prefix}_ea"] = ea.get("typeface") if ea is not None else None
    cs = root.find(".//a:clrScheme", NS)
    if cs is not None:
        colors = {}
        for child in cs:
            name = child.tag.split("}")[1]
            srgb = child.find("a:srgbClr", NS)
            sysclr = child.find("a:sysClr", NS)
            if srgb is not None:
                colors[name] = srgb.get("val")
            elif sysclr is not None:
                colors[name] = sysclr.get("lastClr")
        out["colors"] = colors
        out["theme_file"] = "(from slideMaster1 rels)"
    return out


def rels_map(z, part):
    p = PurePosixPath(part)
    rels_path = str(p.parent / "_rels" / (p.name + ".rels"))
    if rels_path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_path))
    return {r.get("Id"): r.get("Target") for r in root.findall("rel:Relationship", NS)}


def norm(base_dir, target):
    """Resolve a relationship target against its base dir, collapsing '../'."""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def main():
    ap = argparse.ArgumentParser(description="Inspect a pptx template")
    ap.add_argument("pptx", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=Path("template_profile.json"))
    args = ap.parse_args()

    if not args.pptx.exists():
        sys.exit(f"[ERROR] not found: {args.pptx}")

    warnings = []
    z = zipfile.ZipFile(args.pptx)
    names = z.namelist()

    # ---- slide size ----
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    sld_sz = pres.find("p:sldSz", NS)
    sw = emu2in(int(sld_sz.get("cx"))) if sld_sz is not None else None
    sh = emu2in(int(sld_sz.get("cy"))) if sld_sz is not None else None
    ratio = None
    if sw and sh:
        r = round(sw / sh, 4)
        ratio = "16:9" if abs(r - 16 / 9) < 0.02 else \
                "4:3" if abs(r - 4 / 3) < 0.02 else f"{r:.3f}"

    # ---- masters & their layouts (ordered) ----
    master_files = sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
    layouts = []
    idx = 0
    master_styles = {}
    master_hf = {}
    for mf in master_files:
        m_root = ET.fromstring(z.read(mf))
        master_styles[Path(mf).name] = master_text_styles(m_root)
        master_hf[Path(mf).name] = hf_flags(m_root)
        m_rels = rels_map(z, mf)
        for lid in m_root.findall(".//p:sldLayoutIdLst/p:sldLayoutId", NS):
            rid = lid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = m_rels.get(rid)
            if not target:
                continue
            lpath = norm(str(PurePosixPath(mf).parent), target)
            if lpath not in names:
                continue
            l_root = ET.fromstring(z.read(lpath))
            cSld = l_root.find(".//p:cSld", NS)
            lname = cSld.get("name") if cSld is not None else None
            phs = placeholders(l_root)
            decos = decorations(l_root, sw, sh)
            l_texts = part_texts(l_root)
            layout = {
                "index": idx,
                "file": Path(lpath).name,
                "name": lname,
                "kind_guess": guess_kind(lname, phs),
                "placeholders": phs,
                "decorations": decos,
                "fonts": layout_fonts(l_root),
                # prompt text is informational only — see layout_authenticity()
                "stock_prompt": next((t for t in l_texts if STOCK_PROMPT_RE.search(t)), None),
                "design_signals": layout_authenticity(l_root, phs),
                "scheme_refs": scheme_refs(l_root),
                "inherited_size_risks": inherited_size_risks(
                    phs, master_styles.get(Path(mf).name, {})),
            }
            layout["custom_design"] = bool(layout["design_signals"])
            if sw and sh:
                layout["figure_area"] = figure_area(layout, sw, sh)
            layouts.append(layout)
            idx += 1

    # ---- master-level decorations (backgrounds, logos on every page) ----
    master_decos = {}
    for mf in master_files:
        m_root = ET.fromstring(z.read(mf))
        master_decos[Path(mf).name] = decorations(m_root, sw, sh)

    # ---- sample slides ----
    slides = []
    slides_detail = []
    pres_rels = rels_map(z, "ppt/presentation.xml")
    for sld_id in pres.findall(".//p:sldIdLst/p:sldId", NS):
        rid = sld_id.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = pres_rels.get(rid)
        if not target:
            continue
        spath = norm("ppt", target)
        if spath not in names:
            continue
        s_root = ET.fromstring(z.read(spath))
        s_rels = rels_map(z, spath)
        layout_name = None
        for r_target in s_rels.values():
            if "slideLayout" in r_target:
                lp = norm(str(PurePosixPath(spath).parent), r_target)
                if lp in names:
                    lr = ET.fromstring(z.read(lp))
                    cs = lr.find(".//p:cSld", NS)
                    layout_name = cs.get("name") if cs is not None else None
        slides.append({"file": Path(spath).name, "layout": layout_name})
        detail = slide_census(s_root)
        all_texts = part_texts(s_root)
        detail["file"] = Path(spath).name
        detail["layout"] = layout_name
        detail["texts"] = all_texts[:8]
        detail["page_number"] = next((t for t in all_texts if PAGE_NUM_RE.search(t)), None)
        slides_detail.append(detail)

    theme = read_theme(z)

    # ---- derive a usable content area ----
    content_area = None
    body_boxes = []
    for lay in layouts:
        for ph in lay["placeholders"]:
            if ph["type"] in BODY_TYPES and "x" in ph:
                body_boxes.append((ph["x"], ph["y"], ph["w"], ph["h"]))
    if body_boxes:
        x0 = min(b[0] for b in body_boxes)
        y0 = min(b[1] for b in body_boxes)
        x1 = max(b[0] + b[2] for b in body_boxes)
        y1 = max(b[1] + b[3] for b in body_boxes)
        content_area = {"x": x0, "y": y0, "w": round(x1 - x0, 3),
                        "h": round(y1 - y0, 3),
                        "basis": "union of body placeholders across layouts"}

    # ---- warnings ----
    kinds = [l["kind_guess"] for l in layouts]
    if "cover" not in kinds:
        warnings.append("no cover-like layout found (no ctrTitle placeholder / 'title slide' name)")
    # A content page is usable with a title placeholder alone: the generator can
    # carry body copy in free textboxes (this is in fact the preferred route for
    # figure-heavy academic slides, since a body placeholder would fight the
    # figure area for space). Only flag it when *no* layout can host a page at all.
    if "content" not in kinds and "title-only" not in kinds:
        warnings.append("no content layout found (need a layout with at least a title placeholder)")
    elif "content" not in kinds:
        warnings.append(
            "content layouts carry only a title placeholder (no body placeholder) — "
            "body copy must be placed with free textboxes; supported, but plan.json "
            "must then carry explicit textbox coordinates for every text block")
    if "blank" not in kinds:
        warnings.append("no blank layout — figure-only pages will have to reuse a content layout "
                        "and hide its placeholders")
    if "section-divider" not in kinds and "title-only" not in kinds:
        warnings.append("no section-divider layout (and no title-only layout) — Results sections "
                        "will lack a divider page style")

    # ---- does the pptx actually carry the template's own design? ----
    # These checks catch "empty shell" templates: a file that has layouts and a slide
    # count, but whose slides are blank placeholder stubs (thing HTML->pptx exporters
    # emit when the design never made it into the OOXML).
    #
    # NOTE: prompt text alone proves nothing. A hand-built layout keeps the stock
    # prompt string in its placeholders — that is by design. Only a layout with
    # *no* signal of its own (see layout_authenticity) is a stock shell.
    stock_layouts = [l["name"] for l in layouts if not l["custom_design"]]
    if stock_layouts:
        sample = next((l["stock_prompt"] for l in layouts
                       if not l["custom_design"] and l.get("stock_prompt")), None)
        warnings.append(
            f"layout(s) {stock_layouts} carry nothing of the deck's own design "
            f"(no userDrawn flag, no hand-drawn shape, no placeholder geometry or style "
            f"override)" + (f" — only stock prompt text: {sample!r}" if sample else "") +
            " — child slides inherit nothing but Office defaults from them")
    no_style = [l["name"] for l in layouts if l["inherited_size_risks"]]
    if no_style:
        worst = next(r for l in layouts for r in l["inherited_size_risks"])
        warnings.append(
            f"layout(s) {no_style} have placeholders with no font size of their own, so they fall "
            f"back to the master default (Office ships 44pt title / 32pt body). At least one of "
            f"those boxes is smaller than its own text line: {worst['type']} idx={worst['idx']} box "
            f"{worst['box_h_in']}in vs {worst['inherits_pt']}pt line ({worst['line_h_in']}in, "
            f"overflows by {worst['overflow_by_in']}in) — text will spill out of its box and can "
            f"overlap the neighbour. Set an explicit size (or autofit) on every such placeholder "
            f"so the template is WYSIWYG")
    for mf, flags in master_hf.items():
        if flags.get("sldNum") == "0":
            has_num_ph = any(ph["type"] == "sldNum"
                             for lay in layouts for ph in lay["placeholders"])
            if has_num_ph:
                warnings.append(
                    f"{mf} has <p:hf sldNum=\"0\"/> — the slide-number placeholder exists in the "
                    f"master and layouts but the header/footer toggle is switched OFF, so the page "
                    f"number will not render. Turn it back on in Slide Master view (or set "
                    f"sldNum=\"1\") before using the template")
    scheme_used = sorted({s for lay in layouts for s in lay["scheme_refs"]})
    if scheme_used:
        accent = (theme or {}).get("colors", {}).get("accent1")
        warnings.append(
            f"layout(s) paint with theme colour slots {scheme_used} instead of literal hex — the "
            f"rendered colour is whatever the theme defines (accent1 = #{accent}). If the design's "
            f"brand colour differs, either write the literal hex into the layout or update the "
            f"theme palette, otherwise the deck silently comes out in the wrong colour")
    empty_slides = [d["file"] for d in slides_detail if d["with_text"] == 0]
    if empty_slides:
        warnings.append(
            f"slide(s) {empty_slides} carry no text at all — these example pages are empty shells "
            f"(placeholders only, no geometry baked in). Do not treat this pptx as a designed "
            f"template; rebuild it from the design source before generating")
    coded = [d["file"] for d in slides_detail if d["max_descr_len"] > DESCR_CODE_MIN]
    if coded:
        warnings.append(
            f"slide(s) {coded} store embedded source code in cNvPr/@descr (up to "
            f"{max(d['max_descr_len'] for d in slides_detail)} chars) — PowerPoint never renders it; "
            f"strip it, it only inflates the file and leaks the generator's markup")
    paged = [d["file"] for d in slides_detail if d.get("page_number")]
    if paged:
        hit = next(d["page_number"] for d in slides_detail if d.get("page_number"))
        warnings.append(
            f"hardcoded page numbers found in {paged} (e.g. {hit!r}) — they are plain text, so they "
            f"will be wrong on any deck with a different page count; clear them from the template "
            f"and use a real slide-number field on the slides you generate")
    if ratio and ratio != "16:9":
        warnings.append(f"slide ratio is {ratio}, not 16:9 — check projector/venue requirement")
    if theme and not theme.get("major_ea"):
        warnings.append("theme has no east-asian (ea) font — CJK text will fall back to a system font")

    # font consistency: theme vs layout-level overrides, and non-system fonts.
    # Theme references ("+mj-ea", "+mn-lt") are PowerPoint placeholders that resolve
    # to the theme's actual fonts at render time — they're NOT hardcoded fonts, so
    # they shouldn't be flagged as non-system. Filter them out before the check.
    all_fonts = set()
    for lay in layouts:
        all_fonts.update(lay["fonts"])
    theme_fonts = {theme.get("major_latin"), theme.get("minor_latin"),
                   theme.get("major_ea"), theme.get("minor_ea")} - {None}
    if all_fonts and theme_fonts and not all_fonts.issubset(theme_fonts):
        # Theme references like "+mj-ea" / "+mn-lt" resolve to the theme's actual
        # fonts at render time — they're NOT extra hardcoded fonts. Only the
        # literals that don't resolve should be flagged.
        resolved = set()
        for f in all_fonts:
            if f.startswith("+mj-lt"):
                resolved.add(theme.get("major_latin"))
            elif f.startswith("+mj-ea"):
                resolved.add(theme.get("major_ea"))
            elif f.startswith("+mn-lt"):
                resolved.add(theme.get("minor_latin"))
            elif f.startswith("+mn-ea"):
                resolved.add(theme.get("minor_ea"))
            elif f.startswith("+mn-cs"):
                # complex script; rare to be a brand-defining font, accept
                resolved.add(theme.get("minor_latin") or theme.get("major_latin"))
            else:
                resolved.add(f)
        resolved.discard(None)
        extra = sorted(resolved - theme_fonts - {None})
        if extra:
            warnings.append(
                f"fonts used in layouts differ from the theme fonts (after resolving "
                f"theme refs): {extra} (theme declares {sorted(theme_fonts)}) — text will "
                f"look inconsistent between cover and content pages; unify them in the "
                f"theme/layouts")
    # Theme references are NOT non-system fonts; filter them out.
    non_system = sorted(
        f for f in all_fonts
        if not f.lower().startswith("+")
        and f.lower() not in SYSTEM_FONTS
    )
    if non_system:
        warnings.append(
            f"non-system font(s) {non_system} — they will not render on machines that lack them "
            f"(silent fallback, layout shifts). Replace with a system font or embed the font")

    if slides:
        warnings.append(f"template contains {len(slides)} sample slide(s); they must be deleted "
                        f"before generating (or the generator must ignore them)")
    big = []
    for lay in layouts:
        for d in lay["decorations"]:
            if (d["area_pct"] or 0) >= 20 and d["kind"] in ("picture", "shape", "graphicFrame"):
                big.append((lay["name"], d))
    if big:
        warnings.append(f"{len(big)} layout-level decorative shape(s) cover >=20% of the slide — "
                        f"verify they do not intrude into the content area")
    if content_area and content_area["h"] < 4.0:
        warnings.append(f"derived content area height is only {content_area['h']}in — figure pages "
                        f"will be cramped; check title/caption bands")

    profile = {
        "file": str(args.pptx),
        "slide_size": {"w_in": sw, "h_in": sh, "ratio": ratio},
        "theme": theme,
        "master_text_styles": master_styles,
        "master_hf": master_hf,
        "layouts": layouts,
        "master_decorations": master_decos,
        "sample_slides": slides,
        "sample_slides_detail": slides_detail,
        "content_area": content_area,
        # How page body copy is meant to be carried:
        #   "placeholder" — content layouts have a body placeholder, fill it
        #   "textbox"     — content layouts only have a title placeholder,
        #                   so every text block must be a free textbox with
        #                   explicit coordinates in plan.json
        "body_carrier": ("placeholder" if "content" in kinds else "textbox"),
        "warnings": warnings,
    }
    args.out.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")

    # ---- human-readable report ----
    print(f"== {args.pptx.name} ==")
    print(f"slide size : {sw} x {sh} in  ({ratio})")
    if theme:
        print(f"theme fonts: major {theme.get('major_latin')} / ea {theme.get('major_ea')}")
        print(f"             minor {theme.get('minor_latin')} / ea {theme.get('minor_ea')}")
        c = theme.get("colors", {})
        if c:
            print("theme colors: " + ", ".join(f"{k}={v}" for k, v in list(c.items())[:8]))
    print(f"\nlayouts ({len(layouts)}):")
    for lay in layouts:
        phs = ", ".join(p["type"] for p in lay["placeholders"]) or "none"
        dmax = max((d["area_pct"] or 0) for d in lay["decorations"]) if lay["decorations"] else 0
        fa = lay.get("figure_area")
        fa_txt = f" | figure area y={fa['y']} h={fa['h']}in" if fa else ""
        tag = "custom" if lay["custom_design"] else "STOCK-SHELL"
        print(f"  [{lay['index']}] {lay['name']:<24} kind={lay['kind_guess']:<16} "
              f"[{tag}] ph: {phs:<38} deco max {dmax}%{fa_txt}")
        if lay["design_signals"]:
            print(f"       design signals: {', '.join(lay['design_signals'])}")
        if lay["fonts"]:
            print(f"       fonts: {', '.join(lay['fonts'])}")
        if lay["scheme_refs"]:
            print(f"       theme colour slots: {', '.join(lay['scheme_refs'])}")
        for p in lay["placeholders"]:
            extra = []
            if p.get("size_pt_override"):
                extra.append(f"{p['size_pt_override']}pt")
            if p.get("fonts"):
                extra.append("/".join(p["fonts"]))
            if extra:
                print(f"       ph[{p['type']}#{p['idx']}] {' '.join(extra)}")
        for r in lay["inherited_size_risks"]:
            print(f"       ! ph[{r['type']}#{r['idx']}] box {r['box_h_in']}in < "
                  f"inherited {r['inherits_pt']}pt line {r['line_h_in']}in "
                  f"(overflows {r['overflow_by_in']}in)")
    if master_styles:
        for k, v in master_styles.items():
            print(f"master text styles ({k}): {v}")
    if master_hf:
        for k, v in master_hf.items():
            print(f"master header/footer toggles ({k}): {v}")
    if content_area:
        print(f"\ncontent area: x={content_area['x']} y={content_area['y']} "
              f"w={content_area['w']} h={content_area['h']} in")
    print(f"sample slides: {len(slides)}")
    for d in slides_detail:
        print(f"  {d['file']:<14} layout={str(d['layout']):<16} "
              f"shapes={d['shapes']} geom={d['with_geometry']} text={d['with_text']} "
              f"empty_ph={d['empty_placeholders']} pics={d['pictures']} "
              f"max_descr={d['max_descr_len']}")
    print("\nWARNINGS:")
    for w in warnings:
        print(f"  - {w}")
    print(f"\nProfile written to {args.out}")


if __name__ == "__main__":
    main()
