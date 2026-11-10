#!/usr/bin/env python3
"""Phase-5 gate: structural + content checks that the count-based check misses.

Usage:
    python verify_deck.py deck.pptx --profile template_profile.json
                          [--pages 9] [--exempt 1,2,6,8,9]
                          [--content-layout 小标题内页]

Why this exists
---------------
Stage 4 step 7 only counts slide parts and shapes; step 8 (measure_deck.py)
only checks text fit. Neither enforces the *content* rules the outline was
approved against, and those are the ones that silently regress when a deck is
regenerated after an edit. This script checks:

  1. slide part count == --pages
  2. every slide: shapes > 0, text shapes > 0, and exactly one slide-number
     field (`<a:fld type="slidenum">`) — i.e. the manual page number really
     landed, and there is no duplicated page number
  3. per page: figure share of the content layout's `figure_area` — reported as a
     DIAGNOSTIC, not a gate (see --figure-blocking / --figure-pct-override). A page
     whose evidence list is scientifically right but whose panels are tall/thin,
     near-square or simply numerous will legitimately sit below 50%; that is not a
     defect, and it must never be "fixed" by dropping or splitting evidence.
  4. per page: body word count (title and "scaffold" text excluded)
  5. Type scale: content-page text must sit on the 3-step ladder defined in
     `references/typography-spec.md` (26 / 22 / 20 / 14 pt — page title / group
     heading / body bullet / sub-item-caption, the last being the hard floor).
     Anything below 14 pt,
     any off-ladder size, pages with more than `--max-tiers-per-page` distinct
     tiers, two different sizes inside ONE textbox, and inverted hierarchy all
     fail.
     This is the check that catches "shave 0.5 pt until it fits", which is how
     a deck ends up with a dozen near-identical sizes.
  6. Discussion: the four elements (innovations / findings / limitations /
     future) are actually present somewhere in the deck

Figure-pane detection: a shape counts toward the figure area if it is a
Picture or an auto-shape whose name starts with "Rectangle" (the dashed
placeholder panes used for external figures). Anything with text that is not
the title and not below `--scaffold-y` counts as body copy.

Caption detection (named, not positional): the generator names each per-panel
caption row `cap-<label>` ("Fig. 5a  <one-line title>", placed directly under
its own image, see layout-spec.md). Those sit inside the figure area — way
above the bottom caption band — so a y-position rule cannot find them. By name
they are counted as scaffold (not body copy) and excluded from the "body block"
that the hierarchy check uses. A page whose pictures have no such rows is
reported in `[panel captions]`; that report is diagnostic, not a gate.

Scaffold text (excluded from the word budget but reported) is the mandated
metadata inside a placeholder pane plus the citation line under it — those
words are prescribed by outline-spec §0.2, they are not body copy.

Exit code 1 on any failure (2-6 above), so it can gate delivery.

Figure share is deliberately NOT part of the exit code: SKILL.md step 3 ranks
"the page's narrative is complete and its key evidence readable" above the
coverage number. Below-threshold pages print `WARN` and are listed under
[figure coverage] for the human to judge; pass --figure-blocking to put them
back into the gate, or --figure-pct-override to declare accepted per-page
thresholds (mirrors --body-words-override).
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

ELEMENTS = ("innovations", "findings", "limitations", "future")

# Type-scale ladder (references/typography-spec.md). pt is the only unit here:
# PowerPoint stores `a:rPr/@sz` in 1/100 pt, so the ladder is defined directly
# in pt and needs no px conversion. Four tiers only:
# L1 = 26 page title, L2 = 22 group heading, L3 = 20 body bullet,
# L4 = 14 caption / footnote / sub-item (the floor).
DEFAULT_TIERS_PT = (26.0, 22.0, 20.0, 14.0)
TIER_ROLE = {
    1: "page title / section title",
    2: "group heading / module-block title",
    3: "body bullet / statement",
    4: "caption / footnote / citation / sub-item",
}
# Decorative separators a human rarely types into a slide. Separate phrases
# with a space instead (see references/outline-spec.md, section 0).
BANNED_SYMBOLS = "\u00b7\u30fb\u2219\u2022"   # middle dot, katakana dot, bullet op, bullet

# Per-panel captions ("Fig. 5a  <one-line title>", one under EVERY image, see
# layout-spec.md "subfigure label and title") are named `cap-<label>` by the
# generator. They sit INSIDE the figure area, far above the bottom caption
# band, so the y-position rule cannot classify them: recognise them by name.
# Caption text is scaffold, not body copy, and never acts as the page's
# "body block" in the hierarchy check.
CAPTION_PREFIXES = ("cap-", "Cap-", "CAP-", "caption-", "Caption-")


def is_caption(text_frame_owner_name):
    return text_frame_owner_name.startswith(CAPTION_PREFIXES)


def words(s):
    return len(re.findall(r"[A-Za-z0-9~\u2019'\-]+", s))


def _parse_map(raw, flag, cast):
    out = {}
    for part in (raw or "").split(","):
        part = part.strip()
        if not part:
            continue
        k, _, v = part.partition("=")
        try:
            out[int(k)] = cast(v)
        except ValueError:
            raise SystemExit(f"[ERROR] bad {flag} item: {part!r}")
    return out


def parse_int_map(raw, flag):
    return _parse_map(raw, flag, int)


def parse_float_map(raw, flag):
    return _parse_map(raw, flag, float)


def figure_area_of(profile, content_layout):
    for l in profile.get("layouts", []):
        if l.get("name") == content_layout:
            fa = l.get("figure_area")
            if fa:
                return fa["w"] * fa["h"]
    raise SystemExit(f"[ERROR] no figure_area for layout {content_layout!r}")


def template_sizes_pt(profile, min_pt, content_layout):
    """Sizes the *content* layout itself declares (>= the floor) stay legal.

    The content-page title, for instance, is a template decision and may sit
    off the content ladder — but the generator is allowed to keep it (on the
    bundled master it is 26 pt, which happens to be L1 exactly). Only the
    content layout is consulted: sizes that belong to the cover / section pages
    must not leak in as a loophole for content pages.
    """
    out = set()
    for l in profile.get("layouts", []):
        if content_layout and l.get("name") != content_layout:
            continue
        for ph in l.get("placeholders", []):
            sz = ph.get("size_pt_override")
            if sz and float(sz) >= min_pt - 0.01:
                out.add(round(float(sz), 2))
    return out


def type_scale_check(prs, exempt, ladder_pt, max_tiers, legal_extra, scaffold_y):
    """Walk explicit run sizes and judge them against the ladder.

    Only *explicit* `a:rPr/@sz` values are judged: an inherited size comes from
    the layout, i.e. it is the template's decision, not the generator's.
    `a:fld` runs (the slide-number field) are not exposed by python-pptx's
    `paragraph.runs`, so page numbers never enter this check.

    Each used size is classified, first match wins: on-ladder -> its tier;
    template-declared -> "template" (not a content tier); below the floor ->
    `below minimum tier`; anything else -> `off-ladder size`. A 12.75 pt
    citation is therefore reported as "too small" rather than merely
    "off-ladder", which is the actionable message.
    """
    ladder = sorted({round(float(x), 2) for x in ladder_pt}, reverse=True)
    level_of = {pt: i + 1 for i, pt in enumerate(ladder)}
    legal = set(ladder)
    extra = {round(float(x), 2) for x in legal_extra}
    min_pt = min(ladder)

    lines, fails = [], []
    deck_sizes = {}
    for i, slide in enumerate(prs.slides, 1):
        if i in exempt:
            lines.append(f"   slide{i}: exempt")
            continue
        per_shape, used, info = [], set(), []
        body_used = set()          # sizes the generator chose for body copy
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            is_title = (sh.is_placeholder
                        and sh.placeholder_format.idx == 0)
            sizes, txt = [], ""
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    txt += (r.text or "")
                    if r.font.size and (r.text or "").strip():
                        pt = round(r.font.size.pt, 2)
                        sizes.append(pt)
                        used.add(pt)
                        deck_sizes[pt] = deck_sizes.get(pt, 0) + 1
            if sizes:
                if not is_title:
                    body_used.update(sizes)
                per_shape.append((sh.name, sorted(set(sizes)), is_title))
                if is_caption(sh.name):
                    # A caption row must never become the page's "body block":
                    # it is L4 by definition, so it would drag every real
                    # bullet into a false "inverted hierarchy" verdict.
                    continue
                info.append({
                    "name": sh.name, "max": max(sizes), "words": words(txt),
                    "top": sh.top / 914400,
                    "is_title": is_title,
                })

        if not used:
            lines.append(f"   slide{i}: no explicit sizes (all inherited) — OK")
            continue

        tiers, off, below, mixed = set(), [], [], []
        for pt in sorted(used):
            if pt in legal:
                # A size used *only* by the title placeholder is the template's
                # decision, not one of the generator's tiers: the bundled
                # master's 26 pt title happens to equal L1, and counting it
                # would spend one of the three tiers on the template.
                if pt in body_used:
                    tiers.add(level_of[pt])
                else:
                    tiers.add("T")
            elif pt in extra:                          # template-declared
                tiers.add("T")
            elif pt < min_pt - 1e-6:
                below.append(pt)
            else:
                off.append(pt)
        tag = "+".join(
            [f"L{n}" for n in sorted(t for t in tiers if t != "T")]
            + (["template"] if "T" in tiers else []))
        parts = ", ".join(f"{pt:g}pt" for pt in sorted(used))
        line = f"   slide{i}: {parts}  -> {tag or '-'}"
        content_tiers = [t for t in tiers if t != "T"]
        bad = off or below or len(content_tiers) > max_tiers
        if bad:
            line += "  FAIL"
        if below:
            line += (f" [below {min_pt:g}pt minimum: "
                     f"{', '.join(f'{p:g}pt' for p in below)}]")
            fails.append(f"slide{i} below minimum tier")
        if off:
            line += f" [off-ladder: {', '.join(f'{p:g}pt' for p in off)}]"
            fails.append(f"slide{i} off-ladder size")
        if len(content_tiers) > max_tiers:
            line += f" [>{max_tiers} tiers]"
            fails.append(f"slide{i} too many tiers")
        # One textbox, one size (typography-spec section three rule 8): a
        # frame that needs two tiers is two statements — split it into two
        # frames, each carrying a single explicit size. The title placeholder
        # is exempt (template-owned); inherited-only frames never reach here
        # because `sizes` only collects explicit run sizes.
        for name, sizes, sh_is_title in per_shape:
            if len(sizes) > 1 and not sh_is_title:
                mixed.append(
                    f"        ! {name[:24]}: "
                    f"{', '.join(f'{p:g}pt' for p in sizes)} in ONE textbox "
                    f"— split it: one textbox carries one size")
                fails.append(f"slide{i} mixed sizes in one textbox")
        if mixed:
            line += "  FAIL [mixed sizes in one textbox]"
        # Hierarchy sanity: a short label must never sit on a *lower* tier than
        # the page's main body block — that is the "group heading smaller than
        # its own bullets" defect. The floor tier is exempt: L4 is by
        # definition sub-item/caption and is legitimately smaller than body
        # copy, as is anything in the caption band.
        inverted = []
        body = max((d for d in info if d["words"] >= 6),
                   key=lambda d: d["words"], default=None)
        if body and body["max"] in level_of:
            bl = level_of[body["max"]]
            for d in info:
                if d is body or d["is_title"] or d["words"] > 4 or not d["max"]:
                    continue
                if d["top"] >= scaffold_y:
                    continue
                if d["max"] not in level_of:
                    continue
                lv = level_of[d["max"]]
                if lv > bl and lv != len(ladder):
                    inverted.append(
                        f"        ! {d['name'][:22]} ({d['max']:g}pt, "
                        f"{d['words']}w) is SMALLER than the body block "
                        f"{body['name'][:22]} ({body['max']:g}pt, "
                        f"{body['words']}w) — inverted hierarchy")
                    fails.append(f"slide{i} inverted hierarchy")

        if inverted:
            line += "  FAIL [inverted hierarchy]"
        lines.append(line)
        for name, sizes, _t in per_shape:
            pretty = ", ".join(
                f"{pt:g}pt" + (f"(L{level_of[pt]})" if pt in level_of else "")
                for pt in sizes)
            lines.append(f"        {name[:24]:26s} {pretty}")
        lines.extend(mixed)
        lines.extend(inverted)
    
    lines.append("   deck-wide sizes: " + ", ".join(
        f"{pt:g}pt x{deck_sizes[pt]}" for pt in sorted(deck_sizes)))
    return lines, fails


# ============================================================ fonts
# Hard rule from the user (v10): ONE LATIN FAMILY, ONE CJK FAMILY —
# Calibri + 黑体. Explicit typefaces on slides/layouts/master must be those
# (or a theme token, which resolves to them through the theme's major/minor
# font). Two mechanically checkable corollaries:
#
#   - every theme*.xml must declare the same pairing (a template that only
#     fixed theme1 leaves theme2/theme3 on 等线, and every layout wired to
#     those renders in the wrong family);
#   - a `latin` slot must never point at the CJK token (+mj-ea / +mn-ea):
#     that renders ASCII in 黑体, which reads as "wrong font, unnameable".
#
# Colour is deliberately NOT gated here: v10 keeps the four-step type scale
# and the grey-capable PAPER_MINIMAL palette (see theme_override.py).
LEGAL_LATIN = {"Calibri", "Calibri Light"}
LEGAL_EA = {"黑体", "SimHei", "Heiti SC", "Hei", "SimHei Regular"}


def _read_theme_palette(z):
    """slot -> hex, plus the latin/ea faces declared by major/minor font."""
    pal, faces = {}, {}
    for name in sorted(n for n in z.namelist()
                       if n.startswith("ppt/theme/") and n.endswith(".xml")):
        root = etree.fromstring(z.read(name))
        cs = root.find(f".//{A}clrScheme")
        if cs is not None:
            for slot in cs:
                local = etree.QName(slot).localname
                for sub in slot:
                    sl = etree.QName(sub).localname
                    if sl == "srgbClr":
                        pal.setdefault(local, "#" + (sub.get("val") or "").upper())
                    elif sl == "sysClr":
                        pal.setdefault(local, "#" + (sub.get("lastClr") or "000000").upper())
        for group in ("majorFont", "minorFont"):
            holder = root.find(f".//{A}{group}")
            if holder is None:
                continue
            for kind in ("latin", "ea"):
                el = holder.find(f"{A}{kind}")
                if el is not None:
                    faces.setdefault(f"{group}/{kind}",
                                     el.get("typeface") or "")
    return pal, faces


def font_check(z, parts):
    lines, fails = [], []
    _pal, faces = _read_theme_palette(z)
    if faces:
        bad_faces = []
        for k, v in sorted(faces.items()):
            legal = LEGAL_LATIN if k.endswith("latin") else LEGAL_EA
            if v and v not in legal:
                bad_faces.append(f"{k}={v!r}")
        lines.append("   theme: " + ", ".join(
            f"{k.replace('Font/', ' ')}={v or '(empty)'}"
            for k, v in sorted(faces.items())))
        if bad_faces:
            lines.append("   ! theme font scheme off-rule: " + ", ".join(bad_faces))
            fails.append("theme font family")
        elif len({faces.get(k) for k in faces if k.endswith("latin")}) > 1 or \
                len({faces.get(k) for k in faces if k.endswith("ea")}) > 1:
            lines.append("   ! major and minor declare different families — "
                         "this deck must read as ONE family")
            fails.append("theme major/minor mismatch")

    total, bad = 0, []
    for name in parts:
        root = etree.fromstring(z.read(name))
        short = name.rsplit("/", 1)[-1]
        # Same downward walk as color_check: never map an element back to its
        # shape by id(), lxml proxies are recycled and ids collide.
        for sp in list(root.iter(f"{P}sp")) + list(root.iter(f"{P}pic")) + \
                list(root.iter(f"{P}graphicFrame")):
            cnv = sp.find(f".//{P}cNvPr")
            owner = (cnv.get("name") if cnv is not None else None) or "?"
            for el in sp.iter():
                local = etree.QName(el).localname
                if local not in ("latin", "ea"):
                    continue
                tf = el.get("typeface") or ""
                if local == "latin" and tf in ("+mj-ea", "+mn-ea"):
                    # latin slot wired to the EAST-ASIAN token: ASCII text then
                    # renders in 黑体, which shows up as "the font looks wrong
                    # but every explicit typeface is legal".
                    total += 1
                    bad.append((short, owner, "latin", tf,
                                "slot points at the CJK font"))
                    continue
                if not tf or tf.startswith("+"):
                    continue          # theme token / inherit: correct by design
                total += 1
                legal = LEGAL_LATIN if local == "latin" else LEGAL_EA
                if tf in legal:
                    continue
                bad.append((short, owner, local, tf, ""))
    lines.append(f"   {total} explicit typeface(s) on slides/layouts/master")
    for short, owner, kind, tf, why in bad:
        want = "Calibri" if kind == "latin" else "黑体"
        lines.append(f"   ! {short} / {owner[:26]}: {kind}={tf!r} -> must be "
                     f"{want} (or a +mj-/+mn- theme token)"
                     + (f"  [{why}]" if why else ""))
        fails.append(f"{short} font {tf}")
    if not bad and not any("font" in f for f in fails):
        lines.append("   OK: one latin family (Calibri) + one CJK family (黑体)")
    return lines, fails


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--profile", default="template_profile.json")
    ap.add_argument("--pages", type=int, default=0,
                    help="expected page count (0 = auto, from the deck)")
    ap.add_argument("--exempt", default="",
                    help="1-based page numbers exempt from figure/word rules")
    ap.add_argument("--content-layout", default="\u5c0f\u6807\u9898\u5185\u9875")
    ap.add_argument("--scaffold-y", type=float, default=6.85,
                    help="text boxes at or below this y (in) are citation/scaffold")
    ap.add_argument("--max-body-words", type=int, default=30,
                    help="cap on BODY words per page: title placeholder, `cap-*` "
                         "caption rows and everything at y >= --scaffold-y are "
                         "counted separately and never enter this number")
    ap.add_argument("--body-words-override", default="",
                    help="per-slide body-word exceptions, e.g. \"3=40,7=40\"; "
                         "use only for core-evidence pages marked 'Budget: 40' "
                         "in the outline (see SKILL.md 3.2)")
    ap.add_argument("--min-figure-pct", type=float, default=50.0,
                    help="reference line for figure share; below it prints WARN "
                         "and does NOT fail the deck unless --figure-blocking")
    ap.add_argument("--figure-pct-override", default="",
                    help="per-slide accepted figure-share thresholds, e.g. "
                         "\"12=42,7=38\"; declare pages whose panels cannot fill "
                         "50%% for a scientific reason (tall/thin, near-square, "
                         "many necessary panels). Reported as 'accepted'")
    ap.add_argument("--figure-blocking", action="store_true",
                    help="treat below-threshold figure share as a failure again "
                         "(opt-in; default is diagnostic/warn-only)")
    ap.add_argument("--tiers", default=",".join(f"{x:g}" for x in DEFAULT_TIERS_PT),
                    help="type-scale ladder in pt, largest first "
                         "(default %s; see references/typography-spec.md)"
                         % ",".join(f"{x:g}" for x in DEFAULT_TIERS_PT))
    ap.add_argument("--max-tiers-per-page", type=int, default=3,
                    help="fail a page that uses more distinct tiers than this")
    ap.add_argument("--no-type-scale", action="store_true",
                    help="skip the type-scale check")
    ap.add_argument("--no-font-check", action="store_true",
                    help="skip the Calibri + 黑体 font check")
    ap.add_argument("--banned-symbols", default=BANNED_SYMBOLS,
                    help="decorative separator characters that must not appear "
                         "in slide copy (empty string disables the check)")
    args = ap.parse_args()

    exempt = {int(x) for x in args.exempt.split(",") if x.strip()}
    word_cap = parse_int_map(args.body_words_override, "--body-words-override")
    figure_cap = parse_float_map(args.figure_pct_override,
                                 "--figure-pct-override")
    Profile = json.loads(Path(args.profile).read_text(encoding="utf-8"))
    FA = figure_area_of(Profile, args.content_layout)
    fails = []
    warns = []

    z = zipfile.ZipFile(args.deck)
    parts = sorted([n for n in z.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                   key=lambda s: int(re.findall(r"\d+", s)[0]))
    expected = args.pages or len(parts)
    print(f"[pages] {len(parts)} (plan {expected}) -> "
          f"{'OK' if len(parts) == expected else 'FAIL'}")
    if len(parts) != expected:
        fails.append("page count")

    for n in parts:
        root = etree.fromstring(z.read(n))
        tree = root.find(f".//{P}cSld/{P}spTree")
        sp = tree.findall(f"{P}sp")
        pic = tree.findall(f"{P}pic")
        txt = sum(1 for s in sp
                  if "".join(t.text or "" for t in s.findall(f".//{A}t")).strip())
        fld = root.findall(f".//{A}fld")
        num = re.findall(r"\d+", n)[0]
        ok = len(sp) > 0 and txt > 0 and len(fld) == 1
        print(f"   slide {num}: shapes={len(sp):2d} pics={len(pic)} "
              f"text_shapes={txt:2d} slidenum_field={len(fld)} "
              f"{'OK' if ok else 'FAIL'}")
        if not ok:
            fails.append(f"slide {num} structure")

    print("\n[per page]")
    prs = Presentation(args.deck)
    cap_report = []          # (slide, pictures, caption rows) for the diagnostic
    band_pages = []          # (slide, row count) for the bottom `cap-band`
    author_pages = []        # (slide, title text) — author page, fixed title
    for i, slide in enumerate(prs.slides, 1):
        fig = body = scaffold = title = 0.0
        pics = caps = 0
        title_txt = ""
        other_txt = []
        for sh in slide.shapes:
            w, h = sh.width / 914400, sh.height / 914400
            is_title = sh.is_placeholder and sh.placeholder_format.idx == 0
            txt = (sh.text_frame.text or "").strip() if sh.has_text_frame else ""
            if sh.__class__.__name__ == "Picture" or sh.name.startswith("Rectangle"):
                fig += w * h
                pics += 1
                continue
            if not txt:
                continue
            n = words(txt)
            if is_caption(sh.name):
                # per-panel caption ("Fig. 5a  <title>") — a caption, not body
                # copy, wherever it sits on the page
                caps += 1
                scaffold += n
            elif is_title:
                title += n
                title_txt = " ".join(txt.split())
            elif sh.top / 914400 >= args.scaffold_y:
                scaffold += n
                other_txt.append(txt)
            else:
                body += n
                other_txt.append(txt)
        # Author page detection: the page carries author metadata (role label or
        # the photo placeholder) but is not the cover. Its title is a FIXED
        # label, "Authors" (SKILL.md 3.8) — never the paper title, a name, or a
        # question.
        _low = " ".join(other_txt).lower()
        if ("corresponding author" in _low or "photo placeholder" in _low
                or "co-first author" in _low):
            author_pages.append((i, title_txt))
        pct = fig / FA * 100
        line = (f"   slide{i}: figure={fig:5.1f}in2 ({pct:5.1f}% of graphics area) "
                f"| title={title:2.0f}w body={body:2.0f}w scaffold={scaffold:2.0f}w "
                f"| caps={caps}/{pics}")
        if caps < pics:
            cap_report.append((i, pics, caps))
        band = next((sh for sh in slide.shapes if sh.name == "cap-band"
                     and sh.has_text_frame), None)
        if band is not None:
            rows = sum(1 for p in band.text_frame.paragraphs
                       if (p.text or "").strip())
            if rows:
                band_pages.append((i, rows))
        if i not in exempt:
            cap = word_cap.get(i, args.max_body_words)
            fcap = figure_cap.get(i, args.min_figure_pct)
            a = pct >= fcap
            b = body <= cap
            if a:
                fstat = "OK"
            elif i in figure_cap:
                fstat = f"accepted(<{fcap:.0f}%)"
            else:
                fstat = "WARN"
            line += f"  figure>{fcap:.0f}%:{fstat}" \
                    f"  body<={cap}w:{'OK' if b else 'FAIL'}"
            if not a:
                # Diagnostic by default: coverage is an indicator, not a gate.
                # Scientific completeness of the evidence list outranks it.
                if i in figure_cap:
                    # Human declared this page accepted; never silently re-fail it.
                    warns.append(f"slide{i} figure {pct:.1f}% (accepted)")
                elif args.figure_blocking:
                    fails.append(f"slide{i} figure share")
                else:
                    warns.append(f"slide{i} figure {pct:.1f}%")
            if not b:
                fails.append(f"slide{i} word budget")
        else:
            line += "  exempt"
        print(line)

    print("\n[figure coverage]  reference line "
          f"{args.min_figure_pct:.0f}% "
          f"({'BLOCKING' if args.figure_blocking else 'diagnostic, warn-only'})")
    if warns:
        for w in warns:
            print(f"   {w}")
        print("   -> judge these by hand: if every panel is necessary, the page")
        print("      already sits at the layout maximum, and the key evidence is")
        print("      readable, accept the page and note it in the delivery notes.")
        print("      Never drop/split evidence to raise this number.")
    else:
        print("   all non-exempt pages above the reference line")

    print("\n[panel captions]  every figure pane needs its own `cap-<label>` row "
          "(\"Fig. 5a  <title>\")")
    if cap_report:
        for num, pics, caps in cap_report:
            print(f"   slide{num}: {pics} figure pane(s) but {caps} caption row(s) "
                  f"— name them `cap-<label>` and put one under each image")
        print("   -> diagnostic: text under an image is a caption, not body copy,")
        print("      but an untitled subfigure leaves the audience guessing which")
        print("      claim that panel actually supports.")
    else:
        print("   every figure pane carries its own caption row")

    # The bottom cap-band holds two kinds of rows (SKILL.md 3.6): NAME
    # annotations (`Lgr5: intestinal stem-cell marker`), which are normal on
    # every page that introduces a term, and SUPPLEMENTARY rows (bottleneck /
    # reagent principle), which should be rare. Only the latter is worth
    # watching, so the read-out splits by row count: a single-row band is
    # taken as a name annotation, 2+ rows means a supplementary row slipped in.
    # Diagnostic only, never a gate.
    n_content = max(len(prs.slides) - len(exempt), 0)
    if band_pages:
        all_pages = [i for i, _ in band_pages]
        supp = [i for i, rows in band_pages if rows >= 2]
        share = len(supp) / n_content if n_content else 0.0
        print(f"\n[caption band]  {len(all_pages)} of {n_content} content page(s) "
              f"carry a `cap-band` — name annotations are normal at any rate")
        print("   slides: " + ", ".join(str(i) for i in all_pages))
        print(f"   supplementary rows (band with 2+ lines): {len(supp)} page(s), "
              f"{share:.0%} of content pages"
              + (f" — slides {', '.join(str(i) for i in supp)}" if supp else ""))
        if n_content >= 3 and share > 1 / 3:
            print("   -> too many pages carry a supplementary row. Ask of each")
            print("      one: does dropping it make the audience read the figure")
            print("      wrongly or overrate the claim? 'They'd know a bit less'")
            print("      does not qualify. The usual culprits are restating a")
            print("      bullet, restating what the axes/labels already show,")
            print("      and explaining textbook background.")
            print("      Diagnostic only — not a failure.")
        else:
            print("   -> supplementary rows within the expected share")

    print("\n[authors page]  the author page carries a FIXED title: `Authors` "
          "(SKILL.md 3.8)")
    if author_pages:
        bad = []
        for num, ttl in author_pages:
            ok = ttl.strip().lower() == "authors"
            print(f"   slide{num}: title={ttl!r} {'OK' if ok else 'WARN'}")
            if not ok:
                bad.append(num)
        if bad:
            print("   -> the author page is not an argument: the title is the")
            print("      label `Authors`. Not the paper title, not a name, not a")
            print("      question, no subtitle. Put the names, roles and topics")
            print("      inside the columns. Diagnostic only — not a failure.")
    else:
        print("   no author page in this deck (normal for a group meeting)")

    if not args.no_type_scale:
        ladder = tuple(sorted(
            (float(x) for x in args.tiers.split(",") if x.strip()),
            reverse=True))
        legal_extra = (template_sizes_pt(Profile, min(ladder), args.content_layout)
                       - set(ladder))      # on-ladder sizes need no exemption
        print("\n[type scale]  ladder: " + ", ".join(
            f"L{i+1}={pt:g}pt ({TIER_ROLE.get(i+1, '')})"
            for i, pt in enumerate(ladder)))
        if legal_extra:
            print("   template-declared sizes also legal: "
                  + ", ".join(f"{s:g}pt" for s in sorted(legal_extra)))
        tlines, tfails = type_scale_check(prs, exempt, ladder,
                                         args.max_tiers_per_page, legal_extra,
                                         args.scaffold_y)
        for l in tlines:
            print(l)
        fails.extend(tfails)

    ui_parts = sorted(
        n for n in z.namelist()
        if re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", n))

    if not args.no_font_check:
        print("\n[fonts]  rule: latin = Calibri, CJK = 黑体, one family only")
        flines, ffails = font_check(z, ui_parts)
        for l in flines:
            print(l)
        fails.extend(ffails)

    print("\n[copy style]")
    if args.banned_symbols:
        hits = []
        for i, slide in enumerate(prs.slides, 1):
            for sh in slide.shapes:
                if not sh.has_text_frame:
                    continue
                t = sh.text_frame.text or ""
                bad = [c for c in args.banned_symbols if c in t]
                if bad:
                    snippet = t.strip().replace("\n", " / ")[:60]
                    hits.append((i, bad, snippet))
        if hits:
            for i, bad, snippet in hits:
                print(f"   slide{i}: {' '.join(bad)}  in \"{snippet}\"  FAIL")
            fails.append("decorative separator symbol")
        else:
            print(f"   no decorative separators "
                  f"({' '.join(args.banned_symbols)})  OK")
    else:
        print("   disabled")

    print("\n[discussion elements]")
    joined = " ".join((s.text_frame.text or "")
                      for sl in prs.slides for s in sl.shapes
                      if s.has_text_frame).lower()
    for k in ELEMENTS:
        ok = k in joined
        print(f"   {k:12s}: {'present' if ok else 'MISSING'}")
        if not ok:
            fails.append(f"discussion {k}")

    print(f"\nRESULT: {len(fails)} failure(s)" +
          ("" if not fails else " -> " + ", ".join(fails)))
    if warns:
        print(f"        {len(warns)} figure-coverage warning(s) "
              "(diagnostic, not gating) -> " + ", ".join(warns))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
