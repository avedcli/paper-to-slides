#!/usr/bin/env python3
"""把一份 .pptx（模板或成品）的**字体层**归一化成「一拉丁族 + 一中文字族」。

规矩（用户直接指定，允许改母版）：

  **外文 Calibri、中文黑体** —— 主题的 major/minor 字体、master/layout 里
  残留的硬编码字族、项目符号字体，全部收敛到这两个族；`latin` 槽不得指向
  东亚令牌（`+mj-ea` / `+mn-ea`），否则 ASCII 文本被渲染成黑体。

本脚本做两件事（都只动母版层，不动配色、不动字号档位）：

  - ``apply_font_scheme``   改写**所有** ``theme*.xml`` 的 majorFont/minorFont
                            + master/layout 的 ``buFont``
  - ``fix_latin_slots``     把 master/layout 里 ``latin=+mj-ea / +mn-ea``
                            改回 ``+mj-lt / +mn-lt``

配色不在本脚本职责内：v10 沿用 ``PAPER_MINIMAL``（含灰阶），要换色板用
``theme_override.apply_theme_override``。

用法：

    python normalize_template.py assets/_default/template.pptx          # 就地改
    python normalize_template.py in.pptx -o out.pptx
    python normalize_template.py deck.pptx --dry-run                    # 只看报告

改完务必用 ``verify_deck.py`` 复核：``[fonts]`` 段应为 OK。
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation                                    # noqa: E402

from theme_override import (                                      # noqa: E402
    DEFAULT_EA, DEFAULT_LATIN, apply_font_scheme, fix_latin_slots,
)


def theme_report(path: Path) -> None:
    """读回刚存的文件，逐 theme 打印 major/minor 字体与错误槽计数。"""
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist()
                       if n.startswith("ppt/theme/") and n.endswith(".xml"))
        for n in names:
            s = z.read(n).decode("utf-8")
            faces = re.findall(
                r"<a:(major|minor)Font><a:latin typeface=\"([^\"]*)\"/>"
                r"(?:<a:ea typeface=\"([^\"]*)\"/>)?", s)
            summary = ", ".join(
                f"{k} latin={la or '(empty)'} ea={ea or '(empty)'}"
                for k, la, ea in faces)
            print(f"   {n.rsplit('/', 1)[-1]:14s} {summary}")
        bad = 0
        for n in z.namelist():
            if re.match(r"ppt/(slideLayouts|slideMasters|notesMasters)/.*\.xml$", n):
                s = z.read(n).decode("utf-8")
                bad += len(re.findall(r'<a:latin typeface="\+m[ij]-ea"', s))
        print(f"   latin slots still pointing at the CJK token: {bad}")


def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("-o", "--out", default="",
                    help="output path (default: overwrite the input)")
    ap.add_argument("--latin", default=DEFAULT_LATIN)
    ap.add_argument("--ea", default=DEFAULT_EA, help="东亚字体（默认 黑体）")
    ap.add_argument("--dry-run", action="store_true", help="report, do not save")
    args = ap.parse_args()

    src = Path(args.deck)
    if not src.exists():
        sys.exit(f"[ERROR] not found: {src}")
    prs = Presentation(str(src))
    print(f"[in]  {src}  ({len(prs.slides)} slide(s))")

    rep = apply_font_scheme(prs, latin=args.latin, ea=args.ea,
                            bullet_font=args.latin)
    print(f"[font] {len(rep['applied'])} typeface write(s)")
    for a in rep["applied"][:8]:
        print(f"        {a}")
    if len(rep["applied"]) > 8:
        print(f"        ... ({len(rep['applied']) - 8} more)")
    for w in rep["warnings"]:
        print(f"        WARN {w}")

    # latin 槽指到东亚字体（+mj-ea）是本项目踩过的坑：英文标题会被渲染成
    # 黑体，字距变宽而看不出是哪出问题。这里一并纠正。
    rep = fix_latin_slots(prs)
    print(f"[slot] {len(rep['applied'])} latin slot(s) repointed")
    for a in rep["applied"][:8]:
        print(f"        {a}")

    if args.dry_run:
        print("[dry-run] nothing written")
        return 0

    target = Path(args.out) if args.out else src
    prs.save(str(target))

    # ---- 读回校验：不信任内存状态，重新打开刚存的文件 ----------------
    print(f"[out] {target}")
    theme_report(target)
    return 0


if __name__ == "__main__":
    sys.exit(main())
