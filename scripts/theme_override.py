"""运行时主题色板覆写 — 让模板里 schemeClr 引用跟随用户传入的配色。

适用场景：模板用了 schemeClr accent1/accent3 等槽位，但用户希望每次按"学术极简"或
其他主题重新生成。本脚本直接在生成出来的 .pptx 中改写所有 theme*.xml 的 clrScheme，
所有 layout/master/slide 的 schemeClr 引用会随之解析到新色值。

注意：pptx 里通常有多份 theme*.xml（一份主主题 + 多份可能被 themes 文件夹下复用）。
本脚本会**全部**覆写，以避免出现"slide 渲染色 vs layout 引用色不一致"。

用法：
    from theme_override import apply_theme_override
    apply_theme_override(prs, PAPER_MINIMAL)   # prs 是 python-pptx Presentation 对象

色板字典的键名支持两种：
    - 标准槽位：dk1/lt1/dk2/lt2/accent1..accent6/hlink/folHlink
    - 自定义槽位：直接写 sRGB 字面值（不要带 #），例如 "custom1":"0B0F1A"

依赖：python-pptx + lxml（无需新增第三方库）
"""
from __future__ import annotations

import re
from copy import deepcopy

from pptx.oxml.ns import qn
from lxml import etree


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
A = f"{{{A_NS}}}"

# 标准主题色槽
STANDARD_SLOTS = {
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
}

# 学术极简（主黑白 + 极少点缀）默认色板
PAPER_MINIMAL = {
    "dk1":      "0B0F1A",   # 主文字：近黑
    "lt1":      "FFFFFF",   # 背景：白
    "dk2":      "0B0F1A",   # 第二深色：与 dk1 一致，强化极简
    "lt2":      "F5F6F8",   # 第二浅色：极浅灰
    "accent1":  "0B0F1A",   # 强调 = 主文字（极简：以主色作"强调"）
    "accent2":  "64748B",   # 次要灰（对比度 4.76:1，通过 WCAG AA）
    "accent3":  "94A3B8",   # 辅助灰（2.56:1，仅用于大字号次要标注）
    "accent4":  "1E40AF",   # 极少使用的强蓝色（仅在需要"突出关键结论"时使用）
    "accent5":  "64748B",   # 替代 #A6BEDA，与 accent2 一致作次要
    "accent6":  "D35D45",  # 占位（基本不用）
    "hlink":    "1E40AF",
    "folHlink": "64748B",
}


# ---- 字体（硬规则：外文 Calibri、中文黑体） ---------------------------------
DEFAULT_LATIN = "Calibri"
DEFAULT_EA = "黑体"
# 判定字形族是否"合规"时接受的名字（含各平台的同族别名）
LEGAL_LATIN = {"Calibri", "Calibri Light"}
LEGAL_EA = {"黑体", "SimHei", "SimHei Regular", "Heiti SC", "Hei"}


def _hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    h = hex_str.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"hex color must be 6 chars, got {hex_str!r}")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _set_srgb_clr(elem, r: int, g: int, b: int) -> None:
    """把元素内的色彩子元素清空，写入实色 srgbClr。"""
    for child in list(elem):
        # 任何已有的色彩子节点都删掉（srgbClr / schemeClr / sysClr / prstClr / scrgbClr）
        local = etree.QName(child.tag).localname
        if local in ("srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr"):
            elem.remove(child)
    srgb = etree.SubElement(elem, A + "srgbClr")
    srgb.set("val", f"{r:02X}{g:02X}{b:02X}")


def _ensure_slot(clr_scheme: etree._Element, slot: str) -> etree._Element:
    """找到或创建 <a:dk1>/<a:accent1> 等槽位元素（OXML 顺序敏感，按顺序追加）。"""
    target_tag = slot
    for child in clr_scheme:
        if etree.QName(child.tag).localname == target_tag:
            return child
    # 顺序：dk1 lt1 dk2 lt2 accent1..6 hlink folHlink
    order = ["dk1", "lt1", "dk2", "lt2",
             "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
             "hlink", "folHlink"]
    new = etree.Element(A + target_tag)
    # 找到应在的位置
    last_idx = -1
    for i, sib in enumerate(clr_scheme):
        local = etree.QName(sib.tag).localname
        if local in order:
            idx = order.index(local)
            if idx > order.index(target_tag):
                clr_scheme.insert(i, new)
                return new
        last_idx = i
    clr_scheme.append(new)
    return new


def _remove_existing_color_children(slot_elem: etree._Element) -> None:
    """删除槽位元素内的所有色彩子节点，准备写入新色。"""
    # OXML 规定槽位里有且仅有一个色彩子元素（srgbClr 或 sysClr 等）
    for child in list(slot_elem):
        local = etree.QName(child.tag).localname
        if local in ("srgbClr", "sysClr", "schemeClr", "prstClr", "scrgbClr"):
            slot_elem.remove(child)


def apply_theme_override(prs, palette: dict) -> dict:
    """把传入的 palette 写进 presentation 的所有 theme 部分，覆盖 clrScheme。

    返回值：{"applied": [...], "warnings": [...]} 便于 skill 端做报告。
    """
    applied: list[str] = []
    warnings: list[str] = []

    pkg = prs.part.package
    # 找到**所有** theme part（pptx 里通常多份，slideMaster/layout/slide 都引用主题）
    theme_parts = [
        p for p in pkg.iter_parts()
        if str(p.partname).endswith(".xml")
        and "theme" in str(p.partname).rsplit("/", 1)[-1]
    ]
    if not theme_parts:
        warnings.append("no theme*.xml found in package; colors not overridden")
        return {"applied": applied, "warnings": warnings}

    for theme_part in theme_parts:
        root = etree.fromstring(theme_part.blob)
        clr_scheme = root.find(f".//{A}clrScheme")
        if clr_scheme is None:
            warnings.append(f"{theme_part.partname}: no clrScheme; skipped")
            continue

        for slot, hex_value in palette.items():
            # 自定义槽位（不在标准列表内）— 直接忽略，避免破坏未知结构
            if slot not in STANDARD_SLOTS:
                if slot not in warnings:
                    warnings.append(f"slot {slot!r} is non-standard, skipped")
                continue
            try:
                r, g, b = _hex_to_rgb(hex_value)
            except ValueError as e:
                warnings.append(f"slot {slot} in {theme_part.partname}: {e}")
                continue

            slot_elem = _ensure_slot(clr_scheme, slot)
            _remove_existing_color_children(slot_elem)
            srgb = etree.SubElement(slot_elem, A + "srgbClr")
            srgb.set("val", f"{r:02X}{g:02X}{b:02X}")
            applied.append(f"{theme_part.partname}:{slot} -> #{r:02X}{g:02X}{b:02X}")

        # 写回 theme part
        new_blob = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        theme_part._blob = new_blob

    return {"applied": applied, "warnings": warnings}


def _theme_root(part):
    """取 theme part 的可写根节点。

    theme part 在不同 python-pptx 版本里可能是普通 ``Part``（只有 ``_blob``）
    或 ``XmlPart``（``blob`` 由 ``_element`` 反序列化，写 ``_blob`` 无效）。
    两条路径都要覆盖，否则"改完色/字但保存出来没变"。
    """
    el = getattr(part, "_element", None)
    if el is not None:
        return el, None
    return etree.fromstring(part.blob), part


def _commit_theme(part, root, owner) -> None:
    if owner is None:          # XmlPart：就地改，blob 自动反序列化
        return
    owner._blob = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)


def apply_font_scheme(prs, latin: str = DEFAULT_LATIN, ea: str = DEFAULT_EA,
                      bullet_font: str | None = None) -> dict:
    """把**所有** theme*.xml 的 majorFont / minorFont 统一成 Calibri + 黑体。

    为什么不只改一份主题：pptx 里 slideMaster / slideLayout / notesMaster 各自
    引用自己的 theme part。旧模板（如 Office 默认主题）在 theme2/theme3 里把
    latin 写成 ``等线`` / ``等线 Light``，只改 theme1 时，凡引用 theme2 的版式
    仍会用等线渲染——字体"看起来改好了"，实际没改全。

    同时清理：
      - ``<a:latin typeface="等线 Light" panose="..."/>`` 里的 panose 提示
        （换族后旧 panose 会让部分渲染器回退到相近但不同的字族）；
      - master/layout 里 ``<a:buFont typeface="Arial"/>`` 这类项目符号字体
        （``bullet_font`` 不为 None 时改成该值），让整份 deck 只有一个族。

    返回值：{"applied": [...], "warnings": [...]}。
    """
    applied: list[str] = []
    warnings: list[str] = []
    pkg = prs.part.package

    def set_typeface(el, value: str, kind: str) -> bool:
        if el is None:
            return False
        old = el.get("typeface") or ""
        if old == value:
            return False
        el.set("typeface", value)
        # 换族后旧的 panose 提示不再成立，删掉让渲染器按新族重算
        if el.get("panose"):
            del el.attrib["panose"]
        if el.get("pitchFamily"):
            del el.attrib["pitchFamily"]
        if el.get("charset"):
            del el.attrib["charset"]
        applied.append(f"{kind}: {old!r} -> {value!r}")
        return True

    theme_parts = [
        p for p in pkg.iter_parts()
        if str(p.partname).endswith(".xml") and "theme" in str(p.partname)
        and "theme" in str(p.partname).rsplit("/", 1)[-1]
    ]
    if not theme_parts:
        warnings.append("no theme*.xml found; font scheme skipped")
        return {"applied": applied, "warnings": warnings}

    for part in theme_parts:
        root, owner = _theme_root(part)
        pname = str(part.partname).rsplit("/", 1)[-1]
        touched = False
        for group in ("majorFont", "minorFont"):
            holder = root.find(f".//{A}{group}")
            if holder is None:
                continue
            for kind, value in (("latin", latin), ("ea", ea)):
                el = holder.find(f"{A}{kind}")
                if el is None:
                    el = etree.SubElement(holder, A + kind)
                    el.set("typeface", value)
                    applied.append(f"{pname}/{group}/{kind}: (created) -> {value!r}")
                    touched = True
                    continue
                if set_typeface(el, value, f"{pname}/{group}/{kind}"):
                    touched = True
        if touched:
            _commit_theme(part, root, owner)

    # master / layout 里的项目符号字体（buFont 不是"文字"，但换族要一致）
    if bullet_font:
        for part in pkg.iter_parts():
            pn = str(part.partname)
            if not (pn.endswith(".xml") and
                    ("/slideMasters/" in pn or "/slideLayouts/" in pn)):
                continue
            root = part._element
            n = 0
            for el in root.iter(f"{A}buFont"):
                if set_typeface(el, bullet_font, f"{pn}:buFont"):
                    n += 1
            if n:
                applied.append(f"{pn.rsplit('/', 1)[-1]}: {n} bullet font(s)")
    return {"applied": applied, "warnings": warnings}


def fix_latin_slots(prs) -> dict:
    """把 master/layout 里**指错槽**的拉丁字体改回来。

    `<a:latin typeface="+mj-ea"/>` 是错的：`+mj-ea` 指向主题的**东亚**字体
    （黑体），于是 ASCII 文本也按中文字体渲染——英文字距被拉宽、字重变板正，
    看起来"字体不对但查不出是哪种字体"。latin 槽只能指 `+mj-lt` / `+mn-lt`。

    返回值：{"applied": [...]}。
    """
    applied: list[str] = []
    pkg = prs.part.package
    for part in pkg.iter_parts():
        pn = str(part.partname)
        if not (pn.endswith(".xml") and
                ("/slideMasters/" in pn or "/slideLayouts/" in pn
                 or "/notesMasters/" in pn or "/notesSlides/" in pn)):
            continue
        root = part._element
        for el in root.iter(f"{A}latin"):
            old = el.get("typeface") or ""
            if old in ("+mj-ea", "+mn-ea"):
                new = "+mj-lt" if old == "+mj-ea" else "+mn-lt"
                el.set("typeface", new)
                applied.append(f"{pn.rsplit('/', 1)[-1]}: latin {old} -> {new}")
    return {"applied": applied}


def dump_palette(prs) -> dict:
    """从所有 theme 里读出当前色板，便于在覆写前/后做对比报告。返回主主题（theme1）色板。"""
    pkg = prs.part.package
    # 优先 theme1.xml；若没有则用第一个
    theme_part = None
    for part in pkg.iter_parts():
        if str(part.partname).endswith("theme1.xml"):
            theme_part = part
            break
    if theme_part is None:
        for part in pkg.iter_parts():
            if "theme" in str(part.partname) and str(part.partname).endswith(".xml"):
                theme_part = part
                break
    if theme_part is None:
        return {}
    root = etree.fromstring(theme_part.blob)
    clr_scheme = root.find(f".//{A}clrScheme")
    if clr_scheme is None:
        return {}
    out = {}
    for child in clr_scheme:
        local = etree.QName(child.tag).localname
        for sub in child:
            sl = etree.QName(sub.tag).localname
            if sl == "srgbClr":
                out[local] = "#" + sub.get("val", "")
            elif sl == "sysClr":
                out[local] = "sys:" + (sub.get("lastClr") or sub.get("val") or "")
            elif sl == "schemeClr":
                out[local] = "scheme:" + sub.get("val", "")
            else:
                out[local] = sl
    return out


def list_theme_parts(prs) -> list[str]:
    """返回包内所有 theme*.xml 的 partname，便于汇报。"""
    pkg = prs.part.package
    return sorted(
        str(p.partname) for p in pkg.iter_parts()
        if str(p.partname).endswith(".xml")
        and "theme" in str(p.partname).rsplit("/", 1)[-1]
    )


# 默认主题名：让模板里 "自定义 1" 这种占位名有一个一致的替代。
# P2-1 修复入口：模板导出时常把 clrScheme 的 name 写成 "自定义 1"，
# 无法辨识主题归属，skill 端在覆写色板时一并改名。
DEFAULT_THEME_NAME = "paper-minimalist"


def rename_theme(prs, name: str = DEFAULT_THEME_NAME) -> dict:
    """把所有 theme*.xml 的 <a:clrScheme name="..."/> 改成传入的名字。

    模板导出器（如 Keynote / iWork / HTML-to-pptx）经常把主题色板的 name
    属性写成 "自定义 1"、"Theme 1" 这类占位文本。OOXML 把它当成主题
    显示名（PowerPoint 主题面板会显示），改成可读的项目命名后便于回溯
    与跨 deck 复用。

    返回值：{"renamed": [...], "warnings": [...]}。
    """
    renamed: list[str] = []
    warnings: list[str] = []
    pkg = prs.part.package
    theme_parts = list_theme_parts(prs)
    if not theme_parts:
        warnings.append("no theme*.xml found; rename skipped")
        return {"renamed": renamed, "warnings": warnings}

    for partname in theme_parts:
        part = next(p for p in pkg.iter_parts() if str(p.partname) == partname)
        root = etree.fromstring(part.blob)
        clr_scheme = root.find(f".//{A}clrScheme")
        if clr_scheme is None:
            warnings.append(f"{partname}: no clrScheme; skipped")
            continue
        old = clr_scheme.get("name") or ""
        clr_scheme.set("name", name)
        new_blob = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        part._blob = new_blob
        renamed.append(f"{partname}: {old!r} -> {name!r}")
    return {"renamed": renamed, "warnings": warnings}


def apply_typography_normalization(prs, fonts: dict) -> dict:
    """在生成阶段把 layout / master 里硬编码的非主题字体统一换成主题引用。

    适用场景：模板 layout 里残留了像 ``Microsoft YaHei`` 这样的硬编码
    字体引用（典型如 P1-1 报告的 ``小标题内页 body#11`` latin="Microsoft YaHei"），
    与主题声明的字体冲突，会导致 macOS / Linux 渲染回退、跨机器不一致。

    字体归一化策略（传入 ``fonts`` dict）：

    - ``fonts["title_major_latin"]`` / ``fonts["title_major_ea"]``：标题占位符的
        拉丁/东亚字体，统一改成 ``<a:latin typeface="+mj-lt"/>`` /
        ``<a:ea typeface="+mj-ea"/>``。
    - ``fonts["body_minor_latin"]`` / ``fonts["body_minor_ea"]``：正文占位符
        的拉丁/东亚字体，统一改成 ``+mn-lt`` / ``+mn-ea``。
    - 字体以 ``+`` 开头的（如 ``+mj-ea``）是主题引用，本来就跟着主题走，
      **不**做替换。

    实现要点：python-pptx 的 XmlPart（slideLayout/slideMaster/theme 都是）
    ``blob`` 属性**总是**从 ``_element`` 重新序列化，``_blob`` 是私有
    缓存但 XmlPart 不会去读它。所以改写必须**就地修改 ``part._element``**，
    不能重新 ``etree.fromstring`` 再写 ``_blob``，否则保存出来的文件仍是
    原值。

    返回值：{"replaced": [...], "warnings": [...]}。
    """
    replaced: list[tuple[str, str, int]] = []  # (part, kind, count)
    warnings: list[str] = []

    title_maj_lt = fonts.get("title_major_latin", "+mj-lt")
    title_maj_ea = fonts.get("title_major_ea", "+mj-ea")
    body_min_lt = fonts.get("body_minor_latin", "+mn-lt")
    body_min_ea = fonts.get("body_minor_ea", "+mn-ea")

    pkg = prs.part.package
    # 遍历所有 slideMaster / slideLayout；slide 本身一般不再硬编码字体，
    # 但也扫一遍，防御性更好
    targets = [
        p for p in pkg.iter_parts()
        if str(p.partname).endswith(".xml")
        and (
            "/slideLayouts/slideLayout" in str(p.partname)
            or "/slideMasters/slideMaster" in str(p.partname)
        )
    ]
    if not targets:
        warnings.append("no slideLayout/slideMaster parts found; typography skipped")
        return {"replaced": replaced, "warnings": warnings}

    for part in targets:
        # XmlPart 的 _element 是 lxml 树；就地修改它即可（blob 会自动重新序列化）
        root = part._element
        n_replaced = 0
        for el in root.iter():
            if el.tag not in (f"{A}latin", f"{A}ea"):
                continue
            tf = el.get("typeface")
            if not tf:
                continue
            # 主题引用：保留不动
            if tf.startswith("+"):
                continue
            # 推断是标题还是正文：祖先链最近 p:sp 是不是 title 占位符
            is_title = _ancestor_has_title_placeholder(el)
            kind = etree.QName(el.tag).localname  # "latin" 或 "ea"
            if is_title:
                new_typeface = title_maj_lt if kind == "latin" else title_maj_ea
            else:
                new_typeface = body_min_lt if kind == "latin" else body_min_ea
            if tf == new_typeface:
                continue
            el.set("typeface", new_typeface)
            n_replaced += 1
        if n_replaced:
            replaced.append((str(part.partname), "typography", n_replaced))
    return {"replaced": replaced, "warnings": warnings}


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = f"{{{P_NS}}}"

# 标题型占位符的 type 值（OOXML 规定 title/ctrTitle 携带页面主标题文字）
TITLE_PH_TYPES = {"title", "ctrTitle"}


def _ancestor_has_title_placeholder(el) -> bool:
    """检查元素祖先链上是否经过 title 类型占位符。

    占位符元素 ``<p:ph type="title"/>`` 在 ``p:`` 命名空间下；本函数
    沿 ``el`` 的父链向上找最近的 ``<p:sp>``，再看里面有没有 ``<p:ph>``
    的 type 属性是 ``title`` / ``ctrTitle``。找到最近的 sp 就停：要么
    是标题占位符（命中），要么不是（不命中），不再继续往上找。
    """
    cur = el.getparent()
    while cur is not None:
        if cur.tag == f"{P}sp":
            for ph in cur.iter(f"{P}ph"):
                if ph.get("type") in TITLE_PH_TYPES:
                    return True
            return False
        cur = cur.getparent()
    return False