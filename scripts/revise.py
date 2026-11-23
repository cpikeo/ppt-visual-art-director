#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
既有 PPTX 的只读解析与定点修改：只动声明的对象，其余逐项保留。

用法：
    python scripts/revise.py inspect 原稿.pptx [--json inspect.json]          只读解析，不改原文件
    python scripts/revise.py 原稿.pptx edit.json 输出.pptx [--assets DIR] [--report revise.json] [--force]

为什么不是重新编译：母版、动画、切换、SmartArt、嵌入对象、备注在重建里回不来。

保留核对（每次写出前必跑；不通过则交付路径一个字节都不动，证据另存 .preserve-failed.pptx）：
  - 页数、母版、版式、主题字体；原稿含宏或格式不是 .pptx 直接拒绝
  - 逐页特性：隐藏、切换、动画、SmartArt、嵌入、音视频、备注、版式、底色
  - 未被点名的对象：XML（exclusive C14N）逐项一致，前后次序不变
  - 被点名的对象：标识与样式容器（nvXxPr、p:style）不变；超链接地址不变；只改授权的属性
  - 页面非形状部分（时序、切换、扩展、页级属性）与关系：不许丢失；新增关系只允许在改图的页上出现
  - 包内其余部件：规范化指纹一致
  - 输出必须是 .pptx，且包的内容类型与扩展名一致、无宏

退出码：0 写成功且保留核对通过；1 保留核对未通过（交付路径未写）；2 硬错误（未写文件）；3 运行期失败。
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import model as M  # noqa: E402
import ooxml as X  # noqa: E402

qn = M.qn
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

EMU = 914400.0
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DGM_REL = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds"


def inch(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 2)


def kind_of(shape) -> str:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    kind = getattr(shape, "shape_type", None)
    if kind == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT:
        return "ole"
    if kind == MSO_SHAPE_TYPE.MEDIA:
        return "media"
    if kind == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if kind == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if getattr(shape, "has_chart", False):
        return "chart"
    if getattr(shape, "has_table", False):
        return "table"
    if shape._element.find(".//" + DGM_REL) is not None:
        return "smartart"
    if getattr(shape, "is_placeholder", False):
        return "placeholder"
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        return "text"
    return "shape"


def run_styles(shape) -> Dict[str, Any]:
    """逐段逐 run 的字体事实。只看第一个 run 会漏掉混排、局部改色与中文回退。"""
    from pptx.enum.dml import MSO_COLOR_TYPE
    from pptx.oxml.ns import qn

    out: Dict[str, Any] = {"paragraphs": []}
    try:
        paragraphs = shape.text_frame.paragraphs
    except (AttributeError, ValueError):
        return out
    sizes, colors, fonts, weights = set(), set(), set(), set()
    for para in paragraphs:
        runs = []
        for run in para.runs:
            rPr = run._r.find(qn("a:rPr"))
            ea = None
            if rPr is not None:
                node = rPr.find(qn("a:ea"))
                ea = None if node is None else node.get("typeface")
            size = round(run.font.size.pt, 1) if run.font.size else None
            color = None
            try:
                value = run.font.color
                if value is not None and value.type == MSO_COLOR_TYPE.RGB:
                    color = f"#{value.rgb}"
                elif value is not None and value.type == MSO_COLOR_TYPE.SCHEME:
                    color = "(theme)"
            except (AttributeError, ValueError):
                pass
            spc = rPr.get("spc") if rPr is not None else None
            runs.append({"text": run.text, "size_pt": size, "font": run.font.name, "ea": ea,
                         "bold": bool(run.font.bold), "italic": bool(run.font.italic), "color": color,
                         "spacing": int(spc) / 100.0 if spc else None})
            if size:
                sizes.add(size)
            if color:
                colors.add(color)
            for name in (run.font.name, ea):
                if name:
                    fonts.add(name)
            weights.add(bool(run.font.bold))
        out["paragraphs"].append({
            "text": para.text.replace("\n", "⏎"),
            "align": str(para.alignment).split(" ")[0] if para.alignment is not None else None,
            "line_spacing": para.line_spacing,
            "runs": runs,
        })
    out.update({"sizes": sorted(sizes), "colors": sorted(colors), "fonts": sorted(fonts),
                "bold_mixed": len(weights) > 1})
    return out


def _digest(value: Any) -> str:
    """整份载荷的指纹。显示层会截断，核对层必须比全量，否则第 13 行之后的改动看不见。"""
    blob = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]


def describe(shape, depth: int = 0) -> Dict[str, Any]:
    kind = kind_of(shape)
    item: Dict[str, Any] = {
        "id": shape.shape_id,  # 同一页内稳定的标识；保留核对按它逐对象对应，不按可重名的名字
        "name": shape.name,
        "kind": kind,
        "box_in": [inch(shape.left), inch(shape.top), inch(shape.width), inch(shape.height)],
        "rotation": round(shape.rotation or 0, 1) if hasattr(shape, "rotation") else 0,
    }
    if kind in ("text", "placeholder"):
        # 全文保留：保留核对要比对文字，截断到 240 字等于放过第 241 字之后的任何改动。
        # 显示层在 to_markdown 里另行截断。
        item["text"] = shape.text_frame.text.replace("\n", "⏎")
        item["styles"] = run_styles(shape)
        item["links"] = run_links(shape)
        tf = shape.text_frame
        item["frame"] = {"ins": [inch(tf.margin_left), inch(tf.margin_top), inch(tf.margin_right),
                                 inch(tf.margin_bottom)],
                         "anchor": str(tf.vertical_anchor), "wrap": tf.word_wrap}
        if kind == "placeholder":
            try:
                item["ph"] = f"{shape.placeholder_format.type} #{shape.placeholder_format.idx}"
            except (AttributeError, ValueError, KeyError):
                item["ph"] = "unknown"
    elif kind == "picture":
        try:
            image = shape.image
            px_w, px_h = image.size
            width_in = (shape.width or 1) / EMU
            height_in = (shape.height or 1) / EMU
            item.update({"px": [px_w, px_h], "ext": image.ext,
                         "ppi": [round(px_w / width_in) if width_in else None,
                                 round(px_h / height_in) if height_in else None],
                         "crop": [round(shape.crop_left, 3), round(shape.crop_right, 3),
                                  round(shape.crop_top, 3), round(shape.crop_bottom, 3)],
                         "sha256": hashlib.sha256(image.blob).hexdigest()[:16]})
        except Exception:
            item["px"] = None
    elif kind == "chart":
        chart = shape.chart
        item["chart_type"] = str(chart.chart_type).split(" ")[0]
        item["title"] = chart.chart_title.text_frame.text if chart.has_title else None
        plots = []
        for plot in chart.plots:
            series = []
            for s in plot.series:
                try:
                    values = [round(float(v), 4) if not isinstance(v, tuple) else [round(float(x), 4) for x in v]
                              for v in s.values]
                except (TypeError, ValueError):
                    values = []
                series.append({"name": s.name, "values": values[:12], "count": len(values)})
            try:
                categories = [str(c) for c in list(plot.categories or [])]
            except (AttributeError, ValueError):
                categories = []
            plots.append({"series": series, "categories": categories[:12]})
        item["plots"] = plots
        item["has_legend"] = bool(chart.has_legend)
        item["payload"] = _digest([item["title"], [
            [[s.name, list(s.values)] for s in plot.series] + [list(plot.categories or [])]
            for plot in chart.plots]])
    elif kind == "table":
        table = shape.table
        grid = []
        for r in range(min(len(table.rows), 12)):
            grid.append([table.cell(r, c).text for c in range(len(table.columns))])
        full = [[table.cell(r, c).text for c in range(len(table.columns))] for r in range(len(table.rows))]
        item["payload"] = _digest(full)
        item.update({"rows": len(table.rows), "cols": len(table.columns), "grid": grid,
                     "col_widths_in": [round(col.width / EMU, 2) for col in table.columns]})
    elif kind == "group":
        item["children"] = [safe_describe(child, depth + 1) for child in shape.shapes]
    if kind in ("shape", "placeholder", "text"):
        from pptx.enum.dml import MSO_FILL_TYPE

        try:
            fill = shape.fill
            ftype = fill.type
            if ftype == MSO_FILL_TYPE.SOLID:
                item["fill"] = f"#{fill.fore_color.rgb}"
            elif ftype == MSO_FILL_TYPE.GRADIENT:
                item["fill"] = "gradient"
            elif ftype == MSO_FILL_TYPE.BACKGROUND:
                item["fill"] = "none"
            if shape.line.width:
                item["stroke_pt"] = round(shape.line.width.pt, 2)
            try:
                if shape.line.color and shape.line.color.type is not None:
                    item["stroke"] = f"#{shape.line.color.rgb}"
            except (AttributeError, ValueError):
                pass
        except Exception:
            pass
    return item


def safe_describe(shape, depth: int = 0) -> Dict[str, Any]:
    try:
        return describe(shape, depth)
    except Exception as exc:
        return {"name": getattr(shape, "name", "?"), "kind": "unparsed",
                "error": f"{type(exc).__name__}: {exc}"}


def slide_background(slide) -> Optional[str]:
    """页面底色：改稿时要知道"原来是什么场"，评审时要判断深浅场有没有成套换。"""
    from pptx.enum.dml import MSO_FILL_TYPE

    try:
        fill = slide.background.fill
        if fill.type == MSO_FILL_TYPE.SOLID:
            return f"#{fill.fore_color.rgb}"
        if fill.type is None or fill.type == MSO_FILL_TYPE.BACKGROUND:
            return None
        return str(fill.type).split(" ")[0].lower()
    except Exception:
        return None


def part_digest(name: str, data: bytes) -> str:
    """部件的规范化指纹。XML 去掉排版空白、关系按 Id 排序后做 C14N；其余部件逐字节。

    python-pptx 保存时会重新序列化每个 XML 部件（声明的引号、缩进、关系次序都可能变），
    逐字节比会把整包都判成"改过"。规范化只抹掉这些形式差异，语义上的改动仍然会被发现。
    """
    from lxml import etree

    if not name.endswith((".xml", ".rels")):
        return hashlib.sha256(data).hexdigest()
    try:
        root = etree.fromstring(data, etree.XMLParser(remove_blank_text=True, resolve_entities=False, no_network=True, load_dtd=False))
    except etree.XMLSyntaxError:
        return hashlib.sha256(data).hexdigest()
    if name.endswith(".rels"):
        children = sorted(list(root), key=lambda el: el.get("Id") or "")
        for el in children:
            root.remove(el)
            root.append(el)
    return hashlib.sha256(etree.tostring(root, method="c14n")).hexdigest()


def structure(pptx: Path) -> Dict[str, str]:
    """包内部件 → 字节哈希。改稿的"原始内容完整性"就靠它：没被点名的部件必须逐字节一致。

    只比页数、母版数、对象字段是不够的——母版主题、备注页、动画所在的 slide 部件
    被意外重写时，那些字段全都不变。
    """
    import zipfile

    with zipfile.ZipFile(Path(pptx)) as zf:
        return {name: part_digest(name, zf.read(name)) for name in zf.namelist()}


def slide_part_names(pptx: Path) -> List[str]:
    """按页序返回包内每页的 slide 部件名（原始名字）。

    页序由 sldIdLst 决定，部件名不一定是 slideN.xml 的 N：直接按文件名对页会对错页。
    这里直接读原包，不经过 python-pptx（它访问 slides 时会重新编号）。
    """
    import zipfile
    from lxml import etree

    p_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with zipfile.ZipFile(Path(pptx)) as zf:
        pres = M.parse_xml(zf.read("ppt/presentation.xml"))
        rels = M.parse_xml(zf.read("ppt/_rels/presentation.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    names = []
    for sld in pres.iter(f"{p_ns}sldId"):
        t = target[sld.get(f"{r_ns}id")].lstrip("/")
        names.append(t if t.startswith("ppt/") else "ppt/" + t)
    return names


def slide_flags(slide, objects: List[Dict[str, Any]]) -> Dict[str, Any]:
    xml = slide._element.xml

    def has(name: str) -> bool:
        def walk(items: List[Dict[str, Any]]) -> bool:
            return any(item.get("kind") == name or walk(item.get("children", [])) for item in items)

        return walk(objects)

    return {
        "hidden": slide._element.get("show") == "0",
        "transition": "<p:transition" in xml,
        "animation": "<p:timing" in xml,
        "smartart": has("smartart"),
        "ole": has("ole"),
        "media": has("media"),
        "notes": bool(slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip()),
    }


def collect_fonts(shapes, fonts: set, sizes: set) -> None:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for shape in shapes:
        try:
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        if run.font.size:
                            sizes.add(round(run.font.size.pt, 1))
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                if run.font.name:
                                    fonts.add(run.font.name)
                                if run.font.size:
                                    sizes.add(round(run.font.size.pt, 1))
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                collect_fonts(shape.shapes, fonts, sizes)
        except Exception:
            continue


def theme_fonts(prs) -> Dict[str, Any]:
    """主题字族：改稿时判断"原稿本来想用什么字"，run 上的字体名可能只是继承。"""
    from lxml import etree

    out: Dict[str, Any] = {}
    for master in prs.slide_masters:
        for rel in master.part.rels.values():
            if not rel.reltype.endswith("/theme"):
                continue
            try:
                root = M.parse_xml(rel.target_part.blob)
            except Exception:
                continue
            for tag in ("majorFont", "minorFont"):
                node = root.find(".//" + A + tag)
                if node is None:
                    continue
                latin, ea = node.find(A + "latin"), node.find(A + "ea")
                out[tag] = {"latin": None if latin is None else latin.get("typeface"),
                            "ea": None if ea is None else ea.get("typeface")}
    return out


def has_macros(pptx: Path, prs) -> bool:
    if pptx.suffix.lower() in (".pptm", ".ppsm", ".potm"):
        return True
    try:
        return any("vbaProject" in str(part.partname) for part in prs.part.package.iter_parts())
    except Exception:
        return False


def inspect_deck(pptx: Path) -> Dict[str, Any]:
    from pptx import Presentation

    pptx = Path(pptx)
    prs = Presentation(str(pptx))
    out: Dict[str, Any] = {
        "file": str(pptx),
        "canvas_in": [inch(prs.slide_width), inch(prs.slide_height)],
        "slides": len(prs.slides),
        "masters": len(prs.slide_masters),
        "layouts": sum(len(m.slide_layouts) for m in prs.slide_masters),
        "theme_fonts": theme_fonts(prs),
        "macros": has_macros(pptx, prs),
        "pages": [],
        "errors": [],
    }
    fonts: set = set()
    sizes: set = set()
    for index, slide in enumerate(prs.slides):
        try:
            objects = [safe_describe(s) for s in slide.shapes]
            flags = slide_flags(slide, objects)
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        except Exception as exc:
            out["errors"].append(f"第 {index + 1} 页解析失败：{type(exc).__name__}: {exc}")
            continue
        collect_fonts(slide.shapes, fonts, sizes)
        out["pages"].append({
            "index": index + 1,
            "layout": slide.slide_layout.name,
            "background": slide_background(slide),
            "flags": flags,
            "notes": notes,  # 全文保留：备注改动也要被保留核对看见
            "objects": objects,
        })
    out["fonts"] = sorted(fonts)
    xml_facts = slide_xml_facts(pptx, slide_part_names(pptx))
    if len(xml_facts) == len(out["pages"]):
        for page, facts in zip(out["pages"], xml_facts):
            page["xml"] = facts
    else:
        out["errors"].append("XML 级指纹与解析出的页数不一致")
    if sizes:
        distinct = sorted(sizes)
        out["type_scale"] = {"steps": len(distinct), "sizes": distinct, "min": min(distinct), "max": max(distinct)}
    out["risks"] = rebuild_risks(out)
    return out


def rebuild_risks(data: Dict[str, Any]) -> List[str]:
    """改稿路线的判断依据：这些特性在重建里回不来，必须先说明。"""
    risks: List[str] = []
    if data.get("macros"):
        risks.append("含宏（vbaProject）：重建会丢失，只能给逐页指令")
    for page in data["pages"]:
        flags = page["flags"]
        for key, label in (("smartart", "SmartArt"), ("ole", "嵌入对象"), ("media", "音视频"),
                           ("transition", "切换"), ("animation", "动画")):
            if flags.get(key):
                risks.append(f"第 {page['index']} 页含{label}：重建后不存在")
    if data.get("layouts", 0) > 11 or data.get("masters", 0) > 1:
        risks.append(f"母版 {data['masters']} 套、版式 {data['layouts']} 个：自定义母版样式不会随重建保留")
    return risks


def to_markdown(data: Dict[str, Any]) -> str:
    lines = [f"# {Path(data['file']).name}", "",
             f"画布 {data['canvas_in'][0]} × {data['canvas_in'][1]} in｜{data['slides']} 页｜"
             f"母版 {data['masters']}｜版式 {data['layouts']}"]
    scale = data.get("type_scale")
    if scale:
        lines.append(f"字号 {scale['steps']} 档（{scale['min']:g}–{scale['max']:g}pt）")
    if data.get("fonts"):
        lines.append(f"字体 {data['fonts']}")
    theme = data.get("theme_fonts") or {}
    for tag in ("majorFont", "minorFont"):
        if theme.get(tag):
            lines.append(f"主题{tag} 西文 {theme[tag]['latin']}／东亚 {theme[tag]['ea']}")
    if data.get("macros"):
        lines.append("**含宏**")

    def emit(obj: Dict[str, Any], depth: int = 0) -> None:
        pad = "  " * depth
        box = obj.get("box_in") or [None, None, None, None]
        head = f"{pad}- {obj['kind']:11} {obj['name'][:26]:26} [{box[0]},{box[1]},{box[2]}×{box[3]}]"
        if obj.get("rotation"):
            head += f" {obj['rotation']}°"
        detail = ""
        if obj["kind"] in ("text", "placeholder"):
            styles = obj.get("styles", {})
            sizes = "/".join(f"{s:g}" for s in styles.get("sizes", []))
            detail = f"{sizes}pt {'/'.join(styles.get('fonts', []))} 「{(obj.get('text') or '')[:120]}」"
            if obj.get("ph"):
                detail = f"{obj['ph']}｜" + detail
        elif obj["kind"] == "picture":
            detail = f"{obj.get('px')} {obj.get('ppi')}ppi {obj.get('ext')}"
        elif obj["kind"] == "chart":
            plots = obj.get("plots", [])
            bits = [" / ".join(f"{s['name']}({s['count']})" for s in plot["series"]) for plot in plots]
            detail = f"{obj.get('chart_type')} {len(plots)} 层｜" + "；".join(bits)
            if obj.get("title"):
                detail += f"｜题「{obj['title'][:28]}」"
        elif obj["kind"] == "table":
            detail = f"{obj['rows']}×{obj['cols']}  首行 {obj.get('grid', [[]])[0]}"
        elif obj["kind"] == "group":
            detail = f"{len(obj.get('children', []))} 个子对象"
        elif obj["kind"] == "unparsed":
            detail = f"解析失败：{obj.get('error')}"
        lines.append(f"{head} {detail}".rstrip())
        for child in obj.get("children", []):
            emit(child, depth + 1)

    for page in data["pages"]:
        flags = [name for name, on in page["flags"].items() if on]
        head = f"\n## 第 {page['index']} 页 · {page['layout']}"
        if flags:
            head += f" · {','.join(flags)}"
        lines.append(head)
        if page.get("background"):
            lines.append(f"底色 {page['background']}")
        if page["notes"]:
            lines.append(f"备注：{page['notes'][:100]}")
        for obj in page["objects"]:
            emit(obj)
    if data.get("risks"):
        lines.append("\n## 重建会失去什么")
        lines.extend(f"- {item}" for item in data["risks"])
    if data.get("errors"):
        lines.append("\n## 解析异常")
        lines.extend(f"- {item}" for item in data["errors"])
    return "\n".join(lines)




OPS = ("text", "box", "style", "fill", "line", "order", "image", "chart_data", "delete", "background")
# 这些字段一旦出现在操作里，就说明它有意改样式；只改字时一个都不该有。
STYLE_KEYS = ("font", "font_ea", "font_size", "bold", "color", "letter_spacing_pt",
              "align", "line_spacing", "pad_in", "valign")
EMU = M.EMU

# 操作级属性白名单：每个操作只接受表中的键，拼错或不适用的键直接拒绝（不静默忽略）。
COMMON_KEYS = {"op", "page", "target", "target_regex", "target_index", "note"}
OP_KEYS = {
    "text": {"content", "runs", "space_after_pt"} | (set(STYLE_KEYS) - {"pad_in", "valign"}),
    "box": {"x", "y", "w", "h"},
    "style": set(STYLE_KEYS),
    "fill": {"color", "alpha"},
    "line": {"color", "width_pt", "dash", "hide"},
    "order": {"to"},
    "image": {"image_id"},
    "chart_data": {"data"},
    "delete": set(),
    "background": {"color"},
}
# 每种操作授权改动的对象属性（对象指纹的顶层键）。段落样式与文本框字段再按声明收窄。
GRANT_TOP = {
    "text": {"text"}, "box": {"box_in"}, "style": set(), "fill": {"fill"},
    "line": {"stroke", "stroke_pt"}, "image": {"sha256", "px", "ppi", "crop"},
    "chart_data": {"plots", "payload"}, "order": set(), "delete": set(), "background": set(),
}
STYLE_SIG = {"font": "font", "font_ea": "font_ea", "font_size": "font_size", "bold": "bold",
             "color": "color", "align": "align", "line_spacing": "line_spacing",
             "letter_spacing_pt": "letter_spacing_pt"}
STYLE_FRAME = {"pad_in": "ins", "valign": "anchor"}


# ---------------------------------------------------------------- 目标解析


def slide_at(prs, page: Any) -> Any:
    if not isinstance(page, int) or isinstance(page, bool) or not (1 <= page <= len(prs.slides)):
        raise M.BadOp(f"page 必须是 1..{len(prs.slides)} 的整数，收到 {page!r}")
    return prs.slides[page - 1]


def resolve(shape_list, op: Dict[str, Any]) -> Any:
    """按名称、正则或 z 序定位对象。名称多处重名时必须改用 target_index，不替人猜。"""
    name = op.get("target")
    regex = op.get("target_regex")
    index = op.get("target_index")
    if index is not None:
        if not isinstance(index, int) or not (0 <= index < len(shape_list)):
            raise M.BadOp(f"target_index 必须是 0..{len(shape_list) - 1} 的整数，收到 {index!r}")
        return shape_list[index]
    if regex is not None:
        hits = [s for s in shape_list if re.search(str(regex), s.name)]
    elif name is not None:
        hits = [s for s in shape_list if s.name == str(name)]
    else:
        raise M.BadOp("缺少 target / target_regex / target_index")
    if not hits:
        listing = ", ".join(f"{i}:{s.name}" for i, s in enumerate(shape_list))[:400]
        raise M.BadOp(f"找不到目标 {name or regex!r}；本页对象为 {listing}")
    if len(hits) > 1:
        raise M.BadOp(f"目标 {name or regex!r} 命中 {len(hits)} 个对象，改用 target_index："
                      f"{[s.name for s in hits][:6]}")
    return hits[0]


# ---------------------------------------------------------------- 文本


def run_style(run) -> Dict[str, Any]:
    """一个 run 的**显式**样式；没声明的字段是 None。

    None 表示"继承自版式或主题"，不能替它编一个值——真实稿的标题多半是 44pt 主题样式、
    run 上什么都没有，编成 18pt 等于把设计改坏，而保留核对不会响。
    """
    from pptx.enum.dml import MSO_COLOR_TYPE

    color = None
    try:
        if run.font.color and run.font.color.type == MSO_COLOR_TYPE.RGB:
            color = f"#{run.font.color.rgb}"
    except (AttributeError, ValueError):
        color = None
    rPr = run._r.find(qn("a:rPr"))
    ea = None
    if rPr is not None and rPr.find(qn("a:ea")) is not None:
        ea = rPr.find(qn("a:ea")).get("typeface")
    spc = rPr.get("spc") if rPr is not None else None
    return {"font": run.font.name, "font_ea": ea,
            "size": run.font.size.pt if run.font.size else None,
            "bold": run.font.bold, "color": color,
            "spacing": int(spc) / 100.0 if spc else None}


def paragraph_styles(shape) -> List[Dict[str, Any]]:
    """逐段：该段的 run 样式列表，加上段落级对齐与行距。

    只取第一段的第一个 run 是"改一处、坏一处"的根源：多段标题、副题、灰色眉题
    各有各的字级，用第一个 run 的样式重写全部，等于把整块文字压成一种。
    """
    out: List[Dict[str, Any]] = []
    for para in shape.text_frame.paragraphs:
        out.append({
            "runs": [run_style(run) for run in para.runs],
            "align": str(para.alignment).split(" ")[0].lower() if para.alignment is not None else None,
            "line_spacing": para.line_spacing,
        })
    return out


def base_style(shape) -> Dict[str, Any]:
    """兼容旧调用：第一段第一个 run 的样式与段落属性。"""
    paras = paragraph_styles(shape)
    if not paras:
        return {}
    first = paras[0]["runs"][0] if paras[0]["runs"] else {}
    return {**first, "align": paras[0]["align"], "line_spacing": paras[0]["line_spacing"]}


def _first(*values: Any) -> Any:
    """第一个非 None 的值。False 与 0 都是合法值，不能用 or 链代替。"""
    for value in values:
        if value is not None:
            return value
    return None


def _write_run(run, piece: Dict[str, Any], op: Dict[str, Any], base: Dict[str, Any]) -> None:
    """run 级 > 操作级 > 原稿该 run 的显式值；三层都没有就不写，让继承继续生效。"""
    X.set_run_attrs(
        run,
        font=_first(piece.get("font"), op.get("font"), base.get("font")),
        font_ea=_first(piece.get("font_ea"), op.get("font_ea"), base.get("font_ea")),
        size=_first(piece.get("font_size"), op.get("font_size"), base.get("size")),
        color=_first(piece.get("color"), op.get("color"), base.get("color")),
        bold=_first(piece.get("bold"), op.get("bold"), base.get("bold")),
        spacing=_first(piece.get("letter_spacing_pt"), op.get("letter_spacing_pt"), base.get("spacing")),
    )


def _mixed(runs: List[Dict[str, Any]]) -> bool:
    """同一段里是否存在两种以上不同的 run 样式（混排）。"""
    sigs = {(r.get("size"), r.get("bold"), r.get("font"), r.get("font_ea"), r.get("color")) for r in runs}
    return len(sigs) > 1


def apply_style(shape, op: Dict[str, Any]) -> None:
    """只改声明的属性。一个 color 操作不该顺手把字号、字重、字距也重写一遍。"""
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.util import Pt

    tf = shape.text_frame
    for para in tf.paragraphs:
        if op.get("align"):
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[op["align"]]
        if op.get("line_spacing"):
            para.line_spacing = float(op["line_spacing"])
        for run in para.runs:
            # 显式声明 font 就是要换字族：连东亚字族一起写，否则中文仍留在原字族上。
            X.set_run_attrs(run, font=op.get("font"), font_ea=op.get("font_ea") or op.get("font"),
                            size=op.get("font_size"), color=op.get("color"),
                            bold=op.get("bold"), spacing=op.get("letter_spacing_pt"))
    if op.get("valign"):
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[op["valign"]]
    if op.get("pad_in") is not None:
        pad = float(op["pad_in"])
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(pad * 72)


# ---------------------------------------------------------------- 几何与外观


def apply_box(shape, op: Dict[str, Any]) -> None:
    from pptx.util import Inches
    for key, attr in (("x", "left"), ("y", "top"), ("w", "width"), ("h", "height")):
        if op.get(key) is not None:
            setattr(shape, attr, Inches(float(op[key])))


def apply_fill(shape, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from lxml import etree

    color = op.get("color")
    if color is None:
        raise M.BadOp("fill 需要 color")
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(color).lstrip("#"))
    if op.get("alpha") is not None:
        alpha = max(0.0, min(1.0, float(op["alpha"])))
        srgb = shape.fill.fore_color._xFill.find(qn("a:srgbClr"))
        if srgb is not None:
            for old in srgb.findall(qn("a:alpha")):
                srgb.remove(old)
            etree.SubElement(srgb, qn("a:alpha")).set("val", str(int(round(alpha * 100000))))


def apply_line(shape, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    if op.get("color"):
        shape.line.color.rgb = RGBColor.from_string(str(op["color"]).lstrip("#"))
    if op.get("width_pt") is not None:
        shape.line.width = Pt(float(op["width_pt"]))
    if op.get("dash"):
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        shape.line.dash_style = {"solid": MSO_LINE_DASH_STYLE.SOLID, "dash": MSO_LINE_DASH_STYLE.DASH,
                                 "dot": MSO_LINE_DASH_STYLE.ROUND_DOT, "dash_dot": MSO_LINE_DASH_STYLE.DASH_DOT,
                                 "long_dash": MSO_LINE_DASH_STYLE.LONG_DASH}[str(op["dash"])]
    if op.get("hide"):
        shape.line.fill.background()


def apply_order(shape, op: Dict[str, Any]) -> None:
    """层级也是版面：把色块压到底、把标签提到顶，是编辑里最常见的一步。"""
    to = op.get("to")
    node = shape._element
    parent = node.getparent()
    if to == "back":
        parent.remove(node)
        parent.insert(2, node)  # spTree: nvGrpSpPr, grpSpPr, 之后才是第一个对象
    elif to == "front":
        parent.remove(node)
        parent.append(node)
    else:
        raise M.BadOp("order 的 to 取 back 或 front")


def apply_image(shape, path: Path) -> None:

    if not path.is_file():
        raise M.BadOp(f"素材不存在：{path}")
    image_part, rId = shape.part.get_or_add_image_part(str(path))
    blip = shape._element.blipFill.find(qn("a:blip"))
    if blip is None:
        raise M.BadOp("目标不是图片对象")
    blip.set(qn("r:embed"), rId)


def apply_chart(shape, op: Dict[str, Any]) -> None:
    from pptx.chart.data import CategoryChartData

    if not getattr(shape, "has_chart", False):
        raise M.BadOp("目标不是图表对象")
    chart = shape.chart
    if len(chart.plots) > 1:
        raise M.BadOp("该图有多个图层（组合图）：改数据会同时打乱两层格式，请整稿重建")
    data = op.get("data") or {}
    cats, series = data.get("categories") or [], data.get("series") or []
    if not cats or not series:
        raise M.BadOp("chart_data 需要 data.categories 与 data.series")
    existing = [s.name for s in chart.plots[0].series]
    if len(series) != len(existing):
        raise M.BadOp(f"系列数 {len(series)} ≠ 原图 {len(existing)}；加减系列会破坏配色与格式，请重建")
    # 类别顺序与名称决定颜色↔类别的绑定（扇区按 idx 着色、图例按 idx 取色）。
    # 只改数值是安全的；改类别会让同一份颜色落到别的类别上，扇区与图例从此对不上。
    existing_cats = [str(c) for c in chart.plots[0].categories]
    if [str(c) for c in cats] != existing_cats:
        raise M.BadOp(f"chart_data 只改数值：类别 {cats} ≠ 原图 {existing_cats}；"
                      f"改类别顺序或名称会改变颜色↔类别的绑定，扇区与图例将错色，需要改类别请整稿重建")
    for i, s in enumerate(series):
        if len(s.get("values") or []) != len(cats):
            raise M.BadOp(f"系列「{s.get('name')}」值数 {len(s.get('values') or [])} ≠ 类别数 {len(cats)}")
        if s.get("name") and s["name"] != existing[i]:
            raise M.BadOp(f"第 {i + 1} 个系列名 {s['name']!r} ≠ 原图 {existing[i]!r}；改名会丢图例格式")
    payload = CategoryChartData()
    payload.categories = cats
    for s in series:
        payload.add_series(s["name"], [float(v) for v in s["values"]])
    chart.replace_data(payload)


def apply_background(slide, op: Dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor

    color = op.get("color")
    if color is None:
        raise M.BadOp("background 需要 color")
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(str(color).lstrip("#"))


# ---------------------------------------------------------------- 主流程


def _chart_parts(shape) -> set:
    """图表改数会重写自己的 chart 部件与内嵌工作簿：这两个部件因此可以变。

    只放行"这个图表自己的"部件，别的图表被顺手改写照样会报——放行整片
    `ppt/charts/` 就等于关掉了这条检查。
    """
    out: set = set()
    try:
        part = shape.chart.part
    except (AttributeError, ValueError):
        return out
    out.add(str(part.partname).lstrip("/"))
    for rel in getattr(part, "rels", {}).values():
        target = getattr(rel, "target_part", None)
        name = str(getattr(target, "partname", "") or "").lstrip("/")
        if name:
            out.add(name)
    return out


def check_keys(kind: Any, op: Dict[str, Any], where: str) -> None:
    """操作级白名单：未知操作、拼错或不适用的键一律拒绝。静默忽略的结果是报告写着"已改"、文件里什么都没变。"""
    if kind not in OP_KEYS:
        raise M.BadOp(f"{where}: 未知操作 {kind!r}；可用 {', '.join(OPS)}")
    extra = sorted(set(op) - COMMON_KEYS - OP_KEYS[kind])
    if extra:
        raise M.BadOp(f"{where}: {kind} 不接受 {extra}；该操作只认 {sorted(OP_KEYS[kind])}")


def empty_grant() -> Dict[str, Any]:
    return {"top": set(), "sig": set(), "frame": set(), "text": False}


def grant_for(kind: str, op: Dict[str, Any]) -> Dict[str, Any]:
    """一次操作授权改动的属性：顶层键、声明过的段落样式字段、声明过的文本框字段。"""
    declared = set(op) | {k for run in op.get("runs") or [] for k in run}
    grant = empty_grant()
    grant["top"] = set(GRANT_TOP[kind])
    grant["sig"] = {STYLE_SIG[k] for k in declared if k in STYLE_SIG}
    grant["frame"] = {STYLE_FRAME[k] for k in declared if k in STYLE_FRAME}
    grant["text"] = kind == "text"
    return grant


def merge_grant(into: Dict[str, Any], extra: Dict[str, Any]) -> None:
    for key in ("top", "sig", "frame"):
        into[key] |= extra[key]
    into["text"] = into["text"] or extra["text"]


def run_ops(prs, ops: List[Dict[str, Any]], asset_root: Path, log: List[Dict[str, Any]],
            rep: M.Report) -> Tuple[Dict[Tuple[int, int], Dict[str, Any]], set, set, set]:
    """逐条执行操作。对象以 (页, shape id) 为键；返回每个被点名对象的授权清单。"""
    grants: Dict[Tuple[int, int], Dict[str, Any]] = {}
    bg_touched: set = set()
    deleted: set = set()
    movable: set = set()
    for n, op in enumerate(ops, 1):
        where = f"op{n}"
        kind = op.get("op")
        check_keys(kind, op, where)
        slide = slide_at(prs, op.get("page"))
        if kind == "background":
            apply_background(slide, op)
            bg_touched.add(int(op["page"]))
            log.append({"op": n, "kind": kind, "page": op["page"]})
            continue
        shape = resolve(slide.shapes, op)
        key = (int(op["page"]), shape.shape_id)
        merge_grant(grants.setdefault(key, empty_grant()), grant_for(kind, op))
        if kind == "text":
            if op.get("content") is None:
                raise M.BadOp(f"{where}: text 需要 content")
            write_text(shape, op, rep, int(op["page"]), shape.name)
        elif kind == "box":
            if not any(op.get(k) is not None for k in ("x", "y", "w", "h")):
                raise M.BadOp(f"{where}: box 至少给 x/y/w/h 之一")
            apply_box(shape, op)
        elif kind == "style":
            if not getattr(shape, "has_text_frame", False):
                raise M.BadOp(f"{where}: {shape.name} 没有文本可改样式")
            apply_style(shape, op)
        elif kind == "fill":
            apply_fill(shape, op)
        elif kind == "line":
            apply_line(shape, op)
        elif kind == "image":
            if not op.get("image_id"):
                raise M.BadOp(f"{where}: image 需要 image_id")
            path = M.resolve_asset(asset_root, str(op["image_id"]))
            if path is None:
                raise M.BadOp(f"{where}: 找不到素材 {op['image_id']!r}（素材根 {asset_root}）")
            apply_image(shape, path)
        elif kind == "chart_data":
            apply_chart(shape, op)
            movable |= _chart_parts(shape)
        elif kind == "order":
            apply_order(shape, op)
        elif kind == "delete":
            shape._element.getparent().remove(shape._element)
            deleted.add(key)
        log.append({"op": n, "kind": kind, "page": op["page"], "target": shape.name,
                    "note": op.get("note", "")})
    return grants, bg_touched, deleted, movable


def ink_box(obj: Dict[str, Any]) -> Optional[Tuple[float, float, float, float]]:
    """文字实际落点的大致范围：声明框常比字宽，用声明框判遮挡会天天误报。

    估法与本包其他处一致（中文一字一 em、西文约 0.55 em），宽不超过声明框。
    """
    box = obj.get("box_in")
    if not box or None in box:
        return None
    if obj.get("kind") != "text":
        return tuple(box)  # type: ignore[return-value]
    paras = (obj.get("styles") or {}).get("paragraphs") or []
    if not paras:
        return tuple(box)  # type: ignore[return-value]
    widest, tallest = 0.0, 0.0
    for para in paras:
        size = max([r.get("size_pt") or 0 for r in para.get("runs") or []] or [0])
        if not size:
            continue
        text = para.get("text") or ""
        latin = sum(1 for ch in text if ord(ch) < 128)
        cjk = max(len(text) - latin, 0)
        widest = max(widest, (cjk + latin * 0.55) * size / 72)
        tallest += size * 1.35 / 72
    if not widest:
        return tuple(box)  # type: ignore[return-value]
    return (box[0], box[1], min(box[2], max(widest, 0.2)), min(box[3], max(tallest, 0.1)))


def post_check(after: Dict[str, Any], f: M.Report) -> None:
    """改完再按同一套版面规则过一遍：挪过的文字有没有被上层色块盖住、有没有出画布。

    新建稿由 compile 的 E-OVERLAP 守；改动稿不走 plan，所以这里补同一道闸。
    """
    w_in, h_in = after.get("canvas_in") or (13.333, 7.5)
    for page in after["pages"]:
        objects = page["objects"]
        boxes = [(o.get("name"), o.get("kind"), o.get("box_in"), o.get("fill"), o) for o in objects]
        for i, (name, kind, box, _, obj) in enumerate(boxes):
            if not box or None in box:
                continue
            if box[0] < -0.02 or box[1] < -0.02 or box[0] + box[2] > w_in + 0.02 or box[1] + box[3] > h_in + 0.02:
                if kind in ("text", "chart", "table"):
                    f.soft("W-BOUND", f"第 {page['index']} 页 {name}", "改动后超出画布，会被裁掉")
                continue
            if kind != "text":
                continue
            glyph = ink_box(obj) or box
            for up_name, up_kind, up_box, up_fill, _ in boxes[i + 1:]:
                if not up_box or None in up_box:
                    continue
                opaque = bool(up_fill) or up_kind in ("picture", "chart", "table")
                if not opaque:
                    continue
                ix, iy = M._intersection(glyph, up_box)
                area = ix * iy
                if area >= 0.35 and area >= 0.3 * max(glyph[2] * glyph[3], 1e-6):
                    f.soft("W-COVER", f"第 {page['index']} 页 {name}",
                           f"文字落在上层的 {up_name} 上（约 {area:.1f} in²）；挪位置或调层级")
                    break


_SIG_FIELDS = ("font_size", "bold", "font", "font_ea", "color", "align", "line_spacing", "letter_spacing_pt")


def _text_sig(obj: Dict[str, Any]) -> Optional[List[set]]:
    """逐段样式签名：每段用到的（字号、字重、西文/东亚字族、墨色、对齐、行距）集合。

    只看第一个 run 会放过"第二段被压成 32pt"这种最常见的漂移——副题、眉题、括注
    在真实稿里各有各的字级。按段比对而不是按 run 位置比对：换行与增删 run 是内容
    编辑，压平整段样式才是改设计。段落级属性也进签名：只改字的操作顺手改掉对齐，
    是同一类漂移，只是发生在上层。
    """
    paras = (obj.get("styles") or {}).get("paragraphs") or []
    out: List[set] = []
    for para in paras:
        align = para.get("align")
        spacing = para.get("line_spacing")
        out.append({(r.get("size_pt"), r.get("bold"), r.get("font"), r.get("ea"), r.get("color"),
                     align, spacing, r.get("spacing")) for r in para.get("runs") or []})
    return out or None


def _mask(sigs: set, allowed: frozenset) -> set:
    """把"本操作合法声明过"的字段抹平，只留没被授权的位。"""
    return {tuple(None if _SIG_FIELDS[i] in allowed else v for i, v in enumerate(s)) for s in sigs}


def _style_drift(old: Optional[List[set]], new: Optional[List[set]], allowed: frozenset = frozenset()):
    """只改字的操作不许给任何一段换一套样式：新段落用到的样式必须是它原来就有的。

    允许"少用"（内容删掉了那句灰色括注），不允许"新出现"（14pt 副题被压成 32pt，
    或继承的 44pt 被写死成 18pt）。同一目标上另有 `style` 操作显式声明过的字段不参与
    判定——否则"先换字、再改字号"这种正常写法会被自己的闸拦住。
    """
    if not old or not new:
        return None
    olds = [_mask(s, allowed) for s in old]
    news = [_mask(s, allowed) for s in new]

    def ref_for(index: int) -> Optional[set]:
        # 段落数变多时用相邻段作参照：新增的那一段没有"原来的样式"可比。
        for j in range(min(index, len(olds) - 1), -1, -1):
            if olds[j]:
                return olds[j]
        for j in range(index + 1, len(olds)):
            if olds[j]:
                return olds[j]
        return None

    for i, sigs in enumerate(news):
        if not sigs:
            continue
        ref = ref_for(i)
        if not ref:
            continue
        stray = sigs - ref
        if stray:
            return i + 1, sorted(ref, key=str), sorted(stray, key=str)
    return None


def _short(value: Any, n: int = 80) -> str:
    text = repr(value)
    return text if len(text) <= n else text[:n] + "…"


def _object_drift(obj: Dict[str, Any], twin: Dict[str, Any], grant: Optional[Dict[str, Any]],
                  where: str, f: M.Report) -> None:
    """对象的改动是否都落在授权之内。未被点名的对象（grant 为 None）必须逐项一致。"""
    top = grant["top"] if grant else set()
    for key in sorted((set(obj) | set(twin)) - {"styles", "frame"}):
        if key not in top and obj.get(key) != twin.get(key):
            f.hard("E-PRESERVE", where,
                   f"未授权的属性「{key}」被改动：{_short(obj.get(key))} → {_short(twin.get(key))}")
    granted_frame = grant["frame"] if grant else set()
    fb, fa = obj.get("frame") or {}, twin.get("frame") or {}
    bad = sorted(k for k in set(fb) | set(fa) if k not in granted_frame and fb.get(k) != fa.get(k))
    if bad:
        f.hard("E-PRESERVE", where, f"未授权的文本框属性被改动：{bad}")
    allowed = frozenset(grant["sig"]) if grant else frozenset()
    sig_b, sig_a = _text_sig(obj), _text_sig(twin)
    if grant and grant["text"]:
        drift = _style_drift(sig_b, sig_a, allowed)
        if drift:
            para_no, ref, stray = drift
            f.hard("E-PRESERVE", where,
                   f"只改文字的操作用改了样式：第 {para_no} 段出现原稿没有的"
                   f"（字号, 粗体, 西文字族, 东亚字族, 墨色, 对齐, 行距, 字距）={stray}（该段原有 {ref}）")
    elif [_mask(s, allowed) for s in sig_b or []] != [_mask(s, allowed) for s in sig_a or []]:
        f.hard("E-PRESERVE", where, "未授权的段落样式被改动（字号/字重/字族/墨色/对齐/行距/字距）")


FLAG_LABEL = {"hidden": "隐藏", "transition": "切换", "animation": "动画", "smartart": "SmartArt",
              "ole": "嵌入对象", "media": "音视频", "notes": "备注"}


def compare(before: Dict[str, Any], after: Dict[str, Any], grants: Dict[Tuple[int, int], Dict[str, Any]],
            bg_touched: set, deleted: set, f: M.Report, image_pages: frozenset = frozenset()) -> Dict[str, Any]:
    """保留核对：没有任何操作授权改动的东西，必须与原稿逐项一致。这是"没有顺手改坏别处"的证据。"""
    for key, label in (("slides", "页数"), ("masters", "母版数"), ("layouts", "版式数"), ("macros", "宏")):
        if before.get(key) != after.get(key):
            f.hard("E-PRESERVE", label, f"{before.get(key)} → {after.get(key)}")
    if before.get("theme_fonts") != after.get("theme_fonts"):
        f.hard("E-PRESERVE", "主题字体", "改动前后不一致")

    removed = 0
    for i, b_page in enumerate(before["pages"]):
        where_page = f"第 {i + 1} 页"
        if i >= len(after["pages"]):
            f.hard("E-PRESERVE", where_page, "页面丢失")
            continue
        a_page = after["pages"][i]
        page_no = i + 1
        compare_xml(b_page, a_page, page_no,
                    {oid: g for (p, oid), g in grants.items() if p == page_no},
                    {oid for (p, oid) in deleted if p == page_no},
                    page_no in bg_touched, page_no in image_pages, f)
        # 逐页特性（隐藏、切换、动画、SmartArt、嵌入、音视频、备注）没有任何操作授权改它们
        for flag in sorted(set(b_page["flags"]) | set(a_page["flags"])):
            bv, av = b_page["flags"].get(flag), a_page["flags"].get(flag)
            if bv != av:
                f.hard("E-PRESERVE", where_page, f"页面特性「{FLAG_LABEL.get(flag, flag)}」被改动：{bv} → {av}")
        if b_page.get("notes") != a_page.get("notes"):
            f.hard("E-PRESERVE", where_page, "备注文字被改动")
        if b_page.get("layout") != a_page.get("layout"):
            f.hard("E-PRESERVE", where_page, f"版式被改动：{b_page.get('layout')} → {a_page.get('layout')}")
        if i + 1 not in bg_touched and b_page.get("background") != a_page.get("background"):
            f.hard("E-PRESERVE", where_page,
                   f"未点名的底色被改动：{b_page.get('background')} → {a_page.get('background')}")

        b_ids = [o.get("id") for o in b_page["objects"]]
        a_ids = [o.get("id") for o in a_page["objects"]]
        if len(set(b_ids)) != len(b_ids) or len(set(a_ids)) != len(a_ids):
            f.hard("E-PRESERVE", where_page, "对象标识重复，无法逐对象核对；请先在 PowerPoint 中修复重复标识")
            continue
        b_objs = dict(zip(b_ids, b_page["objects"]))
        a_objs = dict(zip(a_ids, a_page["objects"]))
        for oid, obj in b_objs.items():
            where = f"{where_page} {obj.get('name')}"
            if oid not in a_objs:
                removed += 1
                if (i + 1, oid) not in deleted:
                    f.hard("E-PRESERVE", where, "对象消失，但没有任何 delete 操作点名它")
                continue
            _object_drift(obj, a_objs[oid], grants.get((i + 1, oid)), where, f)
        for oid, obj in a_objs.items():
            if oid not in b_objs:
                f.hard("E-PRESERVE", f"{where_page} {obj.get('name')}", "出现原稿没有的对象；定点编辑不新增任何对象")
        # 没被点名的对象，相对前后次序必须不变（order 操作只授权它移动的那一个）
        kept_b = [oid for oid in b_ids if (i + 1, oid) not in grants and oid in a_objs]
        kept_a = [oid for oid in a_ids if (i + 1, oid) not in grants and oid in b_objs]
        if kept_b != kept_a:
            f.hard("E-PRESERVE", where_page, "未点名对象的层级（前后次序）被改动")
    return {"deleted": removed, "touched": len(grants)}


def compare_parts(before: Dict[str, str], after: Dict[str, str], before_slides: List[str],
                  after_slides: List[str], touched_pages: set, f: M.Report,
                  movable: Optional[set] = None) -> Dict[str, Any]:
    """包内部件比对。slide 按页比对（页序不等于部件名，python-pptx 保存时还会重新编号）；
    其余部件按名字逐字节比对。被点名页的 slide 与其 rels、改数图表自己的部件可以变，别的不许变。"""
    import re as _re

    allowed_new = ("ppt/media/", "ppt/embeddings/")
    movable = movable or set()
    slide_re = _re.compile(r"^ppt/slides/(?:_rels/)?slide\d+\.xml(?:\.rels)?$")

    def page_bundle(parts: Dict[str, str], slide_name: str) -> Tuple[Any, Any]:
        rels = slide_name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        return parts.get(slide_name), parts.get(rels)

    changed, lost, added, moved = [], [], [], []
    for i in range(min(len(before_slides), len(after_slides))):
        if i + 1 in touched_pages:
            continue
        if page_bundle(before, before_slides[i]) != page_bundle(after, after_slides[i]):
            changed.append(f"第 {i + 1} 页的 slide 部件")
    for name, digest in before.items():
        if slide_re.match(name) or name == "[Content_Types].xml":
            continue
        if name not in after:
            lost.append(name)
        elif after[name] != digest:
            (moved if name in movable else changed).append(name)
    for name in after:
        if name in before or slide_re.match(name) or name.startswith(allowed_new):
            continue
        added.append(name)

    if lost:
        f.hard("E-PRESERVE", "包内部件", f"部件丢失：{lost[:6]}")
    if changed:
        f.hard("E-PRESERVE", "包内部件", f"未点名的部件被改写：{changed[:6]}")
    if added:
        f.hard("E-PRESERVE", "包内部件", f"出现原稿没有的部件：{added[:6]}")
    return {"parts": len(after), "changed": len(changed), "lost": len(lost), "added": len(added),
            "moved": moved}



# ---------------------------------------------------------------- 文本保真


def _first_rpr(paragraphs: List[Any]) -> Optional[Any]:
    for para in paragraphs:
        for run in para.findall(qn("a:r")):
            rpr = run.find(qn("a:rPr"))
            if rpr is not None:
                return rpr
    return None


def _link_runs(para: Any) -> int:
    count = 0
    for run in para.findall(qn("a:r")):
        rpr = run.find(qn("a:rPr"))
        if rpr is not None and rpr.find(qn("a:hlinkClick")) is not None:
            count += 1
    return count


def write_text(shape, op: Dict[str, Any], rep: M.Report, page: int, name: str) -> None:
    """换字，同时守住原有排版与富文本。

    - 段落属性（pPr：对齐、缩进、项目符号、段距）整份沿用原稿对应段。
    - run 属性（rPr：字号字重字色字族、字距、语言、下划线、超链接 hlinkClick）整份复制自原稿对应 run，
      再只覆盖操作声明过的字段；原 run 的其它属性一概保留。
    - 新段落 i 取原稿第 i 段（超出则取末段）；run 数一致时按位置继承，否则整段取该段首个 run 的样式并披露（W-RUNS）。
    - 原稿段落含超链接而 run 数对不上时拒绝：链接不能被悄悄丢掉或挪到别的字上。
    - 含域（a:fld，如页码字段）或软回车（a:br）的文本框不做 text 改写，免得把域与换行一起抹掉。
    """
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt

    tf = shape.text_frame
    body = tf._txBody
    if body.find(".//" + qn("a:fld")) is not None or body.find(".//" + qn("a:br")) is not None:
        raise M.BadOp(f"{name} 含域（如页码字段）或软回车 a:br，text 不能安全改写；"
                      f"请用 runs 写清每段，或整稿重建")

    content = str(op.get("content", ""))
    runs = op.get("runs")
    if runs and "".join(str(r.get("text", "")) for r in runs) != content:
        raise M.BadOp("runs 拼接与 content 不一致（含 \\n）")
    paras = M.paragraph_runs({"content": content, "runs": runs})
    if paras is None:
        raise M.BadOp("runs 分段与 content 对不上")

    old = paragraph_styles(shape)  # 逐段的 run 样式（用于披露混排与 W-RUNS 判断）
    old_paras = body.findall(qn("a:p"))
    templates = [{"ppr": copy.deepcopy(p.find(qn("a:pPr"))) if p.find(qn("a:pPr")) is not None else None,
                  "end": copy.deepcopy(p.find(qn("a:endParaRPr"))) if p.find(qn("a:endParaRPr")) is not None else None,
                  "rprs": [copy.deepcopy(r.find(qn("a:rPr"))) if r.find(qn("a:rPr")) is not None else None
                           for r in p.findall(qn("a:r"))],
                  "links": _link_runs(p)} for p in old_paras]
    fallback = _first_rpr(old_paras)
    fallback = copy.deepcopy(fallback) if fallback is not None else None
    fallback_style = old[0]["runs"][0] if old and old[0]["runs"] else {}

    for p in old_paras:
        body.remove(p)

    collapsed: List[int] = []
    for idx, (seg, pieces) in enumerate(paras):
        pieces = pieces or ([{"text": seg}] if seg else [])
        tmpl = templates[min(idx, len(templates) - 1)] if templates else {
            "ppr": None, "end": None, "rprs": [], "links": 0}
        if tmpl["links"] and len(pieces) != len(tmpl["rprs"]):
            raise M.BadOp(f"{name} 第 {idx + 1} 段含超链接，新内容的 run 数（{len(pieces)}）与原稿"
                          f"（{len(tmpl['rprs'])}）不一致，链接无法安全保留；请用 runs 逐段写清")
        src = old[idx] if idx < len(old) else (old[-1] if old else {"runs": []})
        src_runs = src.get("runs") or []
        if src_runs and _mixed(src_runs) and len(pieces) != len(src_runs):
            collapsed.append(idx + 1)

        para = tf.add_paragraph()
        if tmpl["ppr"] is not None:
            para._p.insert(0, copy.deepcopy(tmpl["ppr"]))
        for j, piece in enumerate(pieces):
            j_src = j if len(src_runs) == len(pieces) else 0
            base = src_runs[j_src] if j_src < len(src_runs) else (src_runs[0] if src_runs else fallback_style)
            base_rpr = tmpl["rprs"][j_src] if j_src < len(tmpl["rprs"]) else (tmpl["rprs"][0] if tmpl["rprs"] else fallback)
            run = para.add_run()
            run.text = piece.get("text", "")
            r_el = run._r
            stale = r_el.find(qn("a:rPr"))
            if stale is not None:
                r_el.remove(stale)
            if base_rpr is not None:
                r_el.insert(0, copy.deepcopy(base_rpr))
            _write_run(run, piece, op, base)

        if op.get("align"):
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[op["align"]]
        if op.get("line_spacing"):
            para.line_spacing = float(op["line_spacing"])
        if op.get("space_after_pt"):
            para.space_after = Pt(float(op["space_after_pt"]))
        if tmpl["end"] is not None:
            para._p.append(copy.deepcopy(tmpl["end"]))

    if collapsed:
        rep.soft("W-RUNS", f"第 {page} 页 {name}",
                 f"第 {collapsed} 段原有混排（同一段内字号/字色/字重不同），新内容 run 数与之不同，"
                 f"已统一为该段首个 run 的样式；要保留分区样式请用 runs 明确写出")


def run_links(shape) -> List[List[Any]]:
    """文字里每个超链接的位置与目标：[段, run, 地址]。改稿时链接不能被悄悄丢掉或挪走。"""
    out: List[List[Any]] = []
    try:
        paragraphs = list(shape.text_frame.paragraphs)
    except (AttributeError, ValueError):
        return out
    for pi, para in enumerate(paragraphs):
        for ri, run in enumerate(para.runs):
            link = run._r.find(qn("a:rPr") + "/" + qn("a:hlinkClick"))
            if link is None:
                continue
            rid = link.get(REL_ID)
            rel = shape.part.rels.get(rid) if rid else None
            target = rel.target_ref if rel is not None else None
            out.append([pi, ri, target or link.get("action") or "?"])
    return out


# ---------------------------------------------------------------- OOXML 级指纹


def _excl(el: Any) -> str:
    """exclusive C14N：只带本元素真正用到的命名空间，不受根节点 nsmap 变化影响。"""
    return hashlib.sha256(etree.tostring(el, method="c14n", exclusive=True)).hexdigest()


def slide_xml_facts(pptx: Path, part_names: List[str]) -> List[Dict[str, Any]]:
    """逐页的 XML 级指纹，直接读包内原始字节（不经 python-pptx，避免它的重新编号与序列化差异）。

    shapes：spTree 中每个形状的 id、整体指纹（xml）与标识容器指纹（meta：nvXxPr 与 p:style）；
    bg：页面底色节点；rest：slide 除 spTree 与 bg 之外的全部（时序、切换、扩展、页级属性）；
    rels：关系集合（Id、类型、目标、模式）。
    """
    out: List[Dict[str, Any]] = []
    with zipfile.ZipFile(pptx) as zf:
        names = set(zf.namelist())
        for name in part_names:
            root = M.parse_xml(zf.read(name))
            c_sld = root.find(P_NS + "cSld")
            tree = c_sld.find(P_NS + "spTree") if c_sld is not None else None
            shapes: List[Dict[str, Any]] = []
            for child in (tree if tree is not None else []):
                if etree.QName(child).localname in ("nvGrpSpPr", "grpSpPr") or not len(child):
                    continue
                nv = child[0]
                cnv = nv.find(P_NS + "cNvPr")
                sid = int(cnv.get("id")) if cnv is not None and (cnv.get("id") or "").isdigit() else None
                style = child.find(P_NS + "style")
                meta = _excl(nv) + (_excl(style) if style is not None else "")
                shapes.append({"id": sid, "xml": _excl(child),
                               "meta": hashlib.sha256(meta.encode()).hexdigest()})
            bg = c_sld.find(P_NS + "bg") if c_sld is not None else None
            rest = copy.deepcopy(root)
            r_csld = rest.find(P_NS + "cSld")
            if r_csld is not None:
                for el in list(r_csld):
                    if etree.QName(el).localname in ("spTree", "bg"):
                        r_csld.remove(el)
            rels_name = name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            rels = set()
            if rels_name in names:
                for r in M.parse_xml(zf.read(rels_name)):
                    rels.add((r.get("Id"), r.get("Type"), r.get("Target"), r.get("TargetMode") or ""))
            out.append({"shapes": shapes, "bg": _excl(bg) if bg is not None else None,
                        "rest": _excl(rest), "rels": [list(t) for t in sorted(rels)]})
    return out


def compare_xml(b_page: Dict[str, Any], a_page: Dict[str, Any], page_no: int,
                grants: Dict[int, Dict[str, Any]], deleted: set, bg_touched: bool,
                image_page: bool, f: M.Report) -> None:
    """XML 级保留核对：未授权的形状逐字一致；被点名的形状只许动授权的部分；页面其余结构与关系不许丢。"""
    where = f"第 {page_no} 页（XML）"
    bx, ax = b_page.get("xml"), a_page.get("xml")
    if not bx or not ax:
        f.hard("E-PRESERVE", where, "缺少 XML 级指纹，无法完成逐对象核对")
        return
    if bx["rest"] != ax["rest"]:
        f.hard("E-PRESERVE", where, "页面的非形状部分被改动（时序、切换、扩展或页级属性）")
    if not bg_touched and bx["bg"] != ax["bg"]:
        f.hard("E-PRESERVE", where, "页面底色节点（p:bg）被改动")
    b_rels = {tuple(r) for r in bx["rels"]}
    a_rels = {tuple(r) for r in ax["rels"]}
    if b_rels - a_rels:
        f.hard("E-PRESERVE", where, f"关系丢失 {len(b_rels - a_rels)} 条（如超链接、图片或图表引用），"
                                    f"例：{sorted(b_rels - a_rels)[0][:3]}")
    if a_rels - b_rels and not image_page:
        f.hard("E-PRESERVE", where, "出现原稿没有的关系（定点编辑不新增引用）")
    a_by_id = {s["id"]: s for s in ax["shapes"]}
    for s in bx["shapes"]:
        if s["id"] in deleted or s["id"] not in a_by_id:
            continue
        a = a_by_id[s["id"]]
        if a["meta"] != s["meta"]:
            f.hard("E-PRESERVE", where, f"对象 {s['id']} 的标识与样式容器（nvXxPr / p:style）被改动")
        if s["id"] not in grants and a["xml"] != s["xml"]:
            f.hard("E-PRESERVE", where, f"未点名的对象 {s['id']} 的 XML 被改动（格式、超链接或内容）")


# ---------------------------------------------------------------- 入口


def inspect_cli(argv: List[str]) -> int:
    M.utf8_stdio()
    parser = argparse.ArgumentParser(prog="revise.py inspect", description="只读解析既有 PPTX（不改原文件）")
    parser.add_argument("pptx")
    parser.add_argument("--json", default=None, help="同时写出完整 JSON（只能是 .json，不得是原稿；已存在需 --force）")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的 --json 文件")
    args = parser.parse_args(argv)
    src = Path(args.pptx).resolve()
    rep = M.Report()
    if not src.is_file():
        rep.hard("E-SRC", str(src), "找不到原稿")
    M.guard_paths(rep, {"原稿": src}, [("--json", Path(args.json) if args.json else None, ".json")], args.force)
    if not rep.blocked:
        for problem in M.package_problems(src):
            rep.hard("E-FORMAT", str(src), f"原稿{problem}；拒绝解析")
    if rep.blocked:
        return _blocked(rep)
    data = inspect_deck(src)
    print(to_markdown(data))
    if args.json:
        Path(args.json).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv[1:]) if argv is None else list(argv)
    if argv and argv[0] == "inspect":
        return inspect_cli(argv[1:])

    M.utf8_stdio()
    parser = argparse.ArgumentParser(description="在既有 PPTX 上做定点修改")
    parser.add_argument("source")
    parser.add_argument("ops")
    parser.add_argument("out")
    parser.add_argument("--assets", default=None, help="素材目录；缺省为 ops 同级 assets/ 或所在目录")
    parser.add_argument("--report", default=None)
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    args = parser.parse_args(argv)

    src = Path(args.source).resolve()
    ops_path = Path(args.ops).resolve()
    out = Path(args.out).resolve()
    report_path = Path(args.report).resolve() if args.report else None
    asset_root = M.default_asset_root(ops_path, args.assets)

    rep = M.Report()
    if not src.is_file():
        rep.hard("E-SRC", str(src), "找不到原稿")
    M.guard_paths(rep, {"原稿": src, "edit.json": ops_path},
                  [("输出文件", Path(args.out), ".pptx"), ("--report", Path(args.report) if args.report else None, ".json")], args.force)
    if rep.blocked:
        return _blocked(rep)
    for problem in M.package_problems(src):
        rep.hard("E-FORMAT", str(src), f"原稿{problem}；定点编辑只接受无宏的 .pptx")
    if rep.blocked:
        return _blocked(rep)

    from pptx import Presentation
    try:
        prs = Presentation(str(src))
        spec = json.loads(ops_path.read_text(encoding="utf-8-sig"))
        ops = spec.get("ops") or []
        if not ops:
            raise M.BadOp("edit.json 里没有 ops")
    except (OSError, json.JSONDecodeError, M.BadOp) as exc:
        print(f"[HARD] E-OPS 无法读取：{type(exc).__name__}: {exc}")
        return 2

    before = inspect_deck(src)
    before_parts = structure(src)
    before_slides = slide_part_names(src)
    log: List[Dict[str, Any]] = []
    try:
        grants, bg_touched, declared_deleted, movable = run_ops(prs, ops, asset_root, log, rep)
    except M.BadOp as exc:
        print(f"[HARD] E-OPS 未写入：{exc}")
        return 2
    except Exception as exc:  # 运行期异常：不留半成品
        print(f"[HARD] E-RUNTIME {type(exc).__name__}: {exc}")
        return 3

    # 先落到临时件：保留核对不通过时，交付路径一个字节都不动。
    tmp = out.with_name(f".{out.name}.{os.getpid()}.part")
    try:
        prs.save(str(tmp))
        after = inspect_deck(tmp)
        after_parts = structure(tmp)
        after_slides = slide_part_names(tmp)
    except Exception as exc:
        if tmp.exists():
            tmp.unlink()
        print(f"[HARD] E-RUNTIME {type(exc).__name__}: {exc}")
        return 3

    for label, rec in (("原稿", before), ("改动后", after)):
        for err in rec.get("errors") or []:
            rep.hard("E-PRESERVE", label, f"无法完整解析，逐对象核对不可靠：{err}")
    for problem in M.package_problems(tmp, check_suffix=False):  # 临时件带 .part 后缀，扩展名在落盘前已校验
        rep.hard("E-FORMAT", "改动后的文件", problem)
    touched_pages = {int(op["page"]) for op in ops}
    image_pages = {int(op["page"]) for op in ops if op.get("op") == "image"}
    facts = compare(before, after, grants, bg_touched, declared_deleted, rep, image_pages=image_pages)
    facts["parts"] = compare_parts(before_parts, after_parts, before_slides, after_slides,
                                   touched_pages, rep, movable)
    if not rep.blocked:
        post_check(after, rep)
        os.replace(tmp, out)
    else:
        rejected = out.with_name(out.stem + ".preserve-failed.pptx")
        # 诊断稿也守同一条规则：不压输入，也不静默覆盖已有文件（--force 才替换）
        if rejected in (src, ops_path) or (rejected.exists() and not args.force):
            rejected = out.with_name(f"{out.stem}.{os.getpid()}.preserve-failed.pptx")
        os.replace(tmp, rejected)
        facts["rejected"] = str(rejected)
    status = "BLOCK" if rep.blocked else "READY"

    print(f"\n已改写 {out}（{len(ops)} 个操作，{len(after['pages'])} 页）" if not rep.blocked
          else f"\n保留核对未通过：{out} 未写入；诊断稿 {facts.get('rejected')}（不是交付物）")
    for item in log:
        print(f"  {item['op']:>2}. {item['kind']:<10} 第 {item['page']} 页"
              f"{' · ' + item['target'] if item.get('target') else ''}")
    for issue in rep.items:
        print(issue)
    print(f"保留核对：{status}   问题 {rep.count('HARD')}")
    print(f"页数 {len(after['pages'])}｜母版 {after['masters']}｜版式 {after['layouts']}"
          f"｜包内 {facts['parts']['parts']} 部件（越界改动 {facts['parts']['changed']}／丢失 {facts['parts']['lost']}"
          f"／声明过的改写 {len(facts['parts']['moved'])}）")
    if report_path:
        report_path.write_text(json.dumps({
            "status": status,
            "source": str(src), "output": str(out),
            "sha256": hashlib.sha256(out.read_bytes()).hexdigest() if out.is_file() and not rep.blocked else None,
            "ops": log, "preservation": facts,
            "issues": rep.to_dict(),
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    if not rep.blocked and not src.samefile(out):
        print("原稿未被改动；母版、动画、切换、SmartArt、备注按原样保留。")
    print("机器能证明的到此为止：改动后的版面是否更好看，仍需自己看渲染。")
    return 0 if not rep.blocked else 1


def _blocked(rep: M.Report) -> int:
    for issue in rep.items:
        print(issue)
    print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
    return 2


if __name__ == "__main__":
    sys.exit(main())
