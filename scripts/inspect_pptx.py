#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
只读解析既有 PPTX：改稿与评审的第一步。

用法：
    python scripts/inspect_pptx.py deck.pptx [--json inspect.json]

报告画布、母版/版式、主题字体、宏、每页对象的类型·几何·文字·逐 run 字体（含东亚字族）·
图片·图表载荷·表格内容·组合结构·备注·动画与切换·隐藏页，以及字号档数与用色统计。
只给事实，不给审美结论；不改原文件。单个对象解析失败不会中断整份报告。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

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
        root = etree.fromstring(data, etree.XMLParser(remove_blank_text=True))
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
        pres = etree.fromstring(zf.read("ppt/presentation.xml"))
        rels = etree.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
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
                root = etree.fromstring(rel.target_part.blob)
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


def inspect(pptx: Path) -> Dict[str, Any]:
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


def utf8_stdio() -> None:
    """控制台代码页不是 UTF-8 时，⏎ 之类的字符会抛 UnicodeEncodeError；这里统一兜住。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def main(argv: Optional[List[str]] = None) -> int:
    utf8_stdio()
    parser = argparse.ArgumentParser(description="只读解析既有 PPTX")
    parser.add_argument("pptx")
    parser.add_argument("--json", default=None, help="另存完整 JSON")
    args = parser.parse_args(argv)
    path = Path(args.pptx)
    if not path.is_file():
        print(f"找不到 {path}")
        return 2
    try:
        data = inspect(path)
    except Exception as exc:
        print(f"无法解析 {path}：{type(exc).__name__}: {exc}")
        return 2
    print(to_markdown(data))
    if args.json:
        Path(args.json).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
