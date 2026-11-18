#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
只读解析既有 PPTX：改稿与评审的第一步。

用法：
    python scripts/inspect_pptx.py deck.pptx [--json inspect.json]

报告画布、母版/版式、每页对象的类型·几何·文字·逐段字体·图片·图表载荷·表格内容·
组合结构·备注·动画与切换，以及字号档数与用色统计。只给事实，不给审美结论；不改原文件。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

EMU = 914400.0


def inch(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 2)


def kind_of(shape) -> str:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart"
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table"
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        return "placeholder"
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        return "text"
    return "shape"


def run_styles(shape) -> Dict[str, Any]:
    """逐段逐 run 的字体事实：只看第一个 run 会漏掉混排与局部改色。"""
    from pptx.enum.dml import MSO_COLOR_TYPE
    out: Dict[str, Any] = {"paragraphs": []}
    try:
        paragraphs = shape.text_frame.paragraphs
    except (AttributeError, ValueError):
        return out
    sizes, colors, fonts, weights = set(), set(), set(), set()
    for para in paragraphs:
        runs = []
        for run in para.runs:
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
            item = {"text": run.text, "size_pt": size, "font": run.font.name,
                    "bold": bool(run.font.bold), "italic": bool(run.font.italic), "color": color}
            runs.append(item)
            if size:
                sizes.add(size)
            if color:
                colors.add(color)
            if run.font.name:
                fonts.add(run.font.name)
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


def describe(shape, depth: int = 0) -> Dict[str, Any]:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    kind = kind_of(shape)
    item: Dict[str, Any] = {
        "name": shape.name,
        "kind": kind,
        "box_in": [inch(shape.left), inch(shape.top), inch(shape.width), inch(shape.height)],
        "rotation": round(shape.rotation or 0, 1) if hasattr(shape, "rotation") else 0,
    }
    if kind in ("text", "placeholder"):
        item["text"] = shape.text_frame.text.replace("\n", "⏎")[:240]
        item["styles"] = run_styles(shape)
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
        item["title"] = chart.chart_title.text_frame.text if chart.has_title and chart.has_title else None
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
    elif kind == "table":
        table = shape.table
        grid = []
        for r in range(min(len(table.rows), 12)):
            grid.append([table.cell(r, c).text for c in range(len(table.columns))])
        item.update({"rows": len(table.rows), "cols": len(table.columns), "grid": grid,
                     "col_widths_in": [round(col.width / EMU, 2) for col in table.columns]})
    elif kind == "group":
        item["children"] = [describe(child, depth + 1) for child in shape.shapes]
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        try:
            fill = shape.fill
            if fill.type is not None and str(fill.type).startswith("MSO_FILL_TYPE.SOLID"):
                item["fill"] = f"#{fill.fore_color.rgb}"
            elif fill.type is not None and "GRADIENT" in str(fill.type):
                item["fill"] = "gradient"
            if shape.line.width:
                item["stroke_pt"] = round(shape.line.width.pt, 2)
        except Exception:
            pass
    return item


def slide_flags(slide) -> Dict[str, bool]:
    xml = slide._element.xml
    return {"transition": "<p:transition" in xml, "timing": "<p:timing" in xml}


def collect_fonts(shapes, fonts: set, sizes: set) -> None:
    for shape in shapes:
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
        if getattr(shape, "shape_type", None) is not None and shape.shape_type == 6:
            collect_fonts(shape.shapes, fonts, sizes)


def inspect(pptx: Path) -> Dict[str, Any]:
    from pptx import Presentation

    prs = Presentation(str(pptx))
    out: Dict[str, Any] = {
        "file": str(pptx),
        "canvas_in": [inch(prs.slide_width), inch(prs.slide_height)],
        "slides": len(prs.slides),
        "masters": len(prs.slide_masters),
        "layouts": sum(len(m.slide_layouts) for m in prs.slide_masters),
        "pages": [],
    }
    fonts: set = set()
    sizes: set = set()
    for index, slide in enumerate(prs.slides):
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        page = {
            "index": index + 1,
            "layout": slide.slide_layout.name,
            "flags": slide_flags(slide),
            "notes": notes[:400],
            "objects": [describe(s) for s in slide.shapes],
        }
        collect_fonts(slide.shapes, fonts, sizes)
        out["pages"].append(page)
    out["fonts"] = sorted(fonts)
    if sizes:
        distinct = sorted(sizes)
        out["type_scale"] = {"steps": len(distinct), "sizes": distinct, "min": min(distinct), "max": max(distinct)}
    return out


def to_markdown(data: Dict[str, Any]) -> str:
    lines = [f"# {Path(data['file']).name}", "",
             f"画布 {data['canvas_in'][0]} × {data['canvas_in'][1]} in｜{data['slides']} 页｜"
             f"母版 {data['masters']}｜版式 {data['layouts']}"]
    scale = data.get("type_scale")
    if scale:
        lines.append(f"字号 {scale['steps']} 档（{scale['min']:g}–{scale['max']:g}pt）")
    if data.get("fonts"):
        lines.append(f"字体 {data['fonts']}")

    def emit(obj: Dict[str, Any], depth: int = 0) -> None:
        pad = "  " * depth
        box = obj["box_in"]
        head = f"{pad}- {obj['kind']:11} {obj['name'][:26]:26} [{box[0]},{box[1]},{box[2]}×{box[3]}]"
        if obj.get("rotation"):
            head += f" {obj['rotation']}°"
        detail = ""
        if obj["kind"] in ("text", "placeholder"):
            styles = obj.get("styles", {})
            sizes = "/".join(f"{s:g}" for s in styles.get("sizes", []))
            detail = f"{sizes}pt {'/'.join(styles.get('fonts', []))} 「{obj.get('text', '')}」"
        elif obj["kind"] == "picture":
            detail = f"{obj.get('px')} {obj.get('ppi')}ppi {obj.get('ext')}"
        elif obj["kind"] == "chart":
            plots = obj.get("plots", [])
            bits = []
            for plot in plots:
                bits.append(" / ".join(f"{s['name']}({s['count']})" for s in plot["series"]))
            detail = f"{obj.get('chart_type')} {len(plots)} 层｜" + "；".join(bits)
            if obj.get("title"):
                detail += f"｜题「{obj['title'][:28]}」"
        elif obj["kind"] == "table":
            detail = f"{obj['rows']}×{obj['cols']}  首行 {obj.get('grid', [[]])[0]}"
        elif obj["kind"] == "group":
            detail = f"{len(obj.get('children', []))} 个子对象"
        lines.append(f"{head} {detail}".rstrip())
        for child in obj.get("children", []):
            emit(child, depth + 1)

    for page in data["pages"]:
        flags = [name for name, on in page["flags"].items() if on]
        head = f"\n## 第 {page['index']} 页 · {page['layout']}"
        if flags:
            head += f" · 含{'/'.join(flags)}"
        lines.append(head)
        if page["notes"]:
            lines.append(f"备注：{page['notes'][:100]}")
        for obj in page["objects"]:
            emit(obj)
    return "\n".join(lines)


def main(argv: Optional[List[str]] = None) -> int:
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
