#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
只读解析既有 PPTX：改稿与评审的第一步。

用法：
    python scripts/inspect.py deck.pptx [--json inspect.json]

报告画布、母版/版式、每页对象的类型·几何·文字·字体·字号·图片 ppi·图表·表格、
备注、动画与切换。只给事实，不给审美结论；不修改原文件。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

EMU = 914400.0


def inch(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 2)


def shape_kind(shape) -> str:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart"
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if shape.has_text_frame and shape.text_frame.text.strip():
        return "text"
    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
        return "placeholder"
    return "shape"


def first_font(shape) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    try:
        runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
    except (AttributeError, ValueError):
        return out
    if not runs:
        return out
    run = runs[0]
    if run.font.size:
        out["size_pt"] = round(run.font.size.pt, 1)
    if run.font.name:
        out["font"] = run.font.name
    try:
        color = run.font.color
        from pptx.enum.dml import MSO_COLOR_TYPE
        if color is not None and color.type == MSO_COLOR_TYPE.RGB:
            out["color"] = f"#{color.rgb}"
        elif color is not None and color.type == MSO_COLOR_TYPE.SCHEME:
            out["color"] = "(theme)"
    except (AttributeError, ValueError):
        pass
    return out


def describe(shape, asset_root: Optional[Path]) -> Dict[str, Any]:
    kind = shape_kind(shape)
    item: Dict[str, Any] = {
        "name": shape.name,
        "kind": kind,
        "box_in": [inch(shape.left), inch(shape.top), inch(shape.width), inch(shape.height)],
    }
    if kind in ("text", "placeholder"):
        text = shape.text_frame.text.strip()
        item["text"] = text.replace("\n", "⏎")[:120]
        item.update(first_font(shape))
    elif kind == "picture":
        try:
            image = shape.image
            px_w, px_h = image.size
            item["px"] = [px_w, px_h]
            width_in = (shape.width or 1) / EMU
            item["ppi"] = round(px_w / width_in) if width_in else None
            item["ext"] = image.ext
        except Exception:
            item["px"] = None
    elif kind == "chart":
        chart = shape.chart
        item["chart_kind"] = str(chart.chart_type).split(" ")[0]
        try:
            item["series"] = len(chart.plots[0].series)
            item["categories"] = len(list(chart.plots[0].categories or []))
        except (IndexError, AttributeError):
            pass
    elif kind == "table":
        table = shape.table
        item["rows"] = len(table.rows)
        item["cols"] = len(table.columns)
    return item


def slide_flags(slide) -> Dict[str, bool]:
    xml = slide._element.xml
    return {"transition": "<p:transition" in xml, "animation": "<p:timing" in xml}


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
    sizes: List[float] = []
    for index, slide in enumerate(prs.slides):
        shapes = [describe(s, None) for s in slide.shapes]
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        page = {
            "index": index + 1,
            "layout": slide.slide_layout.name,
            "flags": slide_flags(slide),
            "notes": notes[:200],
            "objects": shapes,
        }
        out["pages"].append(page)
        sizes.extend(s["size_pt"] for s in shapes if s.get("size_pt"))

    if sizes:
        distinct = sorted(set(sizes))
        out["type_scale"] = {"steps": len(distinct), "sizes": distinct,
                             "min": min(distinct), "max": max(distinct)}
    return out


def to_markdown(data: Dict[str, Any]) -> str:
    lines = [f"# {Path(data['file']).name}", "",
             f"画布 {data['canvas_in'][0]} × {data['canvas_in'][1]} in｜{data['slides']} 页｜"
             f"母版 {data['masters']}｜版式 {data['layouts']}"]
    scale = data.get("type_scale")
    if scale:
        lines.append(f"字号档数 {scale['steps']}（{scale['min']:g}–{scale['max']:g}pt）")
    for page in data["pages"]:
        flags = [name for name, on in page["flags"].items() if on]
        head = f"\n## 第 {page['index']} 页 · {page['layout']}"
        if flags:
            head += f" · 含{'/'.join(flags)}"
        lines.append(head)
        if page["notes"]:
            lines.append(f"备注：{page['notes'][:80]}")
        for obj in page["objects"]:
            box = obj["box_in"]
            detail = ""
            if obj["kind"] in ("text", "placeholder"):
                font = obj.get("font") or "-"
                detail = f"{obj.get('size_pt', '?')}pt {font} {obj.get('color', '')} 「{obj.get('text', '')}」"
            elif obj["kind"] == "picture":
                detail = f"{obj.get('px')} {obj.get('ppi')}ppi"
            elif obj["kind"] == "chart":
                detail = f"{obj.get('chart_kind')} 系列{obj.get('series')} 类别{obj.get('categories')}"
            elif obj["kind"] == "table":
                detail = f"{obj.get('rows')}×{obj.get('cols')}"
            lines.append(f"- {obj['kind']:9} {obj['name'][:28]:28} "
                         f"[{box[0]},{box[1]},{box[2]}×{box[3]}] {detail}")
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
    data = inspect(path)
    print(to_markdown(data))
    if args.json:
        Path(args.json).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
