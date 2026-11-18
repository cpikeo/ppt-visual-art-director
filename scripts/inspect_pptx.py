#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 既有 PPTX 解析器（改稿 / 评审的第一步）

SKILL 要求“改稿先解析既有尺寸、母版、对象和素材，保留未改内容”。本工具把一份已有
PPTX 读成结构化事实：画布、母版/版式、每页对象（位置·文字·字体·字号·颜色·图片 ppi·
图表·表格）、备注，并附若干**客观线索**。

边界（与 SKILL 一致）：
- 只报告事实与机器可判定的线索，不给审美结论；线索 ≠ 缺陷，须对照真实渲染人工判断。
- 不修改、不保存原文件。
- 颜色只报告显式写入的 RGB；继承自主题/母版的颜色标为 (theme)，不猜测。

用法：
    python scripts/inspect_pptx.py deck.pptx                 # Markdown 摘要到 stdout
    python scripts/inspect_pptx.py deck.pptx --json out.json # 同时写完整 JSON
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor  # noqa: F401  (确认依赖完整)
    from pptx.enum.dml import MSO_COLOR_TYPE
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError as exc:  # pragma: no cover
    print(f"[FATAL] 缺少依赖 python-pptx：{exc}\n        请先安装：pip install -r requirements.txt", file=sys.stderr)
    sys.exit(2)

EMU = 914400.0
POLICY_PATH = Path(__file__).resolve().with_name("policy.json")


def _policy() -> Dict[str, Any]:
    try:
        return json.loads(POLICY_PATH.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return {}


POLICY = _policy()
MIN_INFO_PT = POLICY.get("typography", {}).get("min_pt_sets", {}).get("information", {}).get("caption", 12)
OVERLAP_SQIN = POLICY.get("geometry", {}).get("overlap_area_sq_in", 0.5)
BOUND_TOL = POLICY.get("canvas", {}).get("bound_tolerance_in", 0.01)
MIN_PPI = POLICY.get("image", {}).get("min_effective_ppi", 72)


def inch(value: Optional[int]) -> Optional[float]:
    return None if value is None else round(value / EMU, 3)


def _rgb(font) -> Optional[str]:
    try:
        color = font.color
        if color is not None and color.type == MSO_COLOR_TYPE.RGB:
            return f"#{color.rgb}"
        if color is not None and color.type == MSO_COLOR_TYPE.SCHEME:
            return "(theme)"
    except (AttributeError, ValueError):
        pass
    return None


def _kind(shape) -> str:
    if shape.is_placeholder:
        return "placeholder"
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart"
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if st == MSO_SHAPE_TYPE.LINE:
        return "line"
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame and shape.text_frame.text.strip():
        return "text"
    return "shape"


def _text_facts(shape) -> Dict[str, Any]:
    fonts, sizes, colors, bolds = Counter(), Counter(), Counter(), 0
    paragraphs = []
    for para in shape.text_frame.paragraphs:
        text = "".join(run.text for run in para.runs)
        if text.strip():
            paragraphs.append(text)
        for run in para.runs:
            if not run.text.strip():
                continue
            fonts[run.font.name or "(theme)"] += len(run.text)
            sizes[run.font.size.pt if run.font.size else "(inherit)"] += len(run.text)
            colors[_rgb(run.font) or "(inherit)"] += len(run.text)
            bolds += len(run.text) if run.font.bold else 0
    return {
        "text": "\n".join(paragraphs),
        "fonts": dict(fonts), "sizes_pt": {str(k): v for k, v in sizes.items()},
        "colors": dict(colors), "bold_chars": bolds,
    }


def _picture_facts(shape, sw: float) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    try:
        px_w, px_h = shape.image.size
        crop_w = max(1e-6, 1 - (shape.crop_left or 0) - (shape.crop_right or 0))
        crop_h = max(1e-6, 1 - (shape.crop_top or 0) - (shape.crop_bottom or 0))
        shown_w_in = (shape.width or 0) / EMU
        shown_h_in = (shape.height or 0) / EMU
        ppi = min(px_w * crop_w / shown_w_in, px_h * crop_h / shown_h_in) if shown_w_in and shown_h_in else None
        out.update({
            "pixels": [px_w, px_h], "format": shape.image.ext,
            "crop": {k: round(getattr(shape, f"crop_{k}") or 0, 3) for k in ("left", "right", "top", "bottom")},
            "effective_ppi": None if ppi is None else round(ppi, 1),
        })
    except (AttributeError, ValueError, KeyError) as exc:
        out["error"] = f"{type(exc).__name__}: {exc}"
    descr = None
    try:
        descr = shape._element.nvPicPr.cNvPr.get("descr")
    except AttributeError:
        pass
    out["alt_text"] = descr or ""
    return out


def _chart_facts(shape) -> Dict[str, Any]:
    chart = shape.chart
    out: Dict[str, Any] = {"chart_type": str(chart.chart_type).split(".")[-1].split(" ")[0], "series": []}
    try:
        for plot in chart.plots:
            out["categories"] = [str(c) for c in plot.categories]
            for s in plot.series:
                out["series"].append({"name": s.name, "values": list(s.values)})
    except (AttributeError, ValueError):
        out["note"] = "无法读取图表数据"
    out["has_title"] = bool(chart.has_title)
    return out


def _table_facts(shape) -> Dict[str, Any]:
    tbl = shape.table
    cells = [[c.text for c in row.cells] for row in tbl.rows]
    return {"rows": len(tbl.rows), "cols": len(tbl.columns), "header": cells[0] if cells else [],
            "sample": cells[1:3]}


def _walk(shapes, sw: float, sh: float, depth: int = 0) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    for z, shape in enumerate(shapes):
        kind = _kind(shape)
        item: Dict[str, Any] = {
            "z": z, "name": shape.name, "kind": kind, "depth": depth,
            "x": inch(shape.left), "y": inch(shape.top), "w": inch(shape.width), "h": inch(shape.height),
        }
        if kind == "placeholder":
            item["placeholder"] = str(shape.placeholder_format.type).split(".")[-1].split(" ")[0]
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame and shape.text_frame.text.strip():
            item.update(_text_facts(shape))
        if kind == "picture":
            item.update(_picture_facts(shape, sw))
        elif kind == "chart":
            item.update(_chart_facts(shape))
        elif kind == "table":
            item.update(_table_facts(shape))
        elif kind == "group":
            item["children"] = _walk(shape.shapes, sw, sh, depth + 1)
            item["note"] = "组内坐标未做变换换算；组被缩放时子对象位置仅供参考"
        items.append(item)
    return items


def _flatten(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    flat: List[Dict[str, Any]] = []
    for it in items:
        flat.append(it)
        flat.extend(_flatten(it.get("children", [])))
    return flat


def _overlap(a: Dict[str, Any], b: Dict[str, Any]) -> float:
    if None in (a["x"], a["y"], a["w"], a["h"], b["x"], b["y"], b["w"], b["h"]):
        return 0.0
    dx = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
    dy = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    return dx * dy if dx > 0 and dy > 0 else 0.0


def _clues(items: List[Dict[str, Any]], sw: float, sh: float, slide_no: int) -> List[str]:
    clues: List[str] = []
    flat = [i for i in _flatten(items) if i["depth"] == 0]
    for it in flat:
        if None in (it["x"], it["w"], it["y"], it["h"]):
            continue
        if it["x"] < -BOUND_TOL or it["y"] < -BOUND_TOL or it["x"] + it["w"] > sw + BOUND_TOL or it["y"] + it["h"] > sh + BOUND_TOL:
            if it["kind"] != "picture":
                clues.append(f"[{it['name']}] 超出页面边界（可能为有意出血，请对照渲染）")
        for size in it.get("sizes_pt", {}):
            try:
                if float(size) < MIN_INFO_PT:
                    clues.append(f"[{it['name']}] 含 {size}pt 文字，低于信息级下限 {MIN_INFO_PT}pt")
            except ValueError:
                pass
        if it["kind"] == "picture":
            ppi = it.get("effective_ppi")
            if ppi is not None and ppi < MIN_PPI:
                clues.append(f"[{it['name']}] 图片有效 {ppi} ppi，低于 {MIN_PPI}")
            if not it.get("alt_text"):
                clues.append(f"[{it['name']}] 图片缺少替代文字（无障碍与评审可读性）")
        if it["kind"] == "placeholder" and not it.get("text"):
            clues.append(f"[{it['name']}] 空占位符（成稿可能显示提示语）")
    texty = [i for i in flat if i.get("text")]
    for idx, a in enumerate(texty):
        for b in texty[idx + 1:]:
            area = _overlap(a, b)
            if area >= OVERLAP_SQIN:
                clues.append(f"[{a['name']}] 与 [{b['name']}] 文本框重叠约 {area:.2f} sq in")
    sizes = Counter(s for i in flat for s in i.get("sizes_pt", {}) if s != "(inherit)")
    if len(sizes) > 5:
        clues.append(f"本页出现 {len(sizes)} 种显式字号：{sorted(sizes, key=float)}，层级可能过碎")
    if not any(i.get("text") for i in flat) and not any(i["kind"] in ("chart", "table", "picture") for i in flat):
        clues.append("本页没有可识别的文字或图表/图片对象")
    return clues


def inspect(path: Path) -> Dict[str, Any]:
    prs = Presentation(str(path))
    sw, sh = prs.slide_width / EMU, prs.slide_height / EMU
    deck_fonts, deck_sizes, deck_colors = Counter(), Counter(), Counter()
    slides = []
    for no, slide in enumerate(prs.slides, 1):
        items = _walk(slide.shapes, sw, sh)
        for it in _flatten(items):
            for k, v in it.get("fonts", {}).items():
                deck_fonts[k] += v
            for k, v in it.get("sizes_pt", {}).items():
                deck_sizes[k] += v
            for k, v in it.get("colors", {}).items():
                deck_colors[k] += v
        title = None
        try:
            if slide.shapes.title is not None and slide.shapes.title.text_frame.text.strip():
                title = slide.shapes.title.text_frame.text.strip()
        except AttributeError:
            pass
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        clues = _clues(items, sw, sh, no)
        if title is None:
            clues.append("没有标题占位符（大纲视图/读屏无法识别页标题；若为设计所需可忽略）")
        slides.append({
            "index": no, "layout": slide.slide_layout.name, "title_placeholder": title,
            "notes": notes, "objects": items, "clues": clues,
        })
    ratio = round(sw / sh, 3)
    return {
        "file": str(path), "slide_count": len(slides),
        "canvas_in": {"w": round(sw, 3), "h": round(sh, 3), "ratio": ratio,
                      "label": "16:9" if abs(ratio - 16 / 9) < 0.01 else "4:3" if abs(ratio - 4 / 3) < 0.01 else "custom"},
        "masters": [{"name": m.name, "layouts": [l.name for l in m.slide_layouts]} for m in prs.slide_masters],
        "deck_fonts_by_chars": dict(deck_fonts.most_common()),
        "deck_sizes_by_chars": {k: v for k, v in sorted(deck_sizes.items(), key=lambda kv: -kv[1])},
        "deck_colors_by_chars": dict(deck_colors.most_common()),
        "slides": slides,
    }


def to_markdown(report: Dict[str, Any]) -> str:
    c = report["canvas_in"]
    lines = [f"# 既有 PPTX 解析：{Path(report['file']).name}", "",
             f"- 画布：{c['w']} × {c['h']} in（{c['label']}，比例 {c['ratio']}）；共 {report['slide_count']} 页",
             f"- 母版：" + "；".join(f"{m['name'] or '(未命名母版)'}（{len(m['layouts'])} 个版式）" for m in report["masters"]),
             f"- 字体（按字符数）：{report['deck_fonts_by_chars'] or '无显式字体（全部继承主题）'}",
             f"- 字号（按字符数）：{report['deck_sizes_by_chars']}",
             f"- 显式文字颜色：{report['deck_colors_by_chars']}", ""]
    total_clues = sum(len(s["clues"]) for s in report["slides"])
    lines.append(f"> 线索共 {total_clues} 条。线索只是客观提示，不是缺陷结论；请对照真实渲染复核。")
    for s in report["slides"]:
        lines += ["", f"## 第 {s['index']} 页 · 版式「{s['layout']}」" + (f" · 标题：{s['title_placeholder']}" if s["title_placeholder"] else "")]
        for it in _flatten(s["objects"]):
            pad = "  " * it["depth"]
            geo = f"({it['x']}, {it['y']}) {it['w']}×{it['h']} in"
            extra = ""
            if it.get("text"):
                preview = it["text"].replace("\n", " ⏎ ")
                extra = f" 「{preview[:48]}{'…' if len(preview) > 48 else ''}」 字号{list(it['sizes_pt'])} 字体{list(it['fonts'])}"
            elif it["kind"] == "picture":
                extra = f" {it.get('pixels')}px 有效 {it.get('effective_ppi')} ppi"
            elif it["kind"] == "chart":
                extra = f" {it['chart_type']}，{len(it['series'])} 个系列"
            elif it["kind"] == "table":
                extra = f" {it['rows']}×{it['cols']}"
            lines.append(f"{pad}- z{it['z']} {it['kind']} `{it['name']}` {geo}{extra}")
        if s["notes"]:
            lines.append(f"- 备注：{s['notes'][:80]}")
        for clue in s["clues"]:
            lines.append(f"  - ⚠ {clue}")
    return "\n".join(lines) + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="解析既有 PPTX（只读）")
    parser.add_argument("pptx")
    parser.add_argument("--json", metavar="PATH", help="另存完整 JSON")
    args = parser.parse_args(argv)
    path = Path(args.pptx)
    if not path.is_file():
        print(f"[E001] 文件不存在：{path}", file=sys.stderr)
        return 1
    try:
        report = inspect(path)
    except Exception as exc:  # python-pptx 对损坏文件抛多种异常；统一给出可读失败
        print(f"[E001] 无法解析 PPTX（文件损坏、加密或非 OOXML？）：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sys.stdout.write(to_markdown(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
