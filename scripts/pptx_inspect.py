#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read-only PPTX inspection: observable objects and conservative engineering risks only."""
from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

from compile import readback


class InspectError(ValueError):
    pass


def _preflight(source: Path, policy) -> None:
    budget = policy.req("security.pptx")
    size = source.stat().st_size
    if size > int(budget["max_file_bytes"]):
        raise InspectError(f"PPTX_RESOURCE_BUDGET：文件 {size} bytes > {budget['max_file_bytes']}")
    try:
        with zipfile.ZipFile(source) as package:
            entries = package.infolist()
            if len(entries) > int(budget["max_zip_entries"]):
                raise InspectError(f"PPTX_RESOURCE_BUDGET：ZIP entries {len(entries)} > {budget['max_zip_entries']}")
            total = sum(entry.file_size for entry in entries)
            if total > int(budget["max_uncompressed_bytes"]):
                raise InspectError(f"PPTX_RESOURCE_BUDGET：解压总量 {total} > {budget['max_uncompressed_bytes']}")
            for entry in entries:
                if entry.file_size > int(budget["max_entry_bytes"]):
                    raise InspectError(f"PPTX_RESOURCE_BUDGET：entry {entry.filename} 大小 {entry.file_size} > {budget['max_entry_bytes']}")
                ratio = entry.file_size / max(entry.compress_size, 1)
                if ratio > float(budget["max_compression_ratio"]):
                    raise InspectError(f"PPTX_RESOURCE_BUDGET：entry {entry.filename} 压缩比 {ratio:.1f} > {budget['max_compression_ratio']}")
    except zipfile.BadZipFile:
        raise InspectError(f"PPTX 不是有效 OOXML ZIP：{source}") from None


def inspect_pptx(path, policy) -> dict:
    source = Path(path).expanduser()
    if source.is_symlink():
        raise InspectError(f"PPTX 不得是符号链接：{source}")
    source = source.resolve()
    if not source.is_file() or source.suffix.lower() != ".pptx":
        raise InspectError(f"需要真实 .pptx 文件：{source}")
    _preflight(source, policy)
    try:
        facts = readback(source)
    except Exception as exc:  # python-pptx emits several format-specific exceptions
        raise InspectError(f"PPTX 无法解析：{source} — {exc}") from None
    width, height = facts["slide_size"]
    budget = policy.req("security.pptx")
    if len(facts["slides"]) > int(budget["max_slides"]):
        raise InspectError(f"PPTX_RESOURCE_BUDGET：页面数 {len(facts['slides'])} > {budget['max_slides']}")
    slides, risks = [], []
    counts = {"text": 0, "charts": 0, "images": 0, "shapes": 0}
    explicit_fonts, missing_alt, hyperlinks, slides_with_notes = set(), [], [], 0
    for index, slide in enumerate(facts["slides"], start=1):
        if len(slide["shapes"]) > int(budget["max_objects_per_slide"]):
            raise InspectError(f"PPTX_RESOURCE_BUDGET：第 {index} 页对象数 {len(slide['shapes'])} > {budget['max_objects_per_slide']}")
        objects = []
        if slide.get("notes_text"):
            slides_with_notes += 1
        for shape in slide["shapes"]:
            item = {key: value for key, value in shape.items() if key != "image_blob"}
            kind = item.get("type", "shape")
            count_key = {"text": "text", "chart": "charts", "image": "images", "shape": "shapes"}.get(kind, "shapes")
            counts[count_key] += 1
            objects.append(item)
            name = str(item.get("name") or "")
            x, y = float(item.get("x", 0)), float(item.get("y", 0))
            w, h = float(item.get("width", 0)), float(item.get("height", 0))
            is_line = kind == "shape" and item.get("shape_kind") == "line"
            if (is_line and w <= 0 and h <= 0) or (not is_line and (w <= 0 or h <= 0)):
                risks.append(_risk("GEOMETRY_DEGENERATE", index, name, "对象宽度或高度非正"))
            if x < 0 or y < 0 or x + w > width or y + h > height:
                risks.append(_risk("OUT_OF_CANVAS", index, name, "对象几何超出幻灯片画布"))
            if kind == "text" and item.get("text"):
                explicit_fonts.update(item.get("fonts") or [])
                hyperlinks.extend({"slide": index, "object": name, **link} for link in item.get("hyperlinks", []))
                if not item.get("sizes"):

                    risks.append(_risk("TEXT_SIZE_INHERITED", index, name, "字号来自主题/母版继承，回读未得到显式字号"))
                if not item.get("fonts"):
                    risks.append(_risk("FONT_INHERITED", index, name, "字体来自主题/母版继承，回读未得到显式字体"))
            if kind == "chart" and item.get("kind") == "unknown":
                risks.append(_risk("CHART_KIND_UNKNOWN", index, name, "图表类型不在当前只读识别集合"))
            if kind == "image" and not item.get("alt_text"):
                missing_alt.append({"slide": index, "object": name})
                risks.append(_risk("ALT_TEXT_MISSING", index, name, "图片缺少替代文本；发布前需人工补充或标记为装饰性"))
            if kind == "image" and not item.get("image_sha256"):
                risks.append(_risk("IMAGE_HASH_UNAVAILABLE", index, name, "无法提取内嵌图片散列"))
            if kind == "image" and item.get("image_px"):
                pixels = int(item["image_px"][0]) * int(item["image_px"][1])
                if pixels > int(budget["max_embedded_image_pixels"]):
                    raise InspectError(f"PPTX_RESOURCE_BUDGET：图片 {name} 像素 {pixels} > {budget['max_embedded_image_pixels']}")
        slides.append({"index": index, "id": slide.get("id", ""),
                       "background": slide.get("background"),
                       "layout": slide.get("layout", ""),
                       "notes_present": bool(slide.get("notes_text")),
                       "objects": objects})
    return {
        "schema": "pvad-inspection-v2",
        "source": str(source),
        "artifact_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "slide_size": facts["slide_size"],
        "stats": {"slides": len(slides), "objects": sum(counts.values()), **counts},
        "slides": slides,
        "risks": risks,
        "accessibility": {
            "images": counts["images"],
            "images_missing_alt": missing_alt,
            "hyperlinks": hyperlinks,
            "slides_with_notes": slides_with_notes,
            "reading_order": "z_order recorded; human review required",
        },
        "font_diagnostics": {
            "explicit_fonts": sorted(explicit_fonts),
            "inherited_font_objects": sum(1 for slide in slides for obj in slide["objects"]
                                          if obj.get("type") == "text" and obj.get("text") and not obj.get("fonts")),
            "target_environment_match": "未测量",
        },
        "limitations": [
            "只记录 PPTX 可观察对象，不恢复作者主张、证据状态、叙事任务或设计意图。",
            "空间邻近、连线和页序不自动解释为因果、层级或时间关系。",
            "风险是工程提示，不等于视觉验收；字体回退与真实断行仍需目标 PowerPoint。",
        ],
    }


def _risk(code: str, slide: int, obj: str, detail: str) -> dict:
    return {"code": code, "slide": slide, "object": obj, "detail": detail}
