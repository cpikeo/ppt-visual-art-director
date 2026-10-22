#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fidelity.py · spec 与 PPTX 回读的对账（谁丢了、谁挪了、谁变了）。

这一层回答的是「包里到底有没有按判断落下来」，不是「好不好看」：
  * dropped    —— spec 里的元素在包里找不到（编译器漏写 / 字段名写错被静默丢弃）
  * drift      —— 几何、字号、颜色与 spec 不一致
  * payload    —— 图表类型、类别、数值被改写
  * extra      —— 包里多出 spec 没有的形状（伪元素、残留占位符）

对账失败是编译器缺陷，不是稿子缺陷：先修编译，不要改稿迁就（见 FIX_HINTS.LAYOUT_DRIFT）。
"""
from __future__ import annotations

POSITION_TOL = "thresholds.position_drift_px"
SIZE_TOL = "thresholds.size_drift_px"


def compare(deck: dict, readback: dict, policy) -> dict:
    pos_tol = float(policy.num(POSITION_TOL))
    size_tol = float(policy.num(SIZE_TOL))
    size_pt_tol = float(policy.num("thresholds.size_pt_tolerance"))
    theme_colors = (deck.get("theme") or {}).get("colors") or {}
    mismatches: list[dict] = []
    matched = 0

    if len(readback.get("slides") or []) != len(deck.get("slides") or []):
        mismatches.append({"kind": "slide_count", "expected": len(deck["slides"]),
                           "actual": len(readback.get("slides") or []),
                           "detail": "页数不一致：编译丢页或生成多余页"})

    for index, (spec_slide, rb_slide) in enumerate(zip(deck["slides"], readback.get("slides") or [])):
        sid = spec_slide["id"]
        by_name = {shape["name"]: shape for shape in rb_slide["shapes"]}
        expected = {e["id"]: e for e in spec_slide.get("elements") or []}
        for eid, element in expected.items():
            shape = by_name.get(eid)
            if shape is None:
                mismatches.append({"kind": "dropped", "slide": sid, "element": eid,
                                   "detail": f"{element['type']} 元素在 PPTX 中不存在"})
                continue
            matched += 1
            _check_geometry(mismatches, sid, eid, element, shape, pos_tol, size_tol)
            _check_payload(mismatches, sid, eid, element, shape, size_pt_tol, theme_colors)
        for name in by_name:
            if name not in expected:
                mismatches.append({"kind": "extra", "slide": sid, "element": name,
                                   "detail": "PPTX 中出现 spec 没有的形状"})

    changed = [m for m in mismatches if m["kind"] not in ("note",)]
    return {
        "schema": "pvad-fidelity-v1",
        "status": "MISMATCH" if changed else "MATCH",
        "matched_elements": matched,
        "mismatches": mismatches,
        "summary": _summary(changed),
        "note": "对账的是「判断是否被完整执行」，不是「设计是否成立」。",
    }


def _check_geometry(mismatches, sid, eid, element, shape, pos_tol, size_tol) -> None:
    for axis in ("x", "y"):
        if abs(float(element[axis]) - float(shape[axis])) > pos_tol:
            mismatches.append({"kind": "drift", "slide": sid, "element": eid, "field": axis,
                               "expected": element[axis], "actual": shape[axis],
                               "detail": f"{axis} 偏移 {abs(float(element[axis]) - float(shape[axis])):.2f}px"})
    for axis in ("width", "height"):
        if abs(float(element[axis]) - float(shape[axis])) > size_tol:
            mismatches.append({"kind": "drift", "slide": sid, "element": eid, "field": axis,
                               "expected": element[axis], "actual": shape[axis],
                               "detail": f"{axis} 偏差 {abs(float(element[axis]) - float(shape[axis])):.2f}px"})


def _check_payload(mismatches, sid, eid, element, shape, size_pt_tol, theme_colors) -> None:
    etype = element["type"]
    if etype == "text":
        expected_text = str(element.get("text") or "")
        actual_text = str(shape.get("text") or "")
        if expected_text != actual_text:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "text",
                               "expected": expected_text[:40], "actual": actual_text[:40],
                               "detail": "文字被改写或丢失（含换行与不断行空格）"})
        sizes = shape.get("sizes") or []
        if sizes and abs(float(sizes[0]) - float(element["size"])) > size_pt_tol:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "size",
                               "expected": element["size"], "actual": sizes[0],
                               "detail": "字号与 spec 不一致"})
        colors = shape.get("colors") or []
        expected_color = _resolve_color(element.get("color"), theme_colors)
        if colors and expected_color and colors[0].upper() != expected_color:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "color",
                               "expected": expected_color, "actual": colors[0],
                               "detail": "文字颜色被改写（token 解析失败会静默回落）"})
    elif etype == "chart":
        expected_unit = str(element.get("unit_suffix") or "")
        if expected_unit and str(shape.get("unit_suffix") or "") != expected_unit:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "unit_suffix",
                               "expected": expected_unit, "actual": shape.get("unit_suffix"),
                               "detail": "图表单位在编译/回读中丢失：51 会被读成 51 而不是 51%"})
        if shape.get("kind") and element.get("chart_kind") and shape["kind"] != element["chart_kind"]:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "chart_kind",
                               "expected": element["chart_kind"], "actual": shape["kind"],
                               "detail": "图表类型被替换（形状被渲染器改写）"})
        if [str(c) for c in element.get("categories") or []] != [str(c) for c in shape.get("categories") or []]:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "categories",
                               "expected": element.get("categories"), "actual": shape.get("categories"),
                               "detail": "类别标签与 spec 不一致"})
        expected_series = element.get("series") or []
        actual_series = shape.get("series") or []
        if len(expected_series) != len(actual_series):
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "series",
                               "expected": len(expected_series), "actual": len(actual_series),
                               "detail": "序列数量不一致"})
        else:
            for es, as_ in zip(expected_series, actual_series):
                if str(es.get("name")) != str(as_.get("name")):
                    mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "series.name",
                                       "expected": es.get("name"), "actual": as_.get("name"),
                                       "detail": "序列名不一致"})
                if _values_differ(es.get("values") or [], as_.get("values") or []):
                    mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "series.values",
                                       "expected": es.get("values"), "actual": as_.get("values"),
                                       "detail": "数值被改写或截断"})
    elif etype == "image":
        if shape.get("type") != "image":
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "type",
                               "expected": "image", "actual": shape.get("type"),
                               "detail": "图片元素在 PPTX 里不是图片（可能是占位框或被丢弃）"})
            return
        expected_sha = _hex_of(element.get("asset_sha256")) or element.get("asset_sha256")
        actual_sha = shape.get("image_sha256")
        if expected_sha and actual_sha and str(expected_sha).lower() != str(actual_sha).lower():
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "image_sha256",
                               "expected": str(expected_sha)[:12], "actual": str(actual_sha)[:12],
                               "detail": "包里的图片与 spec 记录的素材不是同一个文件（图被替换或压缩改写）"})
        expected_w, expected_h = (element.get("source_px") or {}).get("width"), (element.get("source_px") or {}).get("height")
        actual_px = shape.get("image_px") or [None, None]
        if expected_w and actual_px[0] and int(expected_w) != int(actual_px[0]):
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "source_px",
                               "expected": [expected_w, expected_h], "actual": actual_px,
                               "detail": "包内图像像素尺寸与源文件不一致（被重采样或缩放写入）"})
        if (element.get("fit") or "cover") == "cover" and not any(shape.get("crop") or []):
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "crop",
                               "expected": "cover", "actual": shape.get("crop"),
                               "detail": "声明 cover 但没有裁切：图片被拉伸变形"})
    elif etype == "shape":
        expected_shape = element.get("shape")
        actual_shape = shape.get("shape_kind")
        if expected_shape in ("line", "arrow") and actual_shape not in (None, "line"):
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "shape",
                               "expected": expected_shape, "actual": actual_shape,
                               "detail": "线元素被写成了形状"})
        expected_stroke = _hex_of(element.get("stroke"))
        if expected_stroke and shape.get("stroke") and shape["stroke"].upper() != expected_stroke:
            mismatches.append({"kind": "payload", "slide": sid, "element": eid, "field": "stroke",
                               "expected": expected_stroke, "actual": shape["stroke"],
                               "detail": "描边色不一致"})


def _values_differ(left, right) -> bool:
    if len(left) != len(right):
        return True
    for a, b in zip(left, right):
        if a is None or b is None:
            if a is not b:
                return True
            continue
        if abs(float(a) - float(b)) > 1e-6:
            return True
    return False


def _hex_of(value) -> str | None:
    if isinstance(value, str) and value.startswith("#"):
        return value.upper()
    return None


def _resolve_color(value, theme_colors: dict) -> str | None:
    """把颜色值（十六进制或 theme token）解析为可比较的大写十六进制。

    fidelity 之前对 token 颜色整体跳过（_hex_of("ink") 返回 None），
    导致 token 映射错误会假 PASS。这里把 token 通过 deck.theme.colors 解析后再比对。
    """
    if isinstance(value, str) and value.startswith("#"):
        return value.upper()
    if isinstance(value, str) and value in theme_colors:
        return str(theme_colors[value]).lstrip("#").upper()
    return None


def _summary(changed: list) -> dict:
    out: dict[str, int] = {}
    for item in changed:
        out[item["kind"]] = out.get(item["kind"], 0) + 1
    return out
