#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render.py · 渲染对照（同一渲染器吃两种真相：spec 与 PPTX 回读）。

为什么这样分层：
  * spec 渲染 = 「作者与判断想让它长什么样」。
  * PPTX 回读渲染 = 「包里实际存了什么几何、字号、文字、数据」。
  * 两者用同一个渲染器、同一套字体度量，所以像素差只可能来自编译或字段丢失，
    不可能来自两个渲染引擎的口径差异——这正是「预览与最终 PPTX 不一致」的可测形式。
  * 这仍然不是目标 PowerPoint 渲染：字体回落、真实断行、投影可读性不在此证明。
    报告里恒写 NOT VERIFIED — target renderer，不拿结构 PASS 冒充视觉终验。
"""
from __future__ import annotations

from pathlib import Path

from typography import FontBook, wrap

_SPEC_PAGE_KEYS = ("background",)


class RenderError(RuntimeError):
    pass


# ──────────────────────────────────────────────────────────────────────────
# 页面模型
# ──────────────────────────────────────────────────────────────────────────
def page_model_from_spec(deck: dict, slide: dict) -> dict:
    colors = (deck.get("theme") or {}).get("colors") or {}
    items = []
    for element in slide.get("elements") or []:
        item = {"id": element["id"], "kind": element["type"],
                "x": float(element["x"]), "y": float(element["y"]),
                "w": float(element["width"]), "h": float(element["height"])}
        etype = element["type"]
        if etype == "text":
            item.update({"text": str(element.get("text") or ""),
                         "size": float(element.get("size") or 16),
                         "color": _hex(element.get("color") or "ink", colors),
                         "bold": bool(element.get("bold")),
                         "align": str(element.get("align") or "left"),
                         "anchor": str(element.get("anchor") or "top"),
                         "line_height": float(element.get("line_height") or 1.4),
                         "tracking": float(element.get("tracking") or 0.0)})
        elif etype == "shape":
            item.update({"shape": element.get("shape"),
                         "fill": _hex(element.get("fill"), colors) if isinstance(element.get("fill"), str) else None,
                         "stroke": _hex(element.get("stroke"), colors) if isinstance(element.get("stroke"), str) else None,
                         "stroke_width": float(element.get("stroke_width") or 1.0)})
        elif etype == "image":
            item.update({"path": element.get("source") or "",
                         "fit": element.get("fit") or "cover",
                         "overlay": float(element.get("overlay") or 0.0),
                         "base_dir": (deck.get("meta") or {}).get("base_dir") or ""})
        elif etype == "chart":
            item.update({"chart_kind": element.get("chart_kind"),
                         "categories": list(element.get("categories") or []),
                         "series": [{"name": s.get("name"), "values": list(s.get("values") or [])}
                                    for s in (element.get("series") or [])],
                         "show_values": bool(element.get("show_values"))})
        items.append(item)
    return {"id": slide.get("id"), "colors": colors,
            "width": float(deck["canvas"]["width"]),
            "height": float(deck["canvas"]["height"]),
            "background": _hex(slide.get("background") or "background", colors),
            "items": items}


def page_model_from_readback(deck: dict, rb_slide: dict, spec_slide: dict | None = None) -> dict:
    """从 PPTX 读回的事实构建页面模型。颜色缺失时回落到 spec 同名元素（并在 fidelity 里记账）。"""
    colors = (deck.get("theme") or {}).get("colors") or {}
    spec_items = {e["id"]: e for e in (spec_slide or {}).get("elements") or []}
    items = []
    for shape in rb_slide["shapes"]:
        name = shape["name"]
        spec = spec_items.get(name, {})
        base = {"id": name, "x": shape["x"], "y": shape["y"], "w": shape["width"], "h": shape["height"]}
        if shape.get("type") == "text":
            size = shape.get("sizes")[0] if shape.get("sizes") else float(spec.get("size") or 16)
            color = "#" + shape["colors"][0] if shape.get("colors") else _hex(spec.get("color") or "ink", colors)
            base.update({"kind": "text", "text": shape.get("text") or "", "size": size,
                         "color": color, "bold": bool(spec.get("bold")),
                         "align": str(spec.get("align") or "left"),
                         "anchor": str(spec.get("anchor") or "top"),
                         "line_height": float(spec.get("line_height") or 1.4),
                         "tracking": float(spec.get("tracking") or 0.0)})
        elif shape.get("type") == "image":
            base.update({"kind": "image", "blob": shape.get("image_blob") or b"",
                         "blob_sha256": shape.get("image_sha256") or "",
                         "taken_from": "pptx_blob", "path": "",
                         "fit": shape.get("fit") or "cover",
                         "overlay": float(spec.get("overlay") or 0.0),
                         "base_dir": (deck.get("meta") or {}).get("base_dir") or ""})
        elif shape.get("type") == "chart":
            base.update({"kind": "chart", "chart_kind": shape.get("kind") or spec.get("chart_kind"),
                         "categories": shape.get("categories") or [],
                         "series": [{"name": s["name"], "values": s["values"]} for s in shape.get("series") or []],
                         "show_values": bool(spec.get("show_values"))})
        else:
            fill = shape.get("fill")
            base.update({"kind": "shape", "shape": shape.get("shape_kind") or spec.get("shape"),
                         "fill": fill, "stroke": shape.get("stroke") or
                         (_hex(spec.get("stroke"), colors) if isinstance(spec.get("stroke"), str) else None),
                         "stroke_width": float(spec.get("stroke_width") or 1.0)})
        items.append(base)
    return {"id": rb_slide.get("id") or (spec_slide or {}).get("id"), "colors": colors,
            "width": float(deck["canvas"]["width"]), "height": float(deck["canvas"]["height"]),
            "background": _hex((spec_slide or {}).get("background") or "background", colors),
            "items": items}


def _hex(value, colors: dict) -> str:
    if isinstance(value, str) and value.startswith("#"):
        return value.upper()
    if isinstance(value, str) and value in colors:
        return str(colors[value]).upper()
    return "#000000"


# ──────────────────────────────────────────────────────────────────────────
# 渲染
# ──────────────────────────────────────────────────────────────────────────
def render_page(model: dict, out_path: str | Path, policy, fontbook: FontBook) -> str:
    from PIL import Image, ImageDraw
    scale = int(policy.opt("render.ghost_supersample", 2))
    w = int(model["width"] * scale)
    h = int(model["height"] * scale)
    image = Image.new("RGB", (w, h), _rgb_tuple(model["background"]))
    draw = ImageDraw.Draw(image)
    palette = model.get("colors") or {}
    for item in model["items"]:
        kind = item["kind"]
        if kind == "chart":
            item = {**item, "colors": palette}
        if kind == "shape":
            _draw_shape(draw, item, scale, fontbook, policy)
        elif kind == "text":
            _draw_text(draw, item, scale, fontbook, policy)
        elif kind == "chart":
            _draw_chart(draw, item, scale, fontbook, policy)
        elif kind == "image":
            _draw_image(image, item, scale, policy)
    if scale > 1:
        image = image.resize((int(model["width"]), int(model["height"])), Image.LANCZOS)
    target = Path(out_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    image.save(target, "PNG", optimize=True)
    return str(target)


def _rgb_tuple(value: str) -> tuple:
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _font(size: float, bold: bool, scale: int, fontbook: FontBook):
    return fontbook._font("cn", max(1, int(round(size * scale))), bold)


def _draw_text(draw, item, scale, fontbook: FontBook, policy) -> None:
    text = item.get("text") or ""
    if not text:
        return
    size = item["size"]
    metrics = wrap(text, item["w"], size, policy=policy, fontbook=fontbook,
                   bold=item["bold"], tracking_pt=item.get("tracking") or 0.0,
                   max_lines=policy.max_lines(_ladder(policy, size)))
    if not metrics.lines:
        return
    font = _font(size, item["bold"], scale, fontbook)
    lh = item.get("line_height") or metrics.line_height
    line_px = size * lh * scale
    total = ((metrics.count - 1) * size * lh + size * max(lh, float(policy.req("typography.ink_height_floor")))) * scale
    if item["anchor"] == "bottom":
        top = (item["y"] + item["h"]) * scale - total
    elif item["anchor"] == "middle":
        top = (item["y"] + item["h"]) * scale - total - (item["h"] * scale - total) / 2
    else:
        top = item["y"] * scale
    color = _rgb_tuple(item["color"])
    for i, line in enumerate(metrics.lines):
        x = item["x"] * scale
        width = font.getlength(line)
        if item["align"] == "center":
            x += (item["w"] * scale - width) / 2
        elif item["align"] == "right":
            x += item["w"] * scale - width
        baseline_top = top + i * line_px
        # PIL 用 top-left 锚点；用 bbox 校正，避免不同字体的内部行差
        bbox = font.getbbox(line)
        draw.text((x - bbox[0], baseline_top - bbox[1]), line, font=font, fill=color)


def visible_image(box_px, source, fit: str = "cover", base_dir=None, policy=None, blob=None):
    """返回「这个框里最终会被看见的图像内容」（已按 fit 裁切/留白）。

    渲染与素材度量共用它：判断文字压图对比度时，量的必须是观众真正看见的那块，
    而不是原图——否则 cover 裁切后的实测就会失真。
    """
    from PIL import Image
    from assets import load, resolve_source
    width = max(1, int(round(box_px[0])))
    height = max(1, int(round(box_px[1])))
    if blob:
        import io
        try:
            image = Image.open(io.BytesIO(blob)).convert("RGB")
        except (OSError, ValueError):
            return None
    else:
        path = resolve_source(source, base_dir)
        if path is None:
            return None
        image = load(path)
    src_w, src_h = image.size
    if fit == "contain":
        ratio = min(width / src_w, height / src_h)
        resized = image.resize((max(1, int(src_w * ratio)), max(1, int(src_h * ratio))), Image.LANCZOS)
        canvas = Image.new("RGB", (width, height), (255, 255, 255))
        canvas.paste(resized, ((width - resized.width) // 2, (height - resized.height) // 2))
        return canvas
    ratio = max(width / src_w, height / src_h)
    resized = image.resize((max(1, int(src_w * ratio)), max(1, int(src_h * ratio))), Image.LANCZOS)
    left = (resized.width - width) // 2
    top = (resized.height - height) // 2
    return resized.crop((left, top, left + width, top + height))


def _draw_image(canvas, item, scale: int, policy) -> None:
    """按 fit 规则贴图：cover 等比铺满后居中裁切，contain 等比装入留白。"""
    from PIL import Image, ImageDraw
    target_w = max(1, int(round(item["w"] * scale)))
    target_h = max(1, int(round(item["h"] * scale)))
    layer = visible_image((target_w, target_h), item.get("path"), item.get("fit") or "cover",
                          item.get("base_dir") or None, policy, blob=item.get("blob"))
    if layer is None:
        # 缺图不能静默：画红框叉号占位，让像素对照与人工预览都看得见
        box = [item["x"] * scale, item["y"] * scale,
               (item["x"] + item["w"]) * scale, (item["y"] + item["h"]) * scale]
        draw = ImageDraw.Draw(canvas)
        draw.rectangle(box, outline=(180, 60, 40), width=max(2, 2 * scale))
        draw.line([box[0], box[1], box[2], box[3]], fill=(180, 60, 40), width=max(1, scale))
        draw.line([box[0], box[3], box[2], box[1]], fill=(180, 60, 40), width=max(1, scale))
        return
    overlay = float(item.get("overlay") or 0.0)
    if overlay > 0:
        veil = Image.new("RGB", layer.size, (20, 22, 26))
        layer = Image.blend(layer, veil, min(1.0, overlay))
    canvas.paste(layer, (int(round(item["x"] * scale)), int(round(item["y"] * scale))))


def _draw_shape(draw, item, scale, fontbook, policy) -> None:
    x0, y0 = item["x"] * scale, item["y"] * scale
    x1, y1 = x0 + item["w"] * scale, y0 + item["h"] * scale
    if item.get("shape") in ("line", "arrow"):
        stroke = item.get("stroke") or "#999999"
        width = max(1, int(round(item.get("stroke_width", 1.0) * scale)))
        draw.line([(x0, y0), (x1, y1)], fill=_rgb_tuple(stroke), width=width)
        return
    fill = item.get("fill")
    stroke = item.get("stroke")
    box = [min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)]
    if item.get("shape") == "ellipse":
        draw.ellipse(box, fill=_rgb_tuple(fill) if fill else None,
                     outline=_rgb_tuple(stroke) if stroke else None)
    else:
        draw.rectangle(box, fill=_rgb_tuple(fill) if fill else None,
                       outline=_rgb_tuple(stroke) if stroke else None)


def _chart_colors(item: dict, policy) -> dict:
    """图表配色跟主题走，不跟策略默认调色板走。

    为什么：颜色是稿件的品牌决定（brief.brand.colors），策略里的默认色只是没有品牌时的
    兜底。曾经这里直接读 policy.colors()，于是渲染图上的柱子/扇区永远是策略默认色
    （例如 #1F4E5A），既不是作者的品牌色，也让「预览与交付一致」的像素对照失去意义
    ——两条路径都错成一样，差值为 0。缺 token 时才回落，不让缺省静默改写品牌色。
    """
    palette = dict(policy.colors())
    theme = {k: v for k, v in (item.get("colors") or {}).items() if isinstance(v, str)}
    palette.update(theme)
    return palette


def _draw_chart(draw, item, scale, fontbook: FontBook, policy) -> None:
    kind = item.get("chart_kind") or "column"
    categories = item.get("categories") or []
    series = item.get("series") or []
    if not categories or not series:
        return
    colors = _chart_colors(item, policy)
    x, y = item["x"] * scale, item["y"] * scale
    w, h = item["w"] * scale, item["h"] * scale
    font = _font(max(10.0, item["size"] if item.get("size") else 13.0), False, scale, fontbook)
    label_color = _rgb_tuple(colors["muted"])
    ink = _rgb_tuple(colors["ink"])
    primary = _rgb_tuple(colors["primary"])
    accent = _rgb_tuple(colors["accent"])

    if kind == "donut":
        values = series[0]["values"]
        total = sum(v for v in values if v > 0) or 1.0
        box = [x + w * 0.08, y + h * 0.06, x + w * 0.08 + min(w, h) * 0.88, y + h * 0.06 + min(w, h) * 0.88]
        start = -90.0
        palette = [colors["primary"], colors["accent"], colors["muted"], colors["ink"]]
        for i, value in enumerate(values):
            extent = 360.0 * (value / total)
            draw.pieslice(box, start, start + extent, fill=_rgb_tuple(palette[i % len(palette)]))
            start += extent
        inner = [box[0] + (box[2] - box[0]) * 0.29, box[1] + (box[3] - box[1]) * 0.29,
                 box[2] - (box[2] - box[0]) * 0.29, box[3] - (box[3] - box[1]) * 0.29]
        draw.ellipse(inner, fill=_rgb_tuple(colors["background"]))
        return

    if kind == "progress":
        values = [v for s in series for v in s["values"]]
        peak = max(values) or 1.0
        ratio = values[0] / peak
        draw.rectangle([x, y, x + w, y + h], fill=_rgb_tuple(colors["hairline"]))
        draw.rectangle([x, y, x + w * ratio, y + h], fill=primary)
        return

    pad_left, pad_bottom, pad_top = w * 0.12, h * 0.18, h * 0.08
    plot_w, plot_h = w - pad_left - w * 0.04, h - pad_bottom - pad_top
    values = [v for s in series for v in s["values"] if v is not None]
    vmax = max(values + [0.0])
    vmin = min(values + [0.0])
    span = (vmax - vmin) or 1.0

    def ypos(value):
        return y + pad_top + plot_h * (1 - (value - vmin) / span)

    if kind == "bar":
        # 横向条形：PowerPoint 的 BAR_CLUSTERED 把第一个类别画在最下面（自下而上），
        # 预览必须跟着一样，否则「预览与交付一致」只是两条路径一起错。
        n = len(categories)
        band = plot_h / max(n, 1)
        pad_left_b = w * 0.24
        plot_w_b = w - pad_left_b - w * 0.08
        bar_h = band * 0.58 / max(len(series), 1)
        draw.line([(x + pad_left_b, y + pad_top), (x + pad_left_b, y + pad_top + plot_h)],
                  fill=_rgb_tuple(colors["hairline"]), width=max(1, scale))
        for i, category in enumerate(categories):
            cy = y + pad_top + plot_h - band * (i + 0.5)
            draw.text((x + w * 0.02, cy), str(category), font=font, fill=label_color, anchor="lm")
        for si, s in enumerate(series):
            color = primary if si == 0 else accent
            for i, value in enumerate(s["values"]):
                if value is None:
                    continue
                cy = y + pad_top + plot_h - band * (i + 0.5) - bar_h * len(series) / 2 + bar_h * si
                length = plot_w_b * (value - vmin) / span
                draw.rectangle([x + pad_left_b, cy, x + pad_left_b + length, cy + bar_h],
                               fill=color)
                if item.get("show_values") and len(series) == 1 and len(categories) <= 8:
                    text = f"{value:g}"
                    draw.text((x + pad_left_b + length + 6 * scale, cy + bar_h / 2), text,
                              font=font, fill=ink, anchor="lm")
        return

    if kind in ("column", "bar"):
        n = len(categories)
        band = plot_w / max(n, 1)
        bar_w = band * 0.62 / max(len(series), 1)
        for si, s in enumerate(series):
            color = primary if si == 0 else accent
            for i, value in enumerate(s["values"]):
                cx = x + pad_left + band * (i + 0.19) + bar_w * si
                top, bottom = ypos(value), ypos(vmin)
                draw.rectangle([cx, min(top, bottom), cx + bar_w, max(top, bottom)], fill=color)
        draw.line([(x + pad_left, ypos(vmin)), (x + pad_left + plot_w, ypos(vmin))],
                  fill=_rgb_tuple(colors["hairline"]), width=max(1, scale))
        for i, category in enumerate(categories):
            cx = x + pad_left + band * (i + 0.5)
            label_w = font.getlength(str(category))
            draw.text((cx - label_w / 2, y + h - pad_bottom * 0.85), str(category), font=font,
                      fill=label_color)
    else:  # line
        step = plot_w / max(len(categories) - 1, 1)
        for si, s in enumerate(series):
            color = primary if si == 0 else accent
            points = [(x + pad_left + step * i, ypos(v)) for i, v in enumerate(s["values"]) if v is not None]
            if len(points) > 1:
                draw.line(points, fill=color, width=max(2, 2 * scale))
            for px_, py_ in points:
                draw.ellipse([px_ - 3 * scale, py_ - 3 * scale, px_ + 3 * scale, py_ + 3 * scale], fill=color)
        draw.line([(x + pad_left, ypos(vmin)), (x + pad_left + plot_w, ypos(vmin))],
                  fill=_rgb_tuple(colors["hairline"]), width=max(1, scale))
        for i, category in enumerate(categories):
            cx = x + pad_left + step * i
            label_w = font.getlength(str(category))
            draw.text((cx - label_w / 2, y + h - pad_bottom * 0.85), str(category), font=font,
                      fill=label_color)
    if item.get("show_values") and len(series) == 1 and len(categories) <= 8:
        for i, value in enumerate(series[0]["values"]):
            cx = x + pad_left + (plot_w / max(len(categories), 1)) * (i + 0.5) if kind in ("column", "bar") \
                else x + pad_left + (plot_w / max(len(categories) - 1, 1)) * i
            text = f"{value:g}"
            draw.text((cx - font.getlength(text) / 2, ypos(value) - 14 * scale), text, font=font, fill=ink)


def _ladder(policy, size: float) -> str:
    scale = policy.req("typography.scale")
    return min(scale, key=lambda k: abs(float(scale[k]) - float(size)))


# ──────────────────────────────────────────────────────────────────────────
# 渲染整稿 + 对照
# ──────────────────────────────────────────────────────────────────────────


def pixel_diff(path_a, path_b, policy) -> dict:
    from PIL import Image, ImageChops
    a = Image.open(path_a).convert("RGB")
    b = Image.open(path_b).convert("RGB")
    if a.size != b.size:
        return {"status": "SIZE_MISMATCH", "mismatch_ratio": 1.0,
                "sizes": [list(a.size), list(b.size)]}
    grid = int(policy.opt("render.diff_grid", 4))
    if grid > 1:
        a = a.resize((a.width // grid, a.height // grid), Image.BOX)
        b = b.resize((b.width // grid, b.height // grid), Image.BOX)
    diff = ImageChops.difference(a, b).convert("L")
    hist = diff.histogram()
    changed = sum(hist[24:])
    total = a.width * a.height
    return {"status": "OK", "mismatch_ratio": round(changed / total, 6),
            "changed_pixels": changed, "total_pixels": total}


def contact_sheet(paths: list, out_path, policy, labels: list | None = None) -> str:
    from PIL import Image, ImageDraw, ImageFont
    if not paths:
        raise RenderError("没有可拼图的页面")
    columns = int(policy.opt("render.contact_sheet_columns", 2))
    thumbs = [Image.open(p) for p in paths]
    gap = 16
    label_h = 26
    cell_w = max(t.width for t in thumbs)
    cell_h = max(t.height for t in thumbs) + label_h
    rows = (len(thumbs) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * cell_w + (columns + 1) * gap,
                              rows * cell_h + (rows + 1) * gap), (240, 238, 233))
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype(policy.req("typography.font_files.latin")[0], 16)
    except (OSError, IndexError):
        font = ImageFont.load_default()
    for i, thumb in enumerate(thumbs):
        r, c = divmod(i, columns)
        x = gap + c * (cell_w + gap)
        y = gap + r * (cell_h + gap)
        sheet.paste(thumb, (x, y))
        text = labels[i] if labels and i < len(labels) else Path(paths[i]).stem
        draw.text((x + 4, y + thumb.height + 6), text, fill=(60, 62, 66), font=font)
    target = Path(out_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(target, "PNG", optimize=True)
    return str(target)
