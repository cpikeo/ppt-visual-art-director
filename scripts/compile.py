#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plan.json → 原生可编辑 PPTX。

用法：
    python scripts/compile.py plan.json out.pptx [--assets DIR] [--report report.json]

只做确定性执行：校验契约、按 plan 建对象、写出文件。不做审美判断，不自动缩字、
换色、补数据或联网取图。硬错误（E-）在写文件前阻断，永不覆盖既有交付物；
运行期异常另写 <out>.error.pptx 供排查，该文件不是交付物。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
LIMITS: Dict[str, Any] = json.loads((HERE / "limits.json").read_text(encoding="utf-8-sig"))
SCHEMA: Dict[str, Any] = json.loads((HERE / "plan.schema.json").read_text(encoding="utf-8-sig"))

TEXT_TYPES = ("title", "body", "text", "number", "caption", "footer")
CHART_KINDS = ("bar", "line", "pie", "donut", "table")
EMU = 914400.0
PT_EMU = 12700.0


# ---------------------------------------------------------------- 诊断


@dataclass
class Issue:
    level: str          # HARD | SOFT | INFO
    code: str
    where: str
    message: str

    def __str__(self) -> str:
        return f"[{self.level}] {self.code} {self.where}: {self.message}"


class Report:
    def __init__(self) -> None:
        self.items: List[Issue] = []

    def add(self, level: str, code: str, where: str, message: str) -> None:
        self.items.append(Issue(level, code, where, message))

    def hard(self, code: str, where: str, message: str) -> None:
        self.add("HARD", code, where, message)

    def soft(self, code: str, where: str, message: str) -> None:
        self.add("SOFT", code, where, message)

    def info(self, code: str, where: str, message: str) -> None:
        self.add("INFO", code, where, message)

    def extend(self, other: "Report") -> None:
        self.items.extend(other.items)

    @property
    def blocked(self) -> bool:
        return any(i.level == "HARD" for i in self.items)

    def count(self, level: str) -> int:
        return sum(1 for i in self.items if i.level == level)

    def to_dict(self) -> List[Dict[str, str]]:
        return [{"level": i.level, "code": i.code, "where": i.where, "message": i.message} for i in self.items]


# ---------------------------------------------------------------- 载入


def load_plan(path: Path, rep: Report) -> Optional[Dict[str, Any]]:
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        rep.hard("E-SCHEMA", str(path), f"无法读取：{exc}")
        return None
    try:
        plan = json.loads(raw)
    except json.JSONDecodeError as exc:
        rep.hard("E-SCHEMA", str(path), f"JSON 语法错误：{exc}")
        return None
    try:
        from jsonschema import Draft202012Validator
    except ImportError:
        rep.hard("E-SCHEMA", "环境", "缺少 jsonschema：pip install -r requirements.txt")
        return plan
    validator = Draft202012Validator(SCHEMA)
    errors = sorted(validator.iter_errors(plan), key=lambda e: list(map(str, e.absolute_path)))
    for err in errors:
        where = "/".join(str(p) for p in err.absolute_path) or "(根)"
        rep.hard("E-SCHEMA", where, err.message[:300])
    return plan if isinstance(plan, dict) else None


# ---------------------------------------------------------------- 度量模型


def canvas(plan: Dict[str, Any]) -> Tuple[float, float]:
    return tuple(LIMITS["canvas"][plan.get("aspect_ratio") or LIMITS["canvas"]["default"]])  # type: ignore


def _em_width(ch: str) -> float:
    if ch in " \u00a0\t":
        return 0.30
    code = ord(ch)
    if code < 128:
        return 0.55 if ch.isalnum() else 0.34
    if ch in "，。、；：？！（）【】《》「」『』“”‘’—…·～":
        return 1.0
    if code >= 0x2E80:
        return 1.0
    return 0.6


def fit_lines(text: str, pt: float, box_w: float) -> int:
    usable = max(box_w - 2 * LIMITS["type"]["inset_in"]["x"], 0.05)
    cap = usable / (pt / 72.0)
    total = 0
    for seg in text.split("\n"):
        width = sum(_em_width(c) for c in seg)
        total += max(1, math.ceil(width / cap - 1e-9))
    return max(1, total)


def fit_height_in(text: str, pt: float, box_w: float, line_spacing: float) -> float:
    pitch = LIMITS["type"]["line_pitch_em"]
    return fit_lines(text, pt, box_w) * pt * pitch * line_spacing / 72.0


def hex_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rel_luminance(hex_color: str) -> float:
    def channel(v: int) -> float:
        s = v / 255.0
        return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4
    r, g, b = hex_rgb(hex_color)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(a: str, b: str) -> float:
    la, lb = rel_luminance(a), rel_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def mix_toward(fg: str, bg: str, t: float) -> str:
    """把 fg 向 bg 混合 t（0=fq, 1=bg），只用于图表退隐色。"""
    fa, ba = hex_rgb(fg), hex_rgb(bg)
    out = [round(a + (b - a) * t) for a, b in zip(fa, ba)]
    return "#" + "".join(f"{v:02X}" for v in out)


def is_dark(hex_color: str) -> bool:
    return rel_luminance(hex_color) < LIMITS["contrast"]["dark_background_luminance"]


def page_bg(page: Dict[str, Any]) -> str:
    return (page.get("background") or LIMITS["defaults"]["background"]).upper()


def deck_palette(plan: Dict[str, Any]) -> List[str]:
    return [c.upper() for c in (plan.get("palette") or LIMITS["defaults"]["palette"])]


def page_numbers(plan: Dict[str, Any]) -> Dict[str, Any]:
    cfg = dict(plan.get("page_numbers") or {})
    cfg.setdefault("corner", LIMITS["page_number"]["corner"])
    cfg.setdefault("start", LIMITS["page_number"]["start"])
    cfg.setdefault("hide", [])
    return cfg


def pn_box(plan: Dict[str, Any], page: Dict[str, Any]) -> Dict[str, float]:
    w, h = canvas(plan)
    box = LIMITS["page_number"]["box_in"]
    mx, my = LIMITS["page_number"]["margin_in"]["x"], LIMITS["page_number"]["margin_in"]["y"]
    corner = page_numbers(plan)["corner"]
    return {"x": round(w - mx - box["w"], 4),
            "y": round(my if corner == "top_right" else h - my - box["h"], 4),
            "w": box["w"], "h": box["h"]}


def ordered(page: Dict[str, Any]) -> List[Dict[str, Any]]:
    """叠放顺序 = 声明顺序：先写的在下。只有一条规则，没有第二个开关。"""
    return list(page["elements"])


# ---------------------------------------------------------------- 文本式样


def text_style(plan: Dict[str, Any], page: Dict[str, Any], el: Dict[str, Any]) -> Dict[str, Any]:
    kind = el["type"]
    size = float(el.get("font_size") or LIMITS["type"]["default_pt"].get(kind, 18))
    fonts = plan.get("fonts") or {}
    family = el.get("font") or fonts.get("heading" if kind in ("title", "number") else "body") \
        or LIMITS["defaults"]["fonts"]["heading" if kind in ("title", "number") else "body"]
    color = el.get("color")
    if not color:
        color = LIMITS["defaults"]["ink_on_dark"] if is_dark(page_bg(page)) else LIMITS["defaults"]["ink"]
    return {
        "size": size,
        "family": family,
        "color": (color or LIMITS["defaults"]["ink"]).upper(),
        "bold": bool(el.get("bold", LIMITS["type"]["bold"].get(kind, False))),
        "line_spacing": float(el.get("line_spacing") or LIMITS["type"]["line_spacing"].get(kind, 1.25)),
        "align": el.get("align", "left"),
        "valign": el.get("valign", "top"),
        "spacing": float(el.get("letter_spacing_pt") or 0),
    }


# ---------------------------------------------------------------- 校验


def _box(el: Dict[str, Any]) -> Tuple[float, float, float, float]:
    return float(el["x"]), float(el["y"]), float(el["w"]), float(el["h"])


def _intersection(a: Dict[str, Any], b: Dict[str, Any]) -> Tuple[float, float]:
    ax, ay, aw, ah = _box(a)
    bx, by, bw, bh = _box(b)
    ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return ix, iy


def substantive_overlap(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
    ix, iy = _intersection(a, b)
    if min(ix, iy) < LIMITS["geometry"]["overlap_min_edge_in"]:
        return False
    area = ix * iy
    smaller = min(float(a["w"]) * float(a["h"]), float(b["w"]) * float(b["h"]))
    return area >= LIMITS["geometry"]["overlap_area_sq_in"] and area / smaller >= LIMITS["geometry"]["overlap_ratio"]


def resolve_asset(asset_root: Path, image_id: str) -> Optional[Path]:
    direct = asset_root / image_id
    if direct.is_file():
        return direct
    for ext in (".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG", ".WEBP"):
        candidate = asset_root / (image_id + ext)
        if candidate.is_file():
            return candidate
    stem = Path(image_id).stem
    if asset_root.is_dir():
        for f in sorted(asset_root.iterdir()):
            if f.is_file() and f.stem == stem:
                return f
    return None


def validate(plan: Dict[str, Any], asset_root: Path) -> Report:
    rep = Report()
    w, h = canvas(plan)
    tol = LIMITS["canvas"]["bound_tolerance_in"]
    pages = plan.get("pages") or []
    seen_pages: set = set()
    assets = plan.get("assets") or {}
    palette = deck_palette(plan)
    pn_cfg = page_numbers(plan)

    for pi, page in enumerate(pages):
        pid = page.get("id", f"p{pi + 1}")
        loc = pid
        if pid in seen_pages:
            rep.hard("E-ID", loc, f"页面 id 重复：{pid}")
        seen_pages.add(pid)

        elements = page.get("elements") or []
        ids: set = set()
        for el in elements:
            eid = el.get("id", "?")
            eloc = f"{pid}/{eid}"
            if eid in ids:
                rep.hard("E-ID", eloc, f"元素 id 在页内重复：{eid}")
            ids.add(eid)

            x, y, ew, eh = _box(el)
            if x + ew <= 0 or y + eh <= 0 or x >= w or y >= h:
                rep.hard("E-GEOM", eloc, f"元素完全在画布外：x={x} y={y} w={ew} h={eh}")
            elif el["type"] in TEXT_TYPES and (x < -tol or y < -tol or x + ew > w + tol or y + eh > h + tol):
                rep.soft("E-GEOM", eloc, f"文本框伸出画布（{w}×{h}），边缘文字会被裁切")

            if el["type"] in TEXT_TYPES:
                content = el.get("content", "")
                runs = el.get("runs")
                if runs:
                    if "\n" in content:
                        rep.hard("E-REF", eloc, "runs 与 content 并存时 content 不能含换行")
                    if "".join(r.get("text", "") for r in runs) != content:
                        rep.hard("E-REF", eloc, "runs 拼接与 content 不一致")
                st = text_style(plan, page, el)
                need = fit_height_in(content, st["size"], ew, st["line_spacing"])
                if need > eh * LIMITS["type"]["fit_tolerance"]:
                    rep.soft("W-FIT", eloc,
                             f"按 {st['size']:g}pt 估算需 {need:.2f}in，框高 {eh:.2f}in：先编辑文案，再加宽/拆页")
                minimum = LIMITS["type"]["min_pt"].get(el["type"], 0)
                if st["size"] < minimum:
                    rep.soft("W-PT", eloc, f"{el['type']} {st['size']:g}pt 低于 {minimum}pt 下限")
                bg = resolve_bg(page, el)
                if bg is None:
                    rep.info("I-CONTRAST", eloc, "文字下方有图片，对比需看真实渲染")
                else:
                    limit = LIMITS["contrast"]["large"] if st["size"] >= LIMITS["contrast"]["large_pt"] \
                        else LIMITS["contrast"]["normal"]
                    ratio = contrast_ratio(st["color"], bg)
                    if ratio < limit:
                        rep.soft("W-CONTRAST", eloc, f"{st['color']} on {bg} = {ratio:.2f}:1，低于 {limit}:1")

            if el["type"] == "image":
                image_id = el.get("image_id", "")
                asset = assets.get(image_id)
                if asset is None:
                    rep.hard("E-REF", eloc, f"image_id 未在 assets 登记：{image_id}")
                else:
                    check_asset(eloc, image_id, asset, asset_root, rep)
                path = resolve_asset(asset_root, image_id)
                if path is None:
                    rep.hard("E-IMG", eloc, f"素材文件不存在：{asset_root / image_id}")
                else:
                    ppi, loss = image_probe(path, ew, eh, el.get("crop_focus"))
                    if ppi is not None and ppi < LIMITS["image"]["min_ppi"]:
                        rep.soft("W-PPI", eloc, f"有效 {ppi:.0f} ppi < {LIMITS['image']['min_ppi']}，投影会发虚")
                    if loss is not None and loss > LIMITS["image"]["max_crop_loss"]:
                        rep.soft("W-CROP", eloc, f"最终裁切丢掉了原图 {loss:.0%}，画幅或落位框大概选错了")

            if el["type"] == "chart":
                check_chart(eloc, el, palette, rep)

        if pid not in pn_cfg["hide"]:
            box = pn_box(plan, page)
            for el in elements:
                if el.get("overlap_allowed"):
                    continue  # 承载层在下，页码本来就压在它上面
                ix, iy = _intersection(el, box)
                if ix <= 0.02 or iy <= 0.02:
                    continue
                where = f"{pid}/{el.get('id')}"
                if el["type"] in TEXT_TYPES or el["type"] == "chart":
                    rep.hard("E-PAGE", where, "文字或图表与页码保留框相交；移开它，或整稿改用另一角区")
                elif el["type"] == "image":
                    rep.info("I-PAGE", where, "图片压到页码角区：用 page_number_tone 声明实际明暗，并确认页码可读")
                else:
                    rep.soft("E-PAGE", where, "色块压到页码角区：确认页码在它上面可读，必要时换角区")

    # 跨元素：实质重叠（叠放顺序即声明顺序）
    for page in pages:
        els = ordered(page)
        for i, lower in enumerate(els):
            for upper in els[i + 1:]:
                if substantive_overlap(lower, upper) and not lower.get("overlap_allowed"):
                    rep.hard("E-OVERLAP", f"{page.get('id')}/{upper.get('id')}",
                             f"与先声明的 {lower.get('id')} 实质重叠；若后者是承载层（背景图/蒙版/色块），"
                             f"在 {lower.get('id')} 上声明 overlap_allowed，否则移开")

    return rep


def resolve_bg(page: Dict[str, Any], el: Dict[str, Any]) -> Optional[str]:
    """文字实际落在什么颜色上：页面底色 + 其下覆盖 ≥90% 的色块；其下有图则无法静态判定。"""
    base = page_bg(page)
    target_area = max(float(el["w"]) * float(el["h"]), 1e-6)
    for below in ordered(page):
        if below is el:
            break
        ix, iy = _intersection(below, el)
        cover = ix * iy / target_area
        if below["type"] == "image" and cover >= 0.6:
            return None
        if below["type"] not in ("shape", "divider") or cover < 0.9:
            continue
        alpha = float(below.get("fill_alpha", 1.0))
        if alpha >= 0.999:
            base = str(below["color"]).upper()
        else:
            base = mix_toward(base, str(below["color"]).upper(), alpha)
    return base


def image_probe(path: Path, w: float, h: float, focus: Optional[Dict[str, float]]):
    """返回（有效 ppi, 被裁掉的原图比例）。裁掉太多说明画幅选错，不是"裁切很紧"。"""
    try:
        from PIL import Image
        with Image.open(path) as img:
            px_w, px_h = img.size
    except Exception:
        return None, None
    if w <= 0 or h <= 0 or not px_w or not px_h:
        return None, None
    left, right, top, bottom = _crop_geometry(px_w, px_h, w, h, focus)
    loss = 1 - (1 - left - right) * (1 - top - bottom)
    return min(px_w / w, px_h / h), loss


def check_asset(loc: str, image_id: str, asset: Dict[str, Any], asset_root: Path, rep: Report) -> None:
    rights = asset.get("rights", "pending")
    generated = bool(asset.get("generated"))
    if rights == "blocked":
        rep.hard("E-ASSET", loc, f"{image_id} 权利状态 blocked，不得使用")
    elif rights == "pending":
        rep.soft("E-ASSET", loc, f"{image_id} 权利未确认（pending）；交付前确认或改用已清权素材")
    if generated and asset.get("used_as_evidence"):
        rep.hard("E-ASSET", loc, f"{image_id} 是生成图，不能作证据")
    gen = asset.get("generation")
    if gen:
        if gen.get("status") != "generated":
            rep.hard("E-ASSET", loc, f"{image_id} 生成记录 status={gen.get('status')}，不得作为正式素材")
        path = resolve_asset(asset_root, image_id)
        gfile = resolve_asset(asset_root, gen.get("file", ""))
        if path and gfile and path.resolve() != gfile.resolve():
            rep.hard("E-ASSET", loc, f"{image_id} 与 generation.file={gen.get('file')} 不是同一物理文件")
        if path and gen.get("sha256"):
            actual = hashlib.sha256(path.read_bytes()).hexdigest()
            if actual.lower() != str(gen["sha256"]).lower():
                rep.hard("E-ASSET", loc, f"{image_id} 实际 SHA-256 与登记不符")


def check_chart(loc: str, el: Dict[str, Any], palette: List[str], rep: Report) -> None:
    chart = el["chart"]
    kind = chart["kind"]
    data = chart["data"]
    if kind == "table":
        if "headers" not in data:
            rep.hard("E-CHART", loc, "table 图表的 data 需为 headers/rows")
            return
        weights = chart.get("column_weights")
        if weights and len(weights) != len(data["headers"]):
            rep.hard("E-CHART", loc, f"column_weights 数量 {len(weights)} ≠ 列数 {len(data['headers'])}")
        for r, row in enumerate(data["rows"]):
            if len(row) != len(data["headers"]):
                rep.hard("E-CHART", loc, f"第 {r + 1} 行 {len(row)} 列 ≠ 表头 {len(data['headers'])} 列")
        return

    if "series" not in data:
        rep.hard("E-CHART", loc, "series 图表的 data 需为 categories/series")
        return
    cats = data["categories"]
    series = data["series"]
    for s in series:
        if len(s["values"]) != len(cats):
            rep.hard("E-CHART", loc, f"系列「{s['name']}」有 {len(s['values'])} 个值，类别 {len(cats)} 个")
    if kind in ("pie", "donut"):
        if len(series) != 1:
            rep.hard("E-CHART", loc, f"{kind} 只能单系列，现有 {len(series)} 个")
        for s in series:
            if any(v < 0 for v in s["values"]):
                rep.hard("E-CHART", loc, "饼/环不能有负值")
    if len(series) > 1:
        units = {s.get("unit") for s in series if s.get("unit")}
        if len(units) > 1 and not chart.get("shared_basis"):
            rep.hard("E-CHART", loc, f"多系列声明了不同单位 {sorted(units)}；共轴前写 shared_basis=true")
        if chart.get("unit") and any(s.get("unit") and s["unit"] != chart["unit"] for s in series):
            rep.hard("E-CHART", loc, "系列单位与图表级 unit 矛盾")
        if chart.get("period") and any(s.get("period") and s["period"] != chart["period"] for s in series):
            rep.hard("E-CHART", loc, "系列期间与图表级 period 矛盾")
    need = len(cats) if kind in ("pie", "donut") else len(series)
    if len(palette) < need:
        rep.soft("E-CHART", loc, f"palette 只有 {len(palette)} 色，需要 {need}；将循环取色")
    if chart.get("axis_labels") is False and not chart.get("data_labels"):
        rep.hard("E-CHART", loc, "axis_labels=false 时必须 data_labels=true，否则无读数")
    hi = chart.get("highlight_category")
    if hi is not None:
        if kind != "bar" or len(series) != 1:
            rep.hard("E-CHART", loc, "highlight_category 只用于单系列柱图")
        elif hi >= len(cats):
            rep.hard("E-CHART", loc, f"highlight_category {hi} 超出类别范围")


# ---------------------------------------------------------------- 绘图


def _apply_alpha(shape, alpha: float) -> None:
    from pptx.oxml.ns import qn
    from lxml import etree
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None:
        return
    clr = solid[0]
    for old in clr.findall(qn("a:alpha")):
        clr.remove(old)
    node = etree.SubElement(clr, qn("a:alpha"))
    node.set("val", str(int(round(alpha * 100000))))


def _ensure_family(font, family: str) -> None:
    """图表文字也要写东亚字体：只写 latin 时，中文会被渲染器换掉，整稿字族就断了。"""
    from pptx.oxml.ns import qn
    from lxml import etree
    font.name = family
    rPr = getattr(font, "_rPr", None)
    if rPr is None:
        return
    for tag in ("a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", family)


def _set_run_font(run, family: str, size: float, color: str, bold: bool, spacing: float = 0.0) -> None:
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree
    font = run.font
    font.size = Pt(size)
    font.bold = bold
    font.name = family
    font.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", family)
    if spacing:
        rPr.set("spc", str(int(round(spacing * 100))))


def add_text(slide, plan, page, el):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.text import MSO_AUTO_SIZE

    st = text_style(plan, page, el)
    box = slide.shapes.add_textbox(Inches(float(el["x"])), Inches(float(el["y"])),
                                   Inches(float(el["w"])), Inches(float(el["h"])))
    box.name = f"{page['id']}.{el['id']}"
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    inset = LIMITS["type"]["inset_in"]
    tf.margin_left = tf.margin_right = Inches(inset["x"])
    tf.margin_top = tf.margin_bottom = Inches(inset["y"])
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[st["valign"]]
    align = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[st["align"]]

    paragraphs = el["content"].split("\n")
    runs_override = el.get("runs")
    for idx, seg in enumerate(paragraphs):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = st["line_spacing"]
        if el.get("space_after_pt"):
            para.space_after = Pt(float(el["space_after_pt"]))
        pieces = runs_override if (runs_override and idx == 0) else [{"text": seg}]
        for piece in pieces:
            run = para.add_run()
            run.text = piece.get("text", "")
            _set_run_font(
                run,
                piece.get("font") or st["family"],
                float(piece.get("font_size") or st["size"]),
                str(piece.get("color") or st["color"]).upper(),
                bool(piece.get("bold", st["bold"])),
                float(piece.get("letter_spacing_pt", st["spacing"])),
            )


def add_shape(slide, page, el):
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor

    kind = {"rect": MSO_SHAPE.RECTANGLE, "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE, "oval": MSO_SHAPE.OVAL}[
        el.get("shape_kind", "rect")]
    shape = slide.shapes.add_shape(kind, Inches(float(el["x"])), Inches(float(el["y"])),
                                   Inches(float(el["w"])), Inches(float(el["h"])))
    shape.name = f"{page['id']}.{el['id']}"
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(el["color"]).lstrip("#").upper())
    shape.line.fill.background()
    shape.shadow.inherit = False
    if el.get("fill_alpha") and float(el["fill_alpha"]) < 1:
        _apply_alpha(shape, float(el["fill_alpha"]))
    if shape.has_text_frame:
        shape.text_frame.text = ""


def _crop_geometry(src_w: float, src_h: float, w: float, h: float, focus: Optional[Dict[str, float]]):
    fx = 0.5 if not focus else float(focus.get("x", 0.5))
    fy = 0.5 if not focus else float(focus.get("y", 0.5))
    src_ratio, dst_ratio = src_w / src_h, w / h
    l = r = t = b = 0.0
    if src_ratio > dst_ratio:
        trim = 1 - (dst_ratio / src_ratio)
        l, r = trim * fx, trim * (1 - fx)
    elif src_ratio < dst_ratio:
        trim = 1 - (src_ratio / dst_ratio)
        t, b = trim * fy, trim * (1 - fy)
    return l, r, t, b


def add_image(slide, page, el, asset_root: Path):
    from pptx.util import Inches
    from PIL import Image

    path = resolve_asset(asset_root, el["image_id"])
    if path is None:
        raise FileNotFoundError(f"素材缺失 {el['image_id']}")
    with Image.open(path) as img:
        src_w, src_h = img.size
    pic = slide.shapes.add_picture(str(path), Inches(float(el["x"])), Inches(float(el["y"])),
                                   Inches(float(el["w"])), Inches(float(el["h"])))
    pic.name = f"{page['id']}.{el['id']}"
    l, r, t, b = _crop_geometry(src_w, src_h, float(el["w"]), float(el["h"]), el.get("crop_focus"))
    if any(v > 0 for v in (l, r, t, b)):
        pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom = l, r, t, b


def _series_colors(palette: List[str], count: int) -> List[str]:
    if not palette:
        palette = LIMITS["defaults"]["palette"]
    return [palette[i % len(palette)] for i in range(count)]


def _deemphasize(bg: str, kind: str) -> str:
    base = LIMITS["chart"]["deemphasis"]["dark" if is_dark(bg) else "light"]
    floor = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    for step in range(0, 21):
        candidate = mix_toward(base, bg, step / 20 * 0.55)
        if contrast_ratio(candidate, bg) >= floor:
            return candidate
    return base


def _left_align_chart_title(obj) -> None:
    """把图表标题框固定到图表区左上：默认 auto layout 会把它居中，与左对齐的轴脱节。"""
    from pptx.oxml.ns import qn
    from lxml import etree
    title = obj._chartSpace.find(qn("c:chart")).find(qn("c:title"))
    layout = None if title is None else title.find(qn("c:layout"))
    if layout is None:
        return
    for child in list(layout):
        layout.remove(child)
    manual = etree.SubElement(layout, qn("c:manualLayout"))
    for tag, val in (("c:layoutTarget", "inner"), ("c:xMode", "edge"), ("c:yMode", "edge")):
        node = etree.SubElement(manual, qn(tag))
        node.set("val", val)
    for tag, val in (("c:x", "0"), ("c:y", "0.02"), ("c:w", "1"), ("c:h", "0.14")):
        node = etree.SubElement(manual, qn(tag))
        node.set("val", val)


def add_chart(slide, plan, page, el, palette: List[str]):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION

    chart = el["chart"]
    kind = chart["kind"]
    bg = page_bg(page)
    tone = "dark" if is_dark(bg) else "light"
    style = LIMITS["chart"]["style"][tone]
    fonts = plan.get("fonts") or {}
    family = el.get("font") or fonts.get("body") or LIMITS["defaults"]["fonts"]["body"]
    x, y, w, h = _box(el)

    if kind == "table":
        add_table(slide, plan, page, el, style, family)
        return

    data = CategoryChartData()
    data.categories = data_categories = chart["data"]["categories"]
    for s in chart["data"]["series"]:
        data.add_series(str(s["name"]), [float(v) for v in s["values"]])

    xl = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "line": XL_CHART_TYPE.LINE,
          "pie": XL_CHART_TYPE.PIE, "donut": XL_CHART_TYPE.DOUGHNUT}[kind]
    frame = slide.shapes.add_chart(xl, Inches(x), Inches(y), Inches(w), Inches(h), data)
    frame.name = f"{page['id']}.{el['id']}"
    obj = frame.chart
    _ensure_family(obj.font, family)
    obj.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])

    series = chart["data"]["series"]
    colors = _series_colors(palette, len(series) if kind not in ("pie", "donut") else len(data_categories))
    plot = obj.plots[0]

    if kind in ("bar", "line"):
        obj.has_legend = len(series) > 1
        if obj.has_legend:
            obj.legend.position = XL_LEGEND_POSITION.BOTTOM
            obj.legend.include_in_layout = False
            obj.legend.font.size = Pt(LIMITS["chart"]["font_pt"]["legend"])
            _ensure_family(obj.legend.font, family)
            obj.legend.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        for i, s in enumerate(plot.series):
            if kind == "bar":
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = RGBColor.from_string(colors[i].lstrip("#"))
                s.format.line.fill.background()
            else:
                s.format.line.color.rgb = RGBColor.from_string(colors[i].lstrip("#"))
                s.format.line.width = Pt(2.25)
                s.smooth = False
        if kind == "bar":
            plot.gap_width = LIMITS["chart"]["gap_width"]
            hi = chart.get("highlight_category")
            if hi is not None:
                dim = _deemphasize(bg, "bar")
                points = plot.series[0].points
                for j, point in enumerate(points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = RGBColor.from_string(
                        (colors[0] if j == hi else dim).lstrip("#"))
                    point.format.line.fill.background()
        axis_on = chart.get("axis_labels", True)
        va = obj.value_axis
        va.has_major_gridlines = axis_on
        if axis_on:
            va.major_gridlines.format.line.color.rgb = RGBColor.from_string(style["grid"].lstrip("#"))
            va.major_gridlines.format.line.width = Pt(0.75)
            va.tick_labels.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])
            _ensure_family(va.tick_labels.font, family)
            va.tick_labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
            va.format.line.color.rgb = RGBColor.from_string(style["rule"].lstrip("#"))
        else:
            va.visible = False
            va.has_major_gridlines = False
        ca = obj.category_axis
        ca.has_major_gridlines = False
        ca.tick_labels.font.size = Pt(LIMITS["chart"]["font_pt"]["axis"])
        _ensure_family(ca.tick_labels.font, family)
        ca.tick_labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        ca.format.line.color.rgb = RGBColor.from_string(style["rule"].lstrip("#"))
        ca.tick_labels.position = XL_TICK_LABEL_POSITION.LOW
    else:
        obj.has_legend = True
        obj.legend.position = XL_LEGEND_POSITION.BOTTOM
        obj.legend.include_in_layout = False
        obj.legend.font.size = Pt(LIMITS["chart"]["font_pt"]["legend"])
        _ensure_family(obj.legend.font, family)
        obj.legend.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        points = plot.series[0].points
        for j, point in enumerate(points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string(colors[j % len(colors)].lstrip("#"))
            point.format.line.color.rgb = RGBColor.from_string(bg.lstrip("#"))

    if chart.get("data_labels"):
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.number_format = "#,##0.###"
        labels.number_format_is_linked = False
        labels.font.size = Pt(LIMITS["chart"]["font_pt"]["label"])
        _ensure_family(labels.font, family)
        labels.font.color.rgb = RGBColor.from_string(style["body"].lstrip("#"))
        labels.position = XL_LABEL_POSITION.OUTSIDE_END if kind == "bar" else XL_LABEL_POSITION.ABOVE

    title = chart.get("title")
    if chart.get("show_title", True) and title:
        from pptx.enum.text import PP_ALIGN
        obj.has_title = True
        frame_title = obj.chart_title
        frame_title.text_frame.text = title
        for para in frame_title.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT
            for run in para.runs:
                _set_run_font(run, family, LIMITS["chart"]["font_pt"]["title"], style["title"], False)
        _left_align_chart_title(obj)
    else:
        obj.has_title = False


def add_table(slide, plan, page, el, style, family):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR
    from pptx.oxml.ns import qn
    from lxml import etree

    chart = el["chart"]
    headers = chart["data"]["headers"]
    rows = chart["data"]["rows"]
    x, y, w, h = _box(el)
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = f"{page['id']}.{el['id']}"
    table = shape.table
    table.first_row = False
    table.horz_banding = False
    from pptx.oxml.ns import qn as _qn
    from lxml import etree as _etree
    tblPr = table._tbl.find(_qn("a:tblPr"))
    if tblPr is not None:
        old = tblPr.find(_qn("a:tableStyleId"))
        if old is not None:
            tblPr.remove(old)
        style_id = _etree.SubElement(tblPr, _qn("a:tableStyleId"))
        style_id.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid：横线由本函数画
    weights = chart.get("column_weights") or [1] * len(headers)
    total = sum(weights)
    for i, weight in enumerate(weights):
        table.columns[i].width = Inches(w * weight / total)

    body_pt = LIMITS["chart"]["font_pt"]["table"]
    header_ink = style["title"]
    body_ink = style["body"]

    def paint(cell, text, ink, bold, fill=None):
        cell.text = str(text)
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(fill.lstrip("#"))
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                _set_run_font(run, family, body_pt, ink, bold)

    for j, head in enumerate(headers):
        paint(table.cell(0, j), head, header_ink, True)
    for r, row in enumerate(rows):
        for j, value in enumerate(row):
            paint(table.cell(r + 1, j), value, body_ink, False)

    # 细底线：比灰底表格干净，也更像编辑设计
    for r in range(len(rows) + 1):
        for c in range(len(headers)):
            cell = table.cell(r, c)
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnT", "a:lnB", "a:lnL", "a:lnR"):
                for node in tcPr.findall(qn(tag)):
                    tcPr.remove(node)
                if tag in ("a:lnT", "a:lnB"):
                    line = etree.SubElement(tcPr, qn(tag))
                    line.set("w", str(int(0.75 * PT_EMU)))
                    line.set("cmpd", "sng")
                    fill = etree.SubElement(line, qn("a:solidFill"))
                    color = etree.SubElement(fill, qn("a:srgbClr"))
                    color.set("val", (style["rule"] if r in (0, 1) else style["grid"]).lstrip("#").upper())


def add_page_number(slide, plan, page, index, total):
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    cfg = page_numbers(plan)
    if page["id"] in cfg["hide"]:
        return
    box = pn_box(plan, page)
    tone = page.get("page_number_tone") or ("dark" if is_dark(page_bg(page)) else "light")
    ink = LIMITS["page_number"]["ink"][tone]
    shape = slide.shapes.add_textbox(Inches(box["x"]), Inches(box["y"]), Inches(box["w"]), Inches(box["h"]))
    shape.name = "_page_number"
    tf = shape.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    run = para.add_run()
    run.text = str(cfg["start"] + index)
    _set_run_font(run, (plan.get("fonts") or {}).get("body") or LIMITS["defaults"]["fonts"]["body"],
                  LIMITS["page_number"]["font_pt"], ink, False)


# ---------------------------------------------------------------- 主流程


def build(plan: Dict[str, Any], asset_root: Path, out: Path, rep: Report) -> None:
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    w, h = canvas(plan)
    prs = Presentation()
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    blank = prs.slide_layouts[6]
    palette = deck_palette(plan)
    pages = plan["pages"]

    for idx, page in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        bg = page_bg(page)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(bg.lstrip("#"))
        for el in ordered(page):
            kind = el["type"]
            if kind in TEXT_TYPES:
                add_text(slide, plan, page, el)
            elif kind in ("shape", "divider"):
                add_shape(slide, page, el)
            elif kind == "image":
                add_image(slide, page, el, asset_root)
            elif kind == "chart":
                add_chart(slide, plan, page, el, palette)
        add_page_number(slide, plan, page, idx, len(pages))

    prs.core_properties.title = plan.get("title", "")
    prs.core_properties.author = "ppt-visual-art-director"
    prs.core_properties.comments = (plan.get("visual_concept") or "")[:500]
    tmp = out.with_suffix(".partial.pptx")
    prs.save(str(tmp))
    os.replace(tmp, out)


def write_error_deck(out: Path, message: str) -> None:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(5.9))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = "编译失败，此文件不是交付物。\n\n" + message
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(14 if run.text.startswith("编译失败") else 11)
    prs.save(str(out))


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="plan.json → 原生可编辑 PPTX")
    parser.add_argument("plan")
    parser.add_argument("out")
    parser.add_argument("--assets", default=None, help="素材目录；缺省为 plan 同级 assets/ 或 plan 所在目录")
    parser.add_argument("--report", default=None, help="把诊断写成 JSON")
    args = parser.parse_args(argv)

    plan_path = Path(args.plan).resolve()
    out_path = Path(args.out).resolve()
    asset_root = Path(args.assets).resolve() if args.assets else (
        plan_path.parent / "assets" if (plan_path.parent / "assets").is_dir() else plan_path.parent)

    rep = Report()
    plan = load_plan(plan_path, rep)
    if plan is None or rep.blocked:
        for issue in rep.items:
            print(issue)
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        return 2

    rep.extend(validate(plan, asset_root))
    for issue in rep.items:
        print(issue)
    if rep.blocked:
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        if args.report:
            Path(args.report).write_text(json.dumps({"blocked": True, "issues": rep.to_dict()},
                                                    ensure_ascii=False, indent=2), encoding="utf-8")
        return 2

    try:
        build(plan, asset_root, out_path, rep)
    except Exception as exc:  # 运行期异常：不覆盖正式输出
        error_path = out_path.with_suffix(".error.pptx")
        write_error_deck(error_path, f"{type(exc).__name__}: {exc}")
        print(f"[HARD] E-RUNTIME 运行期失败：{type(exc).__name__}: {exc}\n诊断文件：{error_path}（不是交付物）")
        return 3

    print(f"\n已写出 {out_path}（{len(plan['pages'])} 页，{canvas(plan)[0]}×{canvas(plan)[1]} in）"
          f"  警告 {rep.count('SOFT')}，提示 {rep.count('INFO')}")
    if args.report:
        Path(args.report).write_text(json.dumps({
            "blocked": False,
            "output": str(out_path),
            "pages": len(plan["pages"]),
            "sha256": hashlib.sha256(out_path.read_bytes()).hexdigest(),
            "issues": rep.to_dict(),
        }, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
