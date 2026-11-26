#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设计系统与版式配方：outline.json → plan.json 初稿。

判断仍归艺术总监（概念、取舍、文案、哪页用哪种讲法）；这里只做**算术**：
  · 色彩系统：由一个种子色按 CIELAB 推出环境/结构/文字/强调/数据色，并保证对比度与灰度可辨；
  · 字号阶梯：按观看场景（hall / meeting / screen）取一整套；
  · 版式配方：12 栏网格 + 同一条左轴，框高用 model.fit_height_in 算出，不靠猜。

用法：
    python scripts/scaffold.py system --seed "#1F3A5F" [--mood warm|cool|auto] [--scene meeting]
    python scripts/scaffold.py init outline.json plan.json [--force]

生成的 plan 是起点不是终点：之后照常 compile → verify → 看图，再手改几何与文案。
配方清单与 outline 写法见 references/build.md「用 scaffold 起稿」一节。
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import model as M

# ---------------------------------------------------------------- 色彩系统（CIELAB）

_WHITE = (0.95047, 1.0, 1.08883)


def _lin(c: float) -> float:
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _gam(c: float) -> float:
    return 12.92 * c if c <= 0.0031308 else 1.055 * (c ** (1 / 2.4)) - 0.055


def hex_to_lch(hex_color: str) -> Tuple[float, float, float]:
    r, g, b = (_lin(v / 255.0) for v in M.hex_rgb(hex_color))
    x = 0.4124564 * r + 0.3575761 * g + 0.1804375 * b
    y = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b
    z = 0.0193339 * r + 0.1191920 * g + 0.9503041 * b
    f = lambda t: t ** (1 / 3) if t > 216 / 24389 else (24389 / 27 * t + 16) / 116
    fx, fy, fz = f(x / _WHITE[0]), f(y / _WHITE[1]), f(z / _WHITE[2])
    lab_l, a, b2 = 116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)
    return lab_l, math.hypot(a, b2), math.degrees(math.atan2(b2, a)) % 360


def _lch_rgb(lab_l: float, c: float, h: float) -> Optional[Tuple[float, float, float]]:
    a, b = c * math.cos(math.radians(h)), c * math.sin(math.radians(h))
    fy = (lab_l + 16) / 116
    fx, fz = fy + a / 500, fy - b / 200
    inv = lambda t: t ** 3 if t ** 3 > 216 / 24389 else (116 * t - 16) / (24389 / 27)
    x, y, z = inv(fx) * _WHITE[0], inv(fy) * _WHITE[1], inv(fz) * _WHITE[2]
    lin = (3.2404542 * x - 1.5371385 * y - 0.4985314 * z,
           -0.9692660 * x + 1.8760108 * y + 0.0415560 * z,
           0.0556434 * x - 0.2040259 * y + 1.0572252 * z)
    if any(v < -1e-4 or v > 1 + 1e-4 for v in lin):
        return None
    return tuple(_gam(min(max(v, 0.0), 1.0)) for v in lin)  # type: ignore


def lch_to_hex(lab_l: float, c: float, h: float) -> str:
    """超出色域时降低彩度而不是截断通道，色相才不会漂。"""
    lab_l = min(max(lab_l, 0.0), 100.0)
    rgb = None
    while c >= 0:
        rgb = _lch_rgb(lab_l, c, h)
        if rgb:
            break
        c -= 1.0
    rgb = rgb or (0.0, 0.0, 0.0)
    return "#" + "".join(f"{int(round(min(max(v, 0), 1) * 255)):02X}" for v in rgb)


def _until(contrast_to: str, minimum: float, lab_l: float, c: float, h: float, step: float) -> str:
    """沿明度方向走，直到对 contrast_to 的对比度达标。"""
    for _ in range(100):
        color = lch_to_hex(lab_l, c, h)
        if M.contrast_ratio(color, contrast_to) >= minimum:
            return color
        lab_l += step
    return lch_to_hex(max(0.0, min(100.0, lab_l)), c, h)


MOOD_HUE = {"warm": 75.0, "cool": 255.0}
SCENES = {  # 论断标题 / 正文 / 关键数字 / 注释 / 封面大标题（pt）——design.md 字号阶梯的默认档
    "hall": {"title": 52, "body": 26, "number": 150, "caption": 16, "hero": 72},
    "meeting": {"title": 36, "body": 20, "number": 120, "caption": 14, "hero": 60},
    "screen": {"title": 28, "body": 16, "number": 80, "caption": 13, "hero": 44},
}
L_STEPS = [44, 57, 24, 34, 72.5, 14]            # 浅底数据色明度：首色贴近强调色；名义间距 ≥ 10 L*，留出 sRGB 取整与色域裁剪的损耗，实测两两 ≥ limits.chart.min_series_delta_lstar（9.5）；前 4 色对底 ≥ 3:1（折线/细线不发虚），后 2 色只作面积填充
L_STEPS_DARK = [65, 55, 45, 75, 85, 35]         # 深底数据色明度：名义间距 10 L*（同上）；前 4 色对底 ≥ 3:1，不用近白色（近白会抢过强调色）
HUE_OFFSETS = [0, 190, 40, 240, 90, 320]        # 首色=种子；其余在冷暖两侧错开，彩度收低，避免与强调色“撞色”
CHROMA_TAIL = [22, 28, 20, 26, 22]


def design_system(seed: str, mood: str = "auto", scene: str = "meeting",
                  overrides: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """种子色 → 色彩系统。`overrides` 让判断凌驾于推导：
    ground.light/dark（页底色，如纸色、墨绿、近黑蓝）、accent.light/dark（唯一强调色）、data/data_dark（数据色序列）。
    给了覆盖，文字墨阶与结构色仍按新底色重新推导并保证对比度；底色不够浅/深会抛 ValueError。"""
    ov = overrides or {}
    ground = ov.get("ground") or {}
    seed_l, seed_c, seed_h = hex_to_lch(seed)
    tint_h = MOOD_HUE.get(mood, seed_h)
    tint_c = 3.0 if mood in MOOD_HUE else 2.0

    def neutral(lab_l: float, c_mul: float = 1.0) -> str:
        return lch_to_hex(lab_l, tint_c * c_mul, tint_h)

    light_bg = ground.get("light") or neutral(97.5)
    dark_bg = ground.get("dark") or neutral(12.0, 1.5)
    if ground.get("light") and hex_to_lch(light_bg)[0] < 70:
        raise ValueError(f"ground.light {light_bg} 明度 L*<70，不是浅底；深底请写 ground.dark")
    if ground.get("dark") and hex_to_lch(dark_bg)[0] > 35:
        raise ValueError(f"ground.dark {dark_bg} 明度 L*>35，不是深底；浅底请写 ground.light")
    lbg_l, lbg_c, lbg_h = hex_to_lch(light_bg)
    dbg_l, dbg_c, dbg_h = hex_to_lch(dark_bg)
    if ground.get("light"):
        tint_h, tint_c = lbg_h, min(max(lbg_c, 1.5), 8.0)
    light = {
        "bg": light_bg, "bg_alt": lch_to_hex(max(lbg_l - 4.0, 0), tint_c, tint_h) if ground.get("light") else neutral(93.5),
        "ink": _until(light_bg, 12.0, 14, tint_c, tint_h, -1),
        "ink2": _until(light_bg, 7.0, 36, tint_c, tint_h, -1),
        "muted": _until(light_bg, 4.8, 52, tint_c, tint_h, -1),
        "rule": lch_to_hex(max(lbg_l - 13.5, 0), tint_c, tint_h) if ground.get("light") else neutral(84),
        "quiet": _until(light_bg, 3.0, 72, tint_c, tint_h, -1),   # 大字配角：≥24pt 够 3:1，又与主角拉开明度
        "accent": _until(light_bg, 4.5, min(seed_l, 55), max(seed_c, 30), seed_h, -1),
    }
    if ground.get("dark"):
        tint_h, tint_c = dbg_h, min(max(dbg_c, 1.5), 8.0)
    dark = {
        "bg": dark_bg, "bg_alt": lch_to_hex(dbg_l + 6.0, tint_c * 1.5, tint_h) if ground.get("dark") else neutral(18.0, 1.5),
        "ink": _until(dark_bg, 12.0, 95, tint_c, tint_h, 1),
        "ink2": _until(dark_bg, 7.0, 78, tint_c, tint_h, 1),
        "muted": _until(dark_bg, 4.8, 62, tint_c, tint_h, 1),
        "rule": lch_to_hex(dbg_l + 18.0, tint_c * 1.5, tint_h) if ground.get("dark") else neutral(30, 1.5),
        "quiet": _until(dark_bg, 3.0, 40, tint_c, tint_h, 1),
        "accent": _until(dark_bg, 4.5, max(seed_l, 62), max(seed_c, 30), seed_h, 1),
    }
    acc = ov.get("accent") or {}
    if acc.get("light"):
        al, ac_, ah = hex_to_lch(acc["light"])
        light["accent"] = _until(light_bg, 4.5, al, ac_, ah, -1)
    if acc.get("dark"):
        al, ac_, ah = hex_to_lch(acc["dark"])
        dark["accent"] = _until(dark_bg, 4.5, al, ac_, ah, 1)
    hues = [seed_h + o for o in HUE_OFFSETS]
    chroma = [max(min(seed_c, 70), 35)] + CHROMA_TAIL  # 只有首色是饱和的，其余收彩度，整套才耐看
    data = [lch_to_hex(L_STEPS[i], chroma[i], hues[i] % 360) for i in range(6)]
    data_dark = [lch_to_hex(L_STEPS_DARK[i], chroma[i], hues[i] % 360) for i in range(6)]
    data = list(ov.get("data") or data)
    data_dark = list(ov.get("data_dark") or data_dark)
    return {
        "scene": scene, "sizes": SCENES[scene], "light": light, "dark": dark,
        "data": data, "data_dark": data_dark,
        "checks": {
            "light_ink_on_bg": round(M.contrast_ratio(light["ink"], light["bg"]), 1),
            "light_muted_on_bg": round(M.contrast_ratio(light["muted"], light["bg"]), 1),
            "light_accent_on_bg": round(M.contrast_ratio(light["accent"], light["bg"]), 1),
            "dark_ink_on_bg": round(M.contrast_ratio(dark["ink"], dark["bg"]), 1),
            "dark_accent_on_bg": round(M.contrast_ratio(dark["accent"], dark["bg"]), 1),
            "accent_vs_quiet_dL_light": round(abs(hex_to_lch(light["accent"])[0] - hex_to_lch(light["quiet"])[0]), 1),
            "data_gray_conflicts": len(M.gray_conflicts(data)),
            "data_dark_gray_conflicts": len(M.gray_conflicts(data_dark)),
            "data_min_contrast_on_light": round(min(M.contrast_ratio(c, light["bg"]) for c in data), 1),
            "data_min_contrast_on_dark": round(min(M.contrast_ratio(c, dark["bg"]) for c in data_dark), 1),
        },
    }


# ---------------------------------------------------------------- 网格与度量

W, H = M.canvas()   # 16:9 唯一画布（1280×720 px @96 px/in），来自 limits.json
MX, MY, GUT, COLS = 0.8, 0.7, 0.25, 12
COL_W = (W - 2 * MX - GUT * (COLS - 1)) / COLS
_PN = M.LIMITS["page_number"]
PN_LEFT = W - _PN["margin_in"]["x"] - _PN["box_in"]["w"]   # 页码框左缘：页脚文字不得越过
TOP = MY + 0.34          # 页内内容起点：母题短横线在 MY，页题从这里开始
FOOT = H - 1.0           # 页脚上沿：note 与页码在它下面
FIT_SAFETY = 0.86        # 自适应大字号时的宽度安全系数：渲染实测，大字（≥60pt）的真实宽度比估算宽 10–15%


def col(start: int, span: int) -> Tuple[float, float]:
    """第 start 列（从 1 起）起、跨 span 列的 (x, w)。"""
    return MX + (start - 1) * (COL_W + GUT), span * COL_W + (span - 1) * GUT


def make_grid(spec: Any) -> Tuple[Any, Any, int, int]:
    """页级网格：网格是选出来的构图工具，不是固定框架。
    spec 缺省 → 12 栏 / 6 行；整数 N → N 栏；对象 {cols, rows, margin, margin_y, gutter, gutter_y}；
    `"none"` 或 0 → 不用网格（元素自己写 x y w h）。返回 (col, row, 栏数, 行数)。"""
    if spec in ("none", 0, False):
        return None, None, 0, 0
    g = {"cols": spec} if isinstance(spec, int) and not isinstance(spec, bool) else dict(spec or {})
    n, rows = int(g.get("cols", COLS)), int(g.get("rows", 6))
    mx, my = float(g.get("margin", MX)), float(g.get("margin_y", MY))
    gx, gy = float(g.get("gutter", GUT)), float(g.get("gutter_y", 0.2))
    cw, rh = (W - 2 * mx - gx * (n - 1)) / n, (H - 2 * my - gy * (rows - 1)) / rows

    def gcol(start: int, span: int) -> Tuple[float, float]:
        return mx + (start - 1) * (cw + gx), span * cw + (span - 1) * gx

    def grow(start: int, span: int) -> Tuple[float, float]:
        return my + (start - 1) * (rh + gy), span * rh + (span - 1) * gy

    return gcol, grow, n, rows


SCALE_FLOOR_PT = 28   # 尺度隐喻里较小数字的字号下限
BASELINE_EM = 0.20    # valign=bottom 的数字框：基线在框底之上的距离（em）。LibreOffice 渲染实测标定；PowerPoint 未实机验证


class Ctx:
    def __init__(self, outline: Dict[str, Any], asset_root: Optional[Path] = None):
        self.asset_root = asset_root
        design = outline.get("design") or {}
        overrides = {k: design[k] for k in ("ground", "accent", "data", "data_dark") if design.get(k)}
        self.system = design_system(design.get("seed", "#1F3A5F"), design.get("mood", "auto"),
                                    design.get("scene", "meeting"), overrides)
        self.motif = design.get("motif", "none")  # 15.2 起默认不画；"bar" 恢复 15.1 的每页短横线。母题应从内容里长出来，不是固定位置的装饰
        self.sz = self.system["sizes"]
        self.fonts = design.get("fonts") or {"heading": "微软雅黑", "body": "微软雅黑"}
        self.stub = {"fonts": self.fonts}

    def pal(self, dark: bool) -> Dict[str, str]:
        return self.system["dark" if dark else "light"]

    def _style(self, el: Dict[str, Any], bg: str) -> Dict[str, Any]:
        return M.text_style(self.stub, {"background": bg}, el)

    def need(self, el: Dict[str, Any], w: float, bg: str) -> float:
        """框高 = 估算需求 × 1.05 + 0.02，向上取 0.05。折行按 98% 的宽度估：
        估宽留 2% 余量足以吸收中西文交界处的渲染差异；再收紧会把本来一行的标题估成两行，
        框被撑高，底边锚点与上方元素随之错位（14.0 的封面碰撞就是这样来的）。"""
        el = dict(el, x=0, y=0, w=w, h=1)
        st = self._style(el, bg)
        return math.ceil((M.fit_height_in(el, st, w * 0.98) * 1.05 + 0.02) * 20) / 20

    def fit_pt(self, kind: str, content: str, w: float, bg: str, hi: float, lo: float,
               runs_fn: Any = None, **kw: Any) -> int:
        """取不使任何一段折行的最大字号（hi→lo，步长 2）。字号随内容长度走，而不是一刀切。"""
        paras = max(1, len(str(content).split("\n")))
        size = int(hi)
        while size > lo:
            el: Dict[str, Any] = {"id": "x", "type": kind, "content": content, "pad_in": 0,
                                  "font_size": size, "x": 0, "y": 0, "w": w, "h": 1}
            el.update(kw)
            if runs_fn:
                el["runs"] = runs_fn(size)
            lines = M.fit_lines(el, self._style(el, bg), w * FIT_SAFETY)
            if sum(lines) <= paras:
                return size
            size -= 2
        return int(lo)

    def balanced_w(self, kind: str, content: str, w: float, size: float, bg: str, **kw: Any) -> float:
        """真折行时，在行数不变的前提下收窄到最窄宽度（text-wrap: balance），避免末行只剩一两个字。
        是否折行按 need() 同一把尺（98% 宽）判断，估成一行的不动；收窄按大字安全系数（86%）量，再留 8% 余量。"""
        def lines(width: float, safety: float) -> int:
            el: Dict[str, Any] = {"id": "x", "type": kind, "content": content, "pad_in": 0,
                                  "font_size": size, "x": 0, "y": 0, "w": width, "h": 1}
            el.update(kw)
            return sum(M.fit_lines(el, self._style(el, bg), width * safety))

        if lines(w, 0.98) < 2:
            return w
        n = lines(w, FIT_SAFETY)
        lo, hi = w * 0.5, w
        for _ in range(12):
            mid = (lo + hi) / 2
            lo, hi = (mid, hi) if lines(mid, FIT_SAFETY) > n else (lo, mid)
        return round(min(hi * 1.08, w), 3)

    def widow_w(self, kind: str, content: str, w: float, size: float, bg: str, max_w: Optional[float] = None, **kw: Any) -> float:
        """CJK 孤字防护：单段文字折行后末行只剩约一个字时，先在 max_w 内加宽，加宽不行再收窄，返回不产生孤字的宽度。
        估算按 em 宽度：末行余量 < 3 em（约一个词或两个汉字）即视为孤字。估不出（多段、含 runs）时原样返回，由 verify 的 V-ORPHAN 渲染兜底。"""
        if "\n" in content:
            return w
        el: Dict[str, Any] = {"id": "x", "type": kind, "content": content, "pad_in": 0, "font_size": size,
                              "x": 0, "y": 0, "w": w, "h": 1}
        el.update(kw)
        st = self._style(el, bg)
        em = M.para_width_em(content, [], size)
        inset = st["pad"][0]

        def last_em(width: float) -> Tuple[int, float]:
            cap = max((width * FIT_SAFETY - 2 * inset - M.list_indent(el, size)) / (size / 72.0), 0.5)
            n = max(1, math.ceil(em / cap - 1e-9))
            return n, em - (n - 1) * cap

        n0, tail0 = last_em(w)
        if n0 < 2 or tail0 >= 3.0:
            return w
        hi = max_w or w
        cand = [w + (hi - w) * k / 10 for k in range(1, 11)] if hi > w else []
        cand += [w * (1 - 0.04 * k) for k in range(1, 11)]
        for width in cand:
            n, tail = last_em(width)
            if n <= n0 and (n == 1 or tail >= 3.0):
                return round(width, 3)
        return w

    def image_tone(self, image: Dict[str, Any], box: Tuple[float, float, float, float]) -> Optional[bool]:
        """封面文字区下方图像的明暗：True＝偏暗（用浅墨），False＝偏亮（用深墨），读不了图返回 None。
        box = 文字区（英寸）；按 cover 裁切还原到原图坐标，只取平均相对亮度。"""
        if not self.asset_root:
            return None
        path = M.resolve_asset(self.asset_root, image.get("id", ""))
        if not path:
            return None
        try:
            from PIL import Image
            with Image.open(path) as im:
                im = im.convert("L")
                pw, ph = im.size
                l, r, t, b = M.crop_geometry(pw, ph, W, H, image.get("focus"))
                vx0, vy0, vx1, vy1 = pw * l, ph * t, pw * (1 - r), ph * (1 - b)
                x, y, w, h = box
                crop = im.crop((int(vx0 + (vx1 - vx0) * x / W), int(vy0 + (vy1 - vy0) * y / H),
                                int(vx0 + (vx1 - vx0) * (x + w) / W), int(vy0 + (vy1 - vy0) * (y + h) / H)))
                crop = crop.resize((24, 24))
                mean = sum(crop.tobytes()) / (24 * 24 * 255.0)
        except Exception:
            return None
        return mean < 0.5

    def text(self, eid: str, kind: str, x: float, y: float, w: float, content: str, bg: str,
             h: Optional[float] = None, **kw: Any) -> Dict[str, Any]:
        el: Dict[str, Any] = {"id": eid, "type": kind, "content": content, "pad_in": 0}
        el.update({k: v for k, v in kw.items() if v is not None})
        height = h if h is not None else self.need(el, w, bg)
        el.update(x=round(x, 3), y=round(y, 3), w=round(w, 3), h=round(height, 3))
        return el

    def rect(self, eid: str, x: float, y: float, w: float, h: float, color: str, **kw: Any) -> Dict[str, Any]:
        el = {"id": eid, "type": "shape", "shape_kind": kw.pop("shape_kind", "rect"), "x": round(x, 3),
              "y": round(y, 3), "w": round(w, 3), "h": round(h, 3), "color": color}
        el.update(kw)
        return el


# ---------------------------------------------------------------- 母题与页头

def _bar(c: Ctx, dark: bool, x: float = MX, y: float = MY) -> Dict[str, Any]:
    """跨页母题：每页同一位置的一道强调色短横线。它是整稿的“锚”，缩略总览里靠它认出同一套。"""
    return c.rect("bar", x, y, 0.6, 0.07, c.pal(dark)["accent"])


def _head(c: Ctx, p: Dict[str, Any], dark: bool, span: int = 10, start: int = 1,
          size: Optional[int] = None, box: Optional[Tuple[float, float]] = None) -> List[Dict[str, Any]]:
    """母题横线 + 可选眉题（eyebrow）+ 页题。页题是整页的论断。"""
    pal = c.pal(dark)
    x, w = box or col(start, span)
    bar_on = c.motif != "none" and p.get("motif") is not False
    els = [_bar(c, dark, x)]
    off = 0.85 if bar_on else 0.0  # 无母题横线时眉题与页题同左轴，不留下横线的空位
    if p.get("eyebrow"):
        els.append(c.text("eyebrow", "caption", x + off, MY - 0.11, w - off, p["eyebrow"], pal["bg"],
                          font_size=c.sz["caption"], color=pal["muted"], letter_spacing_pt=1.2, h=0.3))
    tsize = size or c.sz["title"]
    tw = c.balanced_w("title", p["title"], w, tsize, pal["bg"], line_spacing=1.12)
    els.append(c.text("t", "title", x, TOP, tw, p["title"], pal["bg"], font_size=tsize,
                      color=pal["ink"], line_spacing=1.12))
    return els


def _note(c: Ctx, p: Dict[str, Any], dark: bool) -> List[Dict[str, Any]]:
    if not p.get("note"):
        return []
    pal = c.pal(dark)
    x, w = col(1, 9)
    return [c.rect("foot_rule", MX, H - 0.98, W - 2 * MX, 0.012, pal["rule"]),
            c.text("note", "caption", x, H - 0.85, w, p["note"], pal["bg"], font_size=c.sz["caption"],
                   color=pal["muted"], h=0.3)]


def _image_el(spec: Dict[str, Any], x: float, y: float, w: float, h: float, carrier: bool = False) -> Dict[str, Any]:
    el = {"id": "img", "type": "image", "image_id": spec["id"], "x": round(x, 3), "y": round(y, 3),
          "w": round(w, 3), "h": round(h, 3),
          "image_role": spec.get("role", "background"), "info_duty": spec.get("info_duty", "atmosphere"),
          "asset_reason": spec.get("reason", "氛围图，承担景深与情绪"),
          "crop_focus": spec.get("focus", {"x": 0.5, "y": 0.5})}
    if carrier or el["image_role"] in ("background", "texture"):
        el["overlap_allowed"] = True  # 只有承载层（整幅图）可以被压；其余角色的图不与文字相交
    if spec.get("duotone"):
        el["duotone"] = spec["duotone"]
    return el


# ---------------------------------------------------------------- 版式配方

def r_cover(c: Ctx, p: Dict[str, Any]):
    """标题底边锚在 64% 水平线（valign 底对齐，框高估不准也不会漂），上方整块留给安静区。
    字号按标题长度取“不折行的最大字号”，短标题就大，长标题自动收。"""
    span = 8 if p.get("image") else 10
    x, w = col(1, span)
    if p.get("dark") is not None:
        dark = bool(p["dark"])
    else:   # 没写就看文字区下的真实像素：亮图配深墨，暗图配浅墨；无图或读不了图维持深底
        tone = c.image_tone(p["image"], (x, H * 0.25, w, H * 0.5)) if p.get("image") else None
        dark = True if tone is None else tone
    pal = c.pal(dark)
    full_w = w   # 标题收窄只为均衡折行；副题、元信息仍用整栏宽，否则会被挤出孤字
    els: List[Dict[str, Any]] = []
    if p.get("image"):
        els.append(_image_el(p["image"], 0, 0, W, H, carrier=True))
    size = c.fit_pt("title", p["title"], w, pal["bg"], c.sz["hero"] * 1.3, c.sz["hero"] * 0.8,
                    line_spacing=1.08, letter_spacing_pt=-0.5)
    w = c.balanced_w("title", p["title"], w, size, pal["bg"], line_spacing=1.08, letter_spacing_pt=-0.5)
    t = c.text("t", "title", x, 0, w, p["title"], pal["bg"], font_size=size, color=pal["ink"],
               line_spacing=1.08, letter_spacing_pt=-0.5, valign="bottom")
    anchor = H * 0.64
    t["y"] = round(anchor - t["h"], 3)
    els.append(_bar(c, dark, x))
    if p.get("kicker"):
        off = 0.85 if (c.motif != "none" and p.get("motif") is not False) else 0.0   # 无横线时与标题同左轴
        els.append(c.text("kicker", "caption", x + off, MY - 0.11, w - off, p["kicker"], pal["bg"],
                          font_size=c.sz["caption"], color=pal["ink2"] if p.get("image") else pal["accent"], letter_spacing_pt=1.5, bold=True, h=0.3))
    els.append(t)
    if p.get("sub"):
        els.append(c.text("sub", "caption", x, anchor + 0.3, full_w, p["sub"], pal["bg"],
                          font_size=c.sz["body"], color=pal["ink2"]))
    if p.get("meta"):
        els.append(c.text("meta", "caption", x, H - 0.85, full_w, p["meta"], pal["bg"], font_size=c.sz["caption"],
                          color=pal["muted"], h=0.3))
    return els, dark


def r_statement(c: Ctx, p: Dict[str, Any]):
    """一句话占满视野：竖线是页内唯一装饰，字号由句长决定。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    x, w = col(1, 10)
    tx = x + 0.45
    size = c.fit_pt("title", p["title"], w - 0.45, pal["bg"], c.sz["hero"] * 1.35, c.sz["hero"] * 0.8,
                    line_spacing=1.1)
    w = c.balanced_w("title", p["title"], w - 0.45, size, pal["bg"], line_spacing=1.1) + 0.45
    t = c.text("t", "title", tx, 0, w - 0.45, p["title"], pal["bg"], font_size=size, color=pal["ink"],
               line_spacing=1.1)
    sub_h = 0.0
    sub = None
    if p.get("sub"):
        sub = c.text("sub", "caption", tx, 0, w - 0.45, p["sub"], pal["bg"], font_size=c.sz["body"],
                     color=pal["muted"])
        sub_h = sub["h"] + 0.3
    t["y"] = round((H - (t["h"] + sub_h)) / 2 - 0.15, 3)
    els = [c.rect("vrule", x, t["y"] + 0.08, 0.07, t["h"] + sub_h - 0.1, pal["accent"]), t]
    if sub:
        sub["y"] = round(t["y"] + t["h"] + 0.3, 3)
        els.append(sub)
    return els, dark


def r_section(c: Ctx, p: Dict[str, Any]):
    """章节页：巨型序号（强调色）压在暗场里，标题底对齐。整稿的“停顿”与换气。"""
    dark = bool(p.get("dark", True))
    pal = c.pal(dark)
    x, w = col(1, 9)
    size = c.fit_pt("title", p["title"], w, pal["bg"], c.sz["title"] + 14, c.sz["title"] - 4, line_spacing=1.12)
    t = c.text("t", "title", x, 0, w, p["title"], pal["bg"], font_size=size, color=pal["ink"],
               line_spacing=1.12, valign="bottom")
    t["y"] = round(H - 1.0 - t["h"], 3)
    num = c.text("n", "number", x, MY - 0.15, w, str(p.get("number", "01")), pal["bg"],
                 font_size=round(c.sz["number"] * 1.7), color=pal["accent"], line_spacing=1.0, letter_spacing_pt=-4)
    return [num, t], dark


def r_big_number(c: Ctx, p: Dict[str, Any]):
    """数字是主角：字号取“放得进 7 栏的最大值”；正文只写结论与限定，被一条细线隔在右栏。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    els = _head(c, p, dark)
    unit = p.get("unit", "")
    num = str(p["number"])
    scale = float(p.get("unit_scale", 0.26))
    x, w = col(1, 7)

    def runs_fn(size: float) -> List[Dict[str, Any]]:
        return [{"text": num}, {"text": unit, "font_size": round(size * scale)}]

    size = c.fit_pt("number", num + unit, w, pal["bg"], c.sz["number"] * 2.0, c.sz["number"] * 0.8,
                    runs_fn=runs_fn if unit else None, line_spacing=1.0, letter_spacing_pt=-2)
    n = c.text("n", "number", x, 0, w, num + unit, pal["bg"], font_size=size, color=pal["accent"],
               line_spacing=1.0, letter_spacing_pt=-2, runs=runs_fn(size) if unit else None)
    n["y"] = round(H * 0.5 - n["h"] * 0.5 + 0.35, 3)
    bx, bw = col(8, 5)
    bw = c.widow_w("body", p["body"], bw, c.sz["body"] + 2, pal["bg"], max_w=bw * 1.12)
    body = c.text("b", "body", bx, 0, bw, p["body"], pal["bg"], font_size=c.sz["body"] + 2, color=pal["ink2"])
    body["y"] = round(n["y"] + n["h"] * 0.5 - body["h"] * 0.5, 3)
    rule_x = bx - 0.35
    top, bottom = min(body["y"], n["y"] + 0.3) - 0.1, max(body["y"] + body["h"], n["y"] + n["h"] - 0.3) + 0.1
    els += [n, c.rect("vrule", rule_x, top, 0.012, bottom - top, pal["rule"]), body]
    return els + _note(c, p, dark), dark


def r_chart(c: Ctx, p: Dict[str, Any]):
    """图承担读数；`highlight`（类别序号）让单系列柱/条只有一根强调色，其余退隐。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    els = _head(c, p, dark)
    top = els[-1]["y"] + els[-1]["h"] + 0.3
    bottom = FOOT - 0.1
    aside = p.get("aside")
    x, w = col(1, 8 if aside else 12)
    chart = dict(p["chart"])
    chart.setdefault("palette", c.system["data_dark" if dark else "data"])
    d = chart.get("data")
    if chart.get("kind") in ("bar", "pie", "donut") and isinstance(d, dict) and len(d.get("categories") or []) <= 8 \
            and len(d.get("series") or []) <= 3:
        chart.setdefault("data_labels", True)  # 类别少：直标读数，读者不必回到坐标轴估值（可显式设 false 关闭）
    fs = p.get("focus_series")
    if isinstance(fs, int) and not isinstance(fs, bool) and chart.get("kind") in ("bar", "line", "area") \
            and isinstance(d, dict) and 0 <= fs < len(d.get("series") or []) and "color_map" not in chart:
        # 强调色只指向结论：被论证的那条系列用强调色，其余退隐为灰（直标读数，不靠图例分辨）
        names = [str(sr.get("name")) for sr in d["series"]]
        chart["color_map"] = {n: (pal["accent"] if i == fs else pal["quiet"]) for i, n in enumerate(names)}
    if p.get("highlight") is not None and chart.get("kind") == "bar":
        chart["highlight_category"] = int(p["highlight"])
        chart.setdefault("legend", False)
    els.append({"id": "ch", "type": "chart", "x": round(x, 3), "y": round(top, 3), "w": round(w, 3),
                "h": round(bottom - top, 3), "chart": chart})
    if aside:
        ax, aw = col(10, 3)
        a = c.text("aside", "body", ax, top + 0.35, aw, aside, pal["bg"], font_size=c.sz["body"] + 2,
                   color=pal["ink"])
        els.append(c.rect("vrule", ax - 0.3, a["y"] - 0.05, 0.012, a["h"] + 0.1, pal["accent"]))
        els.append(a)
    return els + _note(c, p, dark), dark


def _compare_scaled_els(c: Ctx, p: Dict[str, Any], pal: Dict[str, str], keys, starts, spans, sizes, big_i: int,
                        big_pt: int):
    els: List[Dict[str, Any]] = []
    labels, nums, bodies = [], {}, []
    for i, key in enumerate(keys):
        side = p[key]
        x, w = col(starts[i], spans[i])
        accent = bool(i == big_i)
        labels.append(c.text(f"{key}_label", "caption", x, 0, w, side["label"], pal["bg"],
                             font_size=c.sz["caption"] + 1, color=pal["muted"], bold=True, letter_spacing_pt=1))
        nums[i] = c.text(f"{key}_n", "number", x, 0, w, side["number"], pal["bg"], font_size=sizes[i],
                         color=pal["accent"] if accent else pal["quiet"], line_spacing=1.0,
                         letter_spacing_pt=-4 if sizes[i] > 100 else -2, valign="bottom")
        bodies.append((x, w, side))
    label_h = max(lb["h"] for lb in labels)
    top_y = label_h + 0.15
    base_y = top_y + nums[big_i]["h"]
    for i, key in enumerate(keys):
        # 同一条基线：valign=bottom 时基线在框底之上约 BASELINE_EM（LibreOffice 渲染实测标定），按字号补偿，小字不会浮在大字基线之上
        nums[i]["y"] = round(base_y + BASELINE_EM * (sizes[i] - big_pt) / 72.0 - nums[i]["h"], 3)
        labels[i]["y"] = 0.0
        x, w, side = bodies[i]
        b = c.text(f"{key}_b", "body", x, base_y + 0.25, c.widow_w("body", side["body"], w, c.sz["body"], pal["bg"],
                   max_w=w * 1.1), side["body"], pal["bg"], font_size=c.sz["body"], color=pal["ink2"])
        els += [labels[i], nums[i], b]
    return els, base_y


def _compare_scaled(c: Ctx, p: Dict[str, Any], head: List[Dict[str, Any]], vals: List[float], dark: bool):
    """尺度隐喻（`scale: "value"`）：数字字高与 `value` 成正比、落在同一条基线上，落差本身就是论点。
    较大者占 8 栏、较小者占 4 栏；较小者字号不低于 `SCALE_FLOOR_PT`（再小就读不清，不再成比例）。
    这是 design.md「尺度隐喻」生成器的配方化：压住与被压住由数据的真实比例决定，而不是由作者目测。"""
    pal = c.pal(dark)
    title = head[-1]
    big_i = 0 if vals[0] >= vals[1] else 1
    spans = {big_i: 8, 1 - big_i: 4}
    starts = {0: 1, 1: 1 + spans[0]}
    keys = ("left", "right")
    bw = col(starts[big_i], spans[big_i])[1]
    big_pt = c.fit_pt("number", str(p[keys[big_i]]["number"]), bw, pal["bg"], c.sz["number"] * 2.0,
                      c.sz["number"] * 0.6, line_spacing=1.0, letter_spacing_pt=-4)
    ratio = vals[1 - big_i] / vals[big_i]
    small_w = col(starts[1 - big_i], spans[1 - big_i])[1]
    while True:
        small_pt = max(SCALE_FLOOR_PT, int(round(big_pt * ratio / 2.0)) * 2)
        # 较小者必须放得进自己的 4 栏：放不下（两数接近时按比例的字号比窄栏还宽）就整体等比收小，
        # 否则它会折行、框高暴涨，与标签和标题实质重叠（E-OVERLAP）。
        if small_pt <= c.fit_pt("number", str(p[keys[1 - big_i]]["number"]), small_w, pal["bg"], small_pt, SCALE_FLOOR_PT,
                                 line_spacing=1.0, letter_spacing_pt=-4 if small_pt > 100 else -2) or big_pt <= SCALE_FLOOR_PT * 2:
            break
        big_pt -= 2
    free_top, free_bottom = title["y"] + title["h"] + 0.2, FOOT
    while True:
        sizes = {big_i: big_pt, 1 - big_i: max(SCALE_FLOOR_PT, int(round(big_pt * ratio / 2.0)) * 2)}
        els, base_y = _compare_scaled_els(c, p, pal, keys, starts, spans, sizes, big_i, big_pt)
        block = max(e["y"] + e["h"] for e in els)
        # 整块（标签+数字+正文）必须放进标题与页脚之间：放不下就整体等比收小，而不是越界压到脚注（E-OVERLAP）。
        if block <= free_bottom - free_top or big_pt <= SCALE_FLOOR_PT * 2:
            break
        big_pt -= 4
    shift = free_top + max((free_bottom - free_top - block) / 2, 0.0)
    for el in els:
        el["y"] = round(el["y"] + shift, 3)
    rule = c.rect("base_rule", MX, round(base_y + shift + 0.1, 3), W - 2 * MX, 0.012, pal["rule"])
    return head + els + [rule] + _note(c, p, dark), dark


def r_compare(c: Ctx, p: Dict[str, Any]):
    """两个数字同基线同尺度；若给了 `value`，下方再按真实比例（自零起）画两根条，落差本身成为图形。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    head = _head(c, p, dark)
    title = head[-1]
    els: List[Dict[str, Any]] = []
    block = 0.0
    vals = [p["left"].get("value"), p["right"].get("value")]
    use_bars = all(isinstance(v, (int, float)) and v > 0 for v in vals)
    if use_bars and p.get("scale") == "value":
        return _compare_scaled(c, p, head, vals, dark)
    vmax = max(vals) if use_bars else 1
    cw = col(1, 5)[1]
    nsize = min(c.fit_pt("number", str(p[k]["number"]), cw, pal["bg"], c.sz["number"] * 1.2, c.sz["number"] * 0.5,
                         line_spacing=1.0, letter_spacing_pt=-2) for k in ("left", "right"))  # 两数同字号才可比
    for key, start, accent, v in (("left", 1, False, vals[0]), ("right", 7, True, vals[1])):
        side = p[key]
        x, w = col(start, 5)
        label = c.text(f"{key}_label", "caption", x, 0, w, side["label"], pal["bg"], font_size=c.sz["caption"] + 1,
                       color=pal["muted"], bold=True, letter_spacing_pt=1)
        n = c.text(f"{key}_n", "number", x, label["h"] + 0.15, w, side["number"], pal["bg"], font_size=nsize,
                   color=pal["accent"] if accent else pal["quiet"], line_spacing=1.0, letter_spacing_pt=-2)
        b = c.text(f"{key}_b", "body", x, n["y"] + n["h"] + 0.1, c.widow_w("body", side["body"], w, c.sz["body"], pal["bg"], max_w=w * 1.1), side["body"], pal["bg"], font_size=c.sz["body"],
                   color=pal["ink2"])
        els += [label, n, b]
        end = b["y"] + b["h"]
        if use_bars:
            bar_y = end + 0.3
            els.append(c.rect(f"{key}_bar", x, bar_y, w * v / vmax, 0.14, pal["accent"] if accent else pal["rule"]))
            end = bar_y + 0.14
        block = max(block, end)
    free_top, free_bottom = title["y"] + title["h"] + 0.2, FOOT
    shift = free_top + max((free_bottom - free_top - block) / 2, 0.0)
    for el in els:
        el["y"] = round(el["y"] + shift, 3)
    mx, _ = col(6, 1)
    rule = c.rect("vrule", mx + COL_W / 2, shift + 0.1, 0.012, block - 0.1, pal["rule"])
    return head + els + [rule] + _note(c, p, dark), dark


def r_list(c: Ctx, p: Dict[str, Any]):
    """标题与条目在同一条顶线上起跑，中间一条竖线把两栏钉住；条目取大一档，撑得起半页。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    els = _head(c, p, dark, span=5)
    lx, lw = col(7, 6)
    longest = max(p["items"], key=len)   # 条目字号随最长一条取不折行的最大值：条目少时撑得起半页，而不是缩在角落
    size = c.fit_pt("body", longest, lw, pal["bg"], c.sz["body"] + 12, c.sz["body"] + 4, line_spacing=1.2,
                    list=p.get("list", "number"))   # 悬挂缩进随字号变，交给 fit_lines 按每个候选字号算
    lw = c.widow_w("body", longest, lw, size, pal["bg"], line_spacing=1.2, list=p.get("list", "number")) if len(longest) > 6 else lw
    lst = c.text("items", "body", lx, TOP + 0.05, lw, "\n".join(p["items"]), pal["bg"],
                 font_size=size, color=pal["ink"], list=p.get("list", "number"), space_after_pt=round(size * 1.1))
    els.append(lst)
    els.append(c.rect("vrule", col(6, 1)[0] + COL_W / 2, TOP + 0.1, 0.012, max(lst["h"], 2.2), pal["rule"]))
    return els + _note(c, p, dark), dark


def r_table(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", False))
    els = _head(c, p, dark)
    top = els[-1]["y"] + els[-1]["h"] + 0.4
    x, w = col(1, 12)
    rows = len(p["table"]["rows"]) + 1
    els.append({"id": "tb", "type": "chart", "x": round(x, 3), "y": round(top, 3), "w": round(w, 3),
                "h": round(min(FOOT - 0.1 - top, rows * 0.8), 3), "chart": {"kind": "table", "data": {"headers": p["table"]["headers"], "rows": p["table"]["rows"]},
                          **{k: p["table"][k] for k in ("column_align", "column_weights") if k in p["table"]}}})
    return els + _note(c, p, dark), dark


def r_split_image(c: Ctx, p: Dict[str, Any]):
    """图占一侧（`side` left|right，`ratio` 图宽占页宽 0.3–0.65，缺省 0.5）；文字在另一侧：标题压顶、正文压底。
    标题与正文必须在文字栏内留出 ≥0.4in 的中段：放不下就先收标题字号，而不是让两个文本框相交。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    side = p.get("side", "right")
    ratio = min(max(float(p.get("ratio", 0.5)), 0.3), 0.65)
    iw = W * ratio
    img_x = W - iw if side == "right" else 0.0
    x0, w0 = (MX, W - iw - MX - 0.55) if side == "right" else (iw + 0.55, W - iw - MX - 0.55)
    els = [_image_el(dict(p["image"], role=p["image"].get("role", "subject")), img_x, 0, iw, H)]
    foot = H - 1.15 if p.get("note") else H - MY
    if side == "left" and not p.get("note"):
        # 文字在右栏时正文右缘越过页码框：底边止于页码保留框之上，而不是压进去
        pn = M.LIMITS["page_number"]
        foot = min(foot, H - pn["margin_in"]["y"] - pn["box_in"]["h"] - 0.12)
    body = None
    if p.get("body"):
        body = c.text("b", "body", x0, 0, w0, p["body"], pal["bg"], font_size=c.sz["body"] + 2, color=pal["ink2"], valign="bottom")
        body["y"] = round(foot - body["h"], 3)
    size = c.sz["title"]
    while True:
        head = _head(c, p, dark, size=size, box=(x0, w0))
        t = head[-1]
        if body is None or body["y"] - (t["y"] + t["h"]) >= 0.4 or size <= 30:
            break
        size -= 4
    els += head
    if body is not None:
        els.append(body)
    if p.get("note"):
        nw = min(w0, PN_LEFT - 0.14 - x0) if side == "left" else w0   # 文字在右侧时页码也在右下角：note 止于页码框之前
        els.append(c.text("note", "caption", x0, H - 0.85, nw, p["note"], pal["bg"], font_size=c.sz["caption"],
                          color=pal["muted"]))
    return els, dark


def r_steps(c: Ctx, p: Dict[str, Any]):
    """顺序与阶段（2–5 步）。只表达先后，不画箭头：箭头会暗示未经证实的因果或时长。
    各步等分整条版心（不再只占 12 栏中的 10 栏），轴线与步骤同宽；步骤块在页题与页脚之间垂直居中；
    说明文字做孤字防护（末行不只剩一个字）。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    head = _head(c, p, dark)
    title = head[-1]
    items = p["steps"]
    n = len(items)
    pitch = (W - 2 * MX) / n
    gap = 0.35
    cw = pitch - gap
    has_when = any(it.get("when") for it in items)
    top_slot = 0.55 if has_when else 1.25   # 无 when 时，轴线上方放大号序号：给各步一个尺度锚点，页面不只剩一排小字
    nsize = c.sz["number"] * 0.55
    node_d = 0.22
    els: List[Dict[str, Any]] = []
    block = 0.0
    for i, it in enumerate(items):
        x = MX + i * pitch
        y = 0.0
        y += top_slot
        node_y = y
        y += node_d + 0.38
        lsize = c.sz["body"] + {2: 18, 3: 16, 4: 12}.get(n, 8)   # 步数越少，单步越有分量
        label = c.text(f"label{i}", "text", x, y, c.widow_w("text", it["label"], cw, lsize, pal["bg"], bold=True, line_spacing=1.2),
                       it["label"], pal["bg"], font_size=lsize, color=pal["ink"], bold=True, line_spacing=1.2)
        y2 = label["y"] + label["h"] + 0.15
        step_els = [c.rect(f"node{i}", x, node_y, node_d, node_d, pal["accent"], shape_kind="oval")]
        if not has_when:
            step_els.append(c.text(f"num{i}", "number", x, 0.0, cw, f"{i + 1:02d}", pal["bg"], font_size=nsize,
                                   color=pal["quiet"], line_spacing=1.0, letter_spacing_pt=-2, h=top_slot - 0.15))
        if it.get("when"):
            step_els.append(c.text(f"when{i}", "caption", x, 0.0, cw, it["when"], pal["bg"], font_size=c.sz["caption"] + 1,
                                   color=pal["accent"], bold=True, letter_spacing_pt=1, h=0.3))
        step_els.append(label)
        end = label["y"] + label["h"]
        if it.get("body"):
            bw = c.widow_w("body", it["body"], cw, c.sz["body"], pal["bg"], max_w=pitch - 0.1)
            body = c.text(f"body{i}", "body", x, y2, bw, it["body"], pal["bg"], font_size=c.sz["body"], color=pal["ink2"])
            step_els.append(body)
            end = body["y"] + body["h"]
        els += step_els
        block = max(block, end)
    free_top, free_bottom = title["y"] + title["h"] + 0.3, FOOT
    shift = free_top + max((free_bottom - free_top - block) / 2, 0.0)
    for el in els:
        el["y"] = round(el["y"] + shift, 3)
    axis_y = shift + top_slot + node_d / 2 - 0.01
    axis = c.rect("axis", MX, axis_y, W - 2 * MX, 0.02, pal["rule"])
    return head + [axis] + els + _note(c, p, dark), dark


def r_close(c: Ctx, p: Dict[str, Any]):
    dark = bool(p.get("dark", True))
    pal = c.pal(dark)
    x, w = col(1, 9)
    size = c.fit_pt("title", p["title"], w, pal["bg"], c.sz["title"] + 14, c.sz["title"], line_spacing=1.1)
    tw = c.balanced_w("title", p["title"], w, size, pal["bg"], line_spacing=1.1)
    t = c.text("t", "title", x, H * 0.27, tw, p["title"], pal["bg"], font_size=size, color=pal["ink"],
               line_spacing=1.1)
    els = [_bar(c, dark, x, t["y"] - 0.45), t]
    if p.get("items"):
        iy = t["y"] + t["h"] + 0.7
        els.append(c.rect("rule", x, iy - 0.35, W - 2 * MX, 0.012, pal["rule"]))
        els.append(c.text("items", "body", x, iy, w, "\n".join(p["items"]), pal["bg"],
                          font_size=c.sz["body"] + 6, color=pal["ink2"], list=p.get("list", "number"),
                          space_after_pt=18))
        over = els[-1]["y"] + els[-1]["h"] - (H - 0.85)   # 标题折行或大字号档会把条目推出页底：整块上移，不压字号
        if over > 0:
            shift = min(over, max(0.0, els[0]["y"] - 0.5))
            for e in els:
                e["y"] = round(e["y"] - shift, 3)
    return els, dark



TEXT_KINDS = ("title", "body", "text", "number", "caption", "footer")
COLOR_ALIASES = ("color", "ink", "fill")   # outline 速记：最终 plan 只有 color；ink（文字墨色）、fill（形状填充）是别名


def _alias_conflict(el: Dict[str, Any]) -> Optional[str]:
    """color / ink / fill 指向同一个最终字段：同时写多个即冲突，拒绝而不静默覆盖。"""
    used = [k for k in COLOR_ALIASES if k in el]
    return f"同时写了 {' 与 '.join(used)}：三者最终都是 plan 的 color，只留一个" if len(used) > 1 else None


def _alias_errors(els: Any, where: str) -> List[str]:
    """递归（含 group.elements）收集颜色别名冲突。"""
    out: List[str] = []
    for j, e in enumerate(els if isinstance(els, list) else [], 1):
        if isinstance(e, dict):
            msg = _alias_conflict(e)
            if msg:
                out.append(f"{where}第 {j} 个元素{msg}")
            out += _alias_errors(e.get("elements"), f"{where}第 {j} 个元素组内")
    return out


def r_free(c: Ctx, p: Dict[str, Any]):
    """自由构图：作者逐个写元素，系统只负责把“设计语言”翻成几何。不是配方，是配方之外的出口。

    每个元素是普通 plan 元素，外加这些**可选**速记（写了才换算，其余键原样交给 schema）：
      cols:[起列, 跨列]       → x, w（12 栏网格）；要出血或错网格就直接写 x, w
      size: 档位名|pt          → 档位取 title/body/number/caption/hero；  fit:[最大pt, 最小pt] → 取不折行的最大字号
      color（别名 ink＝文字墨色、fill＝形状填充；三者只能写一个，最终都落为 plan 的 color）可写调色名
        （ink ink2 muted quiet accent rule bg bg_alt）→ 换成当前明暗场的色值；组内子元素同样生效
      bottom: y                → 以该 y 为文字底边（valign=bottom），不依赖框高估算
      y 缺省且给了 after:元素id → 紧接该元素下方，gap 缺省 0.2 in
      h 缺省（文字）            → 由 model.fit_height_in 估出
      rows:[起行, 跨行]        → y, h；right: x → 以该 x 为右边（右对齐锚点）
    页级：grid 网格（缺省 12 栏 6 行；整数 N、{cols,rows,margin,margin_y,gutter,gutter_y}、"none"）；
          dark 明暗场；bg 页底色（调色名或 #hex，如 accent 整页色场）。"""
    dark = bool(p.get("dark", False))
    pal = c.pal(dark)
    bg = pal.get(p.get("bg", ""), p.get("bg")) or pal["bg"]
    gcol, grow, _n, _r = make_grid(p.get("grid"))
    def tok(v: Any) -> Any:
        return pal[v] if isinstance(v, str) and v in pal else v

    def paint(el: Dict[str, Any]) -> Dict[str, Any]:
        """调色名在所有会写颜色的地方一致生效；ink / fill 规范化为 color 并删除（递归含组内子元素），最终 plan 不残留别名。"""
        msg = _alias_conflict(el)
        if msg:
            raise ValueError(f"元素 {el.get('id', '?')}{msg}")
        el = dict(el)
        alias = next((k for k in ("ink", "fill") if k in el), None)
        if alias:
            el["color"] = el.pop(alias)
        if "color" in el:
            el["color"] = tok(el["color"])
        if isinstance(el.get("runs"), list):
            el["runs"] = [dict(r, color=tok(r["color"])) if isinstance(r, dict) and "color" in r else r for r in el["runs"]]
        if isinstance(el.get("stroke"), dict) and "color" in el["stroke"]:
            el["stroke"] = dict(el["stroke"], color=tok(el["stroke"]["color"]))
        if isinstance(el.get("gradient"), dict):
            el["gradient"] = {k: (tok(v) if k in ("from", "to") else v) for k, v in el["gradient"].items()}
        if isinstance(el.get("elements"), list):
            el["elements"] = [paint(ch) if isinstance(ch, dict) else ch for ch in el["elements"]]
        return el

    out: List[Dict[str, Any]] = []
    by_id: Dict[str, Dict[str, Any]] = {}
    for i, spec in enumerate(p["els"], 1):
        el = {k: v for k, v in spec.items() if k not in ("cols", "rows", "size", "fit", "bottom", "right", "after", "gap")}
        el.setdefault("id", f"e{i}")
        kind = el.setdefault("type", "text")
        el = paint(el)
        if kind == "chart" and isinstance(el.get("chart"), dict):
            # 暗场上的图表用暗场数据色序列（与 chart 配方同一来源），作者可用 chart.palette 覆盖
            el["chart"] = dict(el["chart"])
            el["chart"].setdefault("palette", c.system["data_dark" if dark else "data"])
        if "cols" in spec:
            el["x"], el["w"] = (round(v, 3) for v in gcol(*spec["cols"]))
        if "rows" in spec:
            el["y"], el["h"] = (round(v, 3) for v in grow(*spec["rows"]))
        if kind in TEXT_KINDS:
            el.setdefault("pad_in", 0)
            if "size" in spec:
                sz = spec["size"]
                el["font_size"] = c.sz.get(sz, c.sz["body"]) if isinstance(sz, str) else sz
            if "fit" in spec:
                hi, lo = spec["fit"]
                el["font_size"] = c.fit_pt(kind, el["content"], el["w"], bg, hi, lo,
                                           line_spacing=el.get("line_spacing"),
                                           letter_spacing_pt=el.get("letter_spacing_pt"))
            if "h" not in el:
                el["h"] = c.need(dict(el, x=0, y=0), el["w"], bg)
            if "bottom" in spec:
                el["y"] = round(spec["bottom"] - el["h"], 3)
                el.setdefault("valign", "bottom")
        if "right" in spec and "w" in el:
            el["x"] = round(spec["right"] - el["w"], 3)
        if "y" not in el and "after" in spec:
            ref = by_id[spec["after"]]
            el["y"] = round(ref["y"] + ref["h"] + spec.get("gap", 0.2), 3)
        for k in ("x", "y", "w", "h"):
            if k in el:
                el[k] = round(el[k], 3)
        by_id[el["id"]] = el
        out.append(el)
    return out, dark, p.get("bg")


# recipe → (函数, 默认 role, 讲法, 必填字段)
RECIPES = {
    "cover": (r_cover, "open", "标题底边锚在 64% 线，上方整块安静区", ("title",)),
    "statement": (r_statement, "hold", "一句话占满视野，竖线是唯一装饰，让人停下来", ("title",)),
    "section": (r_section, "breathe", "巨型序号压暗场，换气与分章", ("title",)),
    "big_number": (r_big_number, "build", "关键数字压住页面，正文只写结论与限定", ("title", "number", "body")),
    "chart": (r_chart, "build", "图承担读数，旁注只写意义", ("title", "chart")),
    "compare": (r_compare, "turn", "两个数字同基线同尺度，落差就是论点", ("title", "left", "right")),
    "list": (r_list, "build", "标题与条目同一顶线，竖线钉住两栏", ("title", "items")),
    "table": (r_table, "build", "表格承担精确读数", ("title", "table")),
    "steps": (r_steps, "build", "只表达先后，不画箭头", ("title", "steps")),
    "split_image": (r_split_image, "build", "半幅图承载空间，文字留在另一侧", ("title", "image")),
    "close": (r_close, "close", "收在决定与下一步", ("title",)),
    "free": (r_free, "build", "自由构图：关系决定几何，系统只供网格、色值与字号档", ("title", "els")),
}
ROOT_REQUIRED = ("deck_id", "title", "audience", "goal", "core_claim", "brief", "pages")


def _span_ok(v: Any, n: int) -> bool:
    return (isinstance(v, list) and len(v) == 2 and all(isinstance(i, int) for i in v)
            and 1 <= v[0] <= n and 1 <= v[1] <= n - v[0] + 1)


def validate_outline(outline: Dict[str, Any]) -> List[str]:
    """把常见的写错变成一句人话，而不是 KeyError 堆栈。"""
    errs: List[str] = []
    if not isinstance(outline, dict):
        return ["outline 根必须是 JSON 对象"]
    for k in ROOT_REQUIRED:
        if k not in outline:
            errs.append(f"根字段缺少 {k}")
    pages = outline.get("pages")
    if not isinstance(pages, list) or not pages:
        errs.append("pages 必须是非空数组")
        return errs
    for i, p in enumerate(pages, 1):
        if not isinstance(p, dict):
            errs.append(f"第 {i} 页必须是对象")
            continue
        r = p.get("recipe")
        if r not in RECIPES:
            errs.append(f"第 {i} 页 recipe={r!r} 不存在；可选：{'、'.join(RECIPES)}")
            continue
        for k in RECIPES[r][3]:
            if p.get(k) in (None, "", []):
                errs.append(f"第 {i} 页（{r}）缺少 {k}")
        if r == "free":
            errs += _alias_errors(p.get("els"), f"第 {i} 页（free）")
            for j, e in enumerate(p.get("els") or [], 1):
                if not isinstance(e, dict):
                    errs.append(f"第 {i} 页（free）第 {j} 个元素必须是对象")
                elif e.get("type", "text") in TEXT_KINDS and not e.get("content"):
                    errs.append(f"第 {i} 页（free）第 {j} 个文字元素缺少 content")
                elif ("cols" in e or "rows" in e) and make_grid(p.get("grid"))[0] is None:
                    errs.append(f"第 {i} 页（free）第 {j} 个元素用了 cols/rows，但本页 grid 为 none")
                elif "cols" in e and not _span_ok(e["cols"], make_grid(p.get("grid"))[2]):
                    errs.append(f"第 {i} 页（free）第 {j} 个元素 cols 须是 [起列, 跨列]，且不超出本页 {make_grid(p.get('grid'))[2]} 栏")
                elif "rows" in e and not _span_ok(e["rows"], make_grid(p.get("grid"))[3]):
                    errs.append(f"第 {i} 页（free）第 {j} 个元素 rows 须是 [起行, 跨行]，且不超出本页 {make_grid(p.get('grid'))[3]} 行")
                elif "after" in e and e["after"] not in [x.get("id") for x in p["els"][:j - 1]]:
                    errs.append(f"第 {i} 页（free）第 {j} 个元素 after 必须指向它之前的元素 id")
        if r == "compare":
            for side in ("left", "right"):
                for k in ("label", "number", "body"):
                    if not isinstance(p.get(side), dict) or k not in p[side]:
                        errs.append(f"第 {i} 页（compare）{side}.{k} 缺失")
        if r == "steps":
            if not (2 <= len(p.get("steps") or []) <= 5):
                errs.append(f"第 {i} 页（steps）steps 需 2–5 项")
            for j, it in enumerate(p.get("steps") or [], 1):
                if not isinstance(it, dict) or "label" not in it:
                    errs.append(f"第 {i} 页（steps）第 {j} 项缺少 label")
    return errs


class ContentDropped(ValueError):
    pass


COMMON_KEYS = {"recipe", "id", "role", "title", "focus", "rationale", "claim", "concept", "beat", "evidence", "open_questions",
               "speaker_notes", "page_number_tone", "focus_element", "depth", "design_state", "composition", "motif", "dark"}
RECIPE_KEYS = {   # 每个配方真正消费的页字段（与 references/build.md 配方表一致）；写了表外的键＝会被静默忽略
    "cover": {"sub", "kicker", "meta", "image"}, "statement": {"sub"}, "section": {"number"},
    "big_number": {"number", "unit", "unit_scale", "body", "note", "eyebrow"},
    "chart": {"chart", "aside", "highlight", "focus_series", "note", "eyebrow"},
    "compare": {"left", "right", "scale", "note", "eyebrow"}, "list": {"items", "list", "note", "eyebrow"},
    "table": {"table", "note", "eyebrow"}, "steps": {"steps", "note", "eyebrow"},
    "split_image": {"image", "body", "side", "ratio", "note", "eyebrow"}, "close": {"items", "list"},
}


def unknown_keys(p: Dict[str, Any]) -> List[str]:
    allowed = COMMON_KEYS | RECIPE_KEYS.get(p.get("recipe"), set())
    return [] if p.get("recipe") == "free" else sorted(k for k in p if k not in allowed)


LANDING_KEYS = ("title", "sub", "kicker", "meta", "eyebrow", "body", "note", "aside")


def unlanded_text(p: Dict[str, Any], elements: List[Dict[str, Any]]) -> List[str]:
    """作者写了、却没有进入任何元素的文字字段。配方只消费自己认识的键，其余会被静默丢弃——
    含 note（口径、示意图声明）这类不该丢的限定语。逐字比对，不做模糊匹配。"""
    def squash(t: str) -> str:
        return "".join(str(t).split())
    landed = squash("".join(str(e.get("content", "")) + "".join(str(r.get("text", "")) for r in e.get("runs") or [])
                            for e in M.all_elements({"elements": elements})))
    return [k for k in LANDING_KEYS if isinstance(p.get(k), str) and p[k].strip() and squash(p[k]) not in landed]


def build_plan(outline: Dict[str, Any], asset_root: Optional[Path] = None) -> Dict[str, Any]:
    """outline → plan。scaffold 只做几何与色值的算术：作者没写的设计判断（focus、构图理由、节奏）
    一律留成 PENDING 并把页标成 draft，不拿配方说明或标题冒充判断；作者两者都写了才算 judged。"""
    ctx = Ctx(outline, asset_root)
    plan: Dict[str, Any] = {k: v for k, v in outline.items() if k not in ("design", "pages")}
    plan["schema_version"] = M.SCHEMA_VERSION
    plan["fonts"] = ctx.fonts
    plan["palette"] = ctx.system["data"]
    pages, rhythm = [], []
    for i, p in enumerate(outline["pages"], 1):
        fn, role, _grammar, _req = RECIPES[p["recipe"]]
        res = fn(ctx, p)
        elements, dark = res[0], res[1]
        if unknown_keys(p):
            raise ContentDropped(f"第 {i} 页（{p['recipe']}）的字段 {'、'.join(unknown_keys(p))} 不被该配方消费，会被静默忽略："
                                 "拼写有误、换配方，或改用 free")
        if p["recipe"] != "free":
            lost = unlanded_text(p, elements)
            if lost:
                raise ContentDropped(f"第 {i} 页（{p['recipe']}）的 {'、'.join(lost)} 没有进入版面（该配方不消费这个字段，会被静默丢弃）："
                                     "换一个消费它的配方，或改用 free 把它放进构图")
        if p["recipe"] != "free" and (ctx.motif == "none" or p.get("motif") is False):   # free 页的元素 id 归作者，不与母题保留 id 撞车
            elements = [e for e in elements if e.get("id") != "bar"]
        authored = bool(p.get("focus") and p.get("rationale"))
        page: Dict[str, Any] = {
            "id": p.get("id", f"p{i}"), "role": p.get("role", role),
            "focus": p.get("focus") or M.PENDING, "elements": elements,
            "design_state": p.get("design_state") or ("judged" if authored else "draft"),
            "composition": {**{"source": "author" if authored else "recipe", "recipe": p["recipe"]},
                            **(p.get("composition") if isinstance(p.get("composition"), dict) else {})},
            "background": (res[2] if len(res) > 2 and res[2] else None) and ctx.pal(dark).get(res[2], res[2])
            or ctx.pal(dark)["bg"],
        }
        if p["recipe"] == "compare" and p.get("scale") == "value":
            vs = [(p["left"].get("value"), "left"), (p["right"].get("value"), "right")]
            if all(isinstance(v, (int, float)) and v > 0 for v, _ in vs):
                hi, lo = max(v for v, _ in vs), min(v for v, _ in vs)
                page.setdefault("focus_element", max(vs)[1] + "_n")
                page["composition"] = {"temperament": "dramatic", "hero_ratio": round(min(max(hi / lo, 1.0), 3.0), 2),
                                       "focus_mode": "single", "strategy": "type_led", **page["composition"]}
        if p.get("claim"):
            page["claim"] = p["claim"]
        if p.get("rationale"):
            page["composition_rationale"] = p["rationale"]
        for key in ("concept", "evidence", "open_questions", "speaker_notes", "page_number_tone", "focus_element", "depth"):
            if p.get(key):
                page[key] = p[key]
        pages.append(page)
        rhythm.append(p.get("beat") or f"{page['role']}：{M.PENDING}")
    plan["pages"] = pages
    plan["rhythm"] = rhythm
    return plan


def _guard(rep: M.Report, inputs: Dict[str, Path], out: Path, force: bool) -> bool:
    """init / migrate 共用：与 compile / verify / revise 同一套路径守卫（model.guard_paths）。"""
    M.guard_paths(rep, inputs, [("输出文件", out, ".json")], force)
    M.show(rep.items)
    return not rep.blocked


def _emit(out: Path, data: Dict[str, Any], force: bool) -> bool:
    def check(b: bytes) -> Optional[str]:
        try:
            return None if isinstance(json.loads(b.decode("utf-8")), dict) else "根不是对象"
        except ValueError as exc:
            return str(exc)
    try:
        M.write_atomic(out, json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"), force, check)
    except FileExistsError:
        print(f"[HARD] E-EXIST {out} 在写入前被创建，未覆盖")
        return False
    except (OSError, ValueError) as exc:
        print(f"[HARD] E-RUNTIME 写出失败，未留下半成品：{exc}")
        return False
    return True


def schema_fields(name: str) -> str:
    """plan.schema.json 的紧凑字段表（按需查看，不必通读 40KB 的 schema）：`key*`＝必填，类型/枚举一行一项。"""
    defs = M.SCHEMA.get("$defs", {})
    node = M.SCHEMA if name == "root" else defs.get(name)
    if node is None:
        return f"未知节点 {name}；可选：root {' '.join(sorted(defs))}"

    def brief(s: Dict[str, Any], deep: bool = True) -> str:
        if "$ref" in s:
            return s["$ref"].rsplit("/", 1)[-1]
        if "enum" in s:
            return "|".join(map(str, s["enum"]))
        if "anyOf" in s:
            return "|".join(brief(x, False) for x in s["anyOf"])
        t = s.get("type", "any")
        if t == "array":
            return f"[{brief(s.get('items', {}), False)}]"
        if t == "object" and s.get("properties"):   # 嵌套对象带一层类型：写错数组/字符串是最常见的返工
            return "{" + ",".join(f"{k}:{brief(v, False)}" if deep else k for k, v in s["properties"].items()) + "}"
        lim = "".join(f" {k}={s[k]}" for k in ("minimum", "maximum", "exclusiveMinimum", "minItems", "maxItems") if k in s)
        return f"{t}{lim}"

    req = set(node.get("required", []))
    return "\n".join(f"{k}{'*' if k in req else ''}: {brief(v)}" for k, v in node.get("properties", {}).items())


def canvas_spec() -> Dict[str, Any]:
    """16:9 画布的坐标契约（只读）：同一份几何同时给英寸（plan 的单位）与 1280×720 设计像素，AI 做空间构成时直接取用。"""
    def both(v: float) -> Dict[str, float]:
        return {"in": round(v, 3), "px": round(M.px(v), 1)}
    cols = [{"col": i + 1, "x": both(MX + i * (COL_W + GUT)), "w": both(COL_W)} for i in range(COLS)]
    return {
        "canvas": {"px": list(M.CANVAS_PX), "in": [round(W, 4), round(H, 4)], "emu": list(M.CANVAS_EMU), "px_per_in": M.PX_PER_IN,
                   "note": "plan 坐标单位是英寸；px = in × 96；1 pt = 1.333 px"},
        "margins": {"x": both(MX), "y": both(MY)},
        "grid_12": {"gutter": both(GUT), "col_w": both(COL_W), "cols": cols},
        "content_top": both(TOP), "footer_top": both(FOOT),
        "page_number_left": both(PN_LEFT),
        "split_lines_x": {"thirds": [both(W / 3), both(2 * W / 3)], "five_seven": both(W * 5 / 12), "golden": [both(W * 0.382), both(W * 0.618)]},
        "split_lines_y": {"thirds": [both(H / 3), both(2 * H / 3)], "golden": [both(H * 0.382), both(H * 0.618)]},
    }


def main(argv: Optional[List[str]] = None) -> int:
    M.utf8_stdio()
    ap = argparse.ArgumentParser(description="设计系统与版式配方")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s1 = sub.add_parser("system", help="由种子色推导色彩系统与字号阶梯")
    s1.add_argument("--seed", required=True)
    s1.add_argument("--mood", default="auto", choices=["auto", "warm", "cool"])
    s1.add_argument("--scene", default="meeting", choices=list(SCENES))
    s2 = sub.add_parser("init", help="outline.json → plan.json 初稿（设计判断留 draft）")
    s2.add_argument("outline")
    s2.add_argument("out")
    s2.add_argument("--assets", help="素材根（封面按图像明暗取墨色时读取；缺省取 outline 同级 assets/，否则 outline 所在目录）")
    s2.add_argument("--force", action="store_true", help="允许替换已存在的输出（仍不得覆盖输入、不经符号链接）")
    s3 = sub.add_parser("migrate", help="旧版 plan.json → 当前 schema_version，不就地改写")
    s3.add_argument("plan")
    s3.add_argument("out")
    s3.add_argument("--force", action="store_true")
    sub.add_parser("canvas", help="16:9 画布坐标契约：1280×720 px 与英寸对照、12 栏网格、安全区、分割线（只读）")
    s4 = sub.add_parser("fields", help="plan.schema.json 的紧凑字段表（root/page/element/chart/…），代替通读 schema")
    s4.add_argument("node", nargs="?", default="root")
    a = ap.parse_args(argv)
    if a.cmd == "fields":
        print(schema_fields(a.node))
        return 0
    if a.cmd == "canvas":
        for key, value in canvas_spec().items():
            print(f"{key}: {json.dumps(value, ensure_ascii=False)}")
        return 0
    if a.cmd == "system":
        try:
            print(json.dumps(design_system(a.seed, a.mood, a.scene), ensure_ascii=False, indent=2))
        except (ValueError, TypeError) as exc:
            print(f"[HARD] E-OUTLINE seed：{exc}")
            return 2
        return 0
    src, out = Path(a.outline if a.cmd == "init" else a.plan), Path(a.out)
    rep = M.Report()
    if not _guard(rep, {"输入": src}, out, a.force):
        print("\n阻断：路径不安全，未写出文件。")
        return 2
    try:
        data = M.strict_json_loads(src.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        print(f"[HARD] E-OUTLINE 读不了 {src}：{exc}")
        return 2
    if a.cmd == "migrate":
        if not isinstance(data, dict):
            print("[HARD] E-SCHEMA plan 根必须是 JSON 对象")
            return 2
        try:
            new, notes = M.migrate_plan(data)
        except ValueError as exc:
            print(f"[HARD] E-SCHEMA {exc}")
            return 2
        if not _emit(out, new, a.force):
            return 3
        print("\n".join(f"  · {n}" for n in notes) or "  · 已是当前版本，无需变更")
        print(f"已写出 {out}。迁移只做可证的变更；请再跑 compile.py 校验。")
        return 0
    errs = validate_outline(data)
    if errs:
        for e in errs:
            print(f"[HARD] E-OUTLINE {e}")
        print(f"\n阻断：{len(errs)} 个 outline 问题，未写出文件。")
        return 2
    try:
        root = Path(a.assets) if a.assets else (src.parent / "assets" if (src.parent / "assets").is_dir() else src.parent)
        plan = build_plan(data, root)
    except ContentDropped as exc:
        print(f"[HARD] E-DROP {exc}")
        return 2
    except (ValueError, KeyError, TypeError, IndexError) as exc:
        print(f"[HARD] E-OUTLINE 无法生成 plan：{type(exc).__name__}: {exc}")
        return 2
    bad = M.schema_errors(plan)
    if bad:   # 在 init 就拦：不写出一份注定编译不过的 plan（路径相对生成后的 plan；字段表见 `fields`）
        for e in bad:
            print(f"[HARD] E-SCHEMA {e}")
        print(f"\n阻断：生成的 plan 有 {len(bad)} 个 schema 问题，未写出文件（用 scaffold.py fields <节点> 查字段与枚举）。")
        return 2
    if not _emit(out, plan, a.force):
        return 3
    draft = [p["id"] for p in plan["pages"] if p["design_state"] == "draft"]
    print(f"已写出 {out}（{len(plan['pages'])} 页）。下一步：compile.py → verify.py → 看图，再手改文案与几何。")
    if draft:
        print(f"[PENDING] {len(draft)} 页设计判断未完成（focus / composition_rationale 待写）：{', '.join(draft)}；"
              f"草稿可继续编译，交付前用 compile.py --production 复核。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
