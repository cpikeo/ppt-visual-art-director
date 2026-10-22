#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""guard.py · 硬错误拦截（只回答「能不能交」）。

纪律：
  * 这里只拦确定性错误：几何越界、墨迹相撞、装不下、图表载荷非法、字段名写错、
    证据缺来源。输出只有 PASS / BLOCK，不评分、不重新设计。
  * 审美与构图平衡不在这里判：guard PASS 不等于设计通过（SKILL「四层架构与权威源」「证据层级」）。
  * 文本容量用 typography 的真实度量；没有字体度量时降级并把 degraded 写进报告。
"""
from __future__ import annotations

from pathlib import Path

from typography import FontBook, ink_box, orphan_issues, wrap

# 元素字段白名单（写错字段名会让渲染器静默丢弃，因此必须拦）
COMMON_FIELDS = {"id", "type", "role", "x", "y", "width", "height", "layout_source",
                 "z", "background", "note"}
TYPE_FIELDS = {
    "text": {"text", "size", "color", "bold", "italic", "align", "anchor", "line_height",
             "tracking", "max_lines", "wrap", "font", "space_before", "space_after", "family",
             "opacity", "uppercase", "text_anchor_v"},
    "shape": {"shape", "fill", "stroke", "stroke_width", "text", "text_size", "text_color",
              "text_bold", "text_align", "radius"},
    "chart": {"chart_kind", "categories", "series", "show_values", "show_legend", "takeaway", "derived",
              "value_size", "label_size", "highlight", "target", "target_label", "unit_suffix"},
    "image": {"source", "alt_text", "fit", "crop", "crop_loss", "overlay", "asset_id", "function", "reason",
              "safe_area", "text_color", "asset_sha256", "source_px", "format", "layer", "brief"},
}
REQUIRED_FIELDS = {
    "text": {"text", "size"},
    "shape": {"shape"},
    "chart": {"chart_kind", "categories", "series"},
    "image": {"source"},
}
CHART_KINDS = {"column", "bar", "line", "donut", "progress"}
# 定位类元素：页码/眉标这类文字不承载信息，重复判定要排除它们（案例 B s05：页码 05 与工序 05）
LOCATOR_ROLES = {"eyebrow"}
NUMERIC_KINDS = {"column", "bar", "line", "donut"}

FIX_HINTS = {
    "GEOMETRY_INVALID": "坐标必须是有限数值且非退化：线元素可以 width/height 为 0，其他元素不行。",
    "OUT_OF_CANVAS": "元素必须落在画布内：先改构图，不要靠缩放画布迁就。",
    "TEXT_SIZE_FLOOR": "字号有下限：内容多先删句，不靠缩字号解决。",
    "EMPTY_SLIDE": "空页要么删掉，要么补一页该有的主张与证据。",
    "EMPTY_DECK": "没有页面就没有交付物。",
    "FONT_FAMILIES": "字体家族 ≤2：一个中文族 + 一个拉丁族。",
    "ELEMENT_ID_DUPLICATE": "元素/页面 id 必须唯一，取证与对账都依赖它。",
    "DONUT_NOT_PERCENT": "环图表达占比：合计不是 100 时改用条形图。",
    "CONTENT_TRUNCATED": "内容放不进构图的列数上限：删条目或换构图，不能静默丢弃。",
    "EVIDENCE_NOT_LAID_OUT": "作者声明的量化证据没有出现在页面上：换构图或删数据，不能只在 plan 里有。",
    "SOURCE_ZONE_COLLISION": "来源区只放来源/口径文字，其他对象不得侵入。",
    "ASSET_MISSING": "素材文件不存在或未声明：缺图不能静默通过。"
                     "该页若已登记生产说明，用 `pvad.py media deck.json` 取出提示词去生产/委托拍摄；"
                     "不需要这张图就删掉图位。",
    "ASSET_FORMAT": "素材格式不在白名单：转成 png/jpg/webp 再入包。",
    "ASSET_RESOLUTION": "素材源像素不足：换更大的图，或缩小图位；不要放大糊图。",
    "ASSET_CROP": "cover 裁掉太多内容：换构图/换图，或写明裁切理由。",
    "ASSET_SAFE_AREA": "文字压在图上但没有声明安全区：先定文字落点，再谈构图。",
    "ASSET_CONTRAST": "文字与图像实测对比度不足：改文字色、加遮罩层或换图位。",
    "ASSET_BUDGET": "单页素材超上限：图片不是靠数量建立氛围。",
    "ASSET_DECLARATION": "图的职责与页面判断不一致：先写清它替观众理解了什么。",
    "TEXT_OVERFLOW": "删内容或改短句；不要缩字号到下限以下。可用 --explain 看每行的度量。",
    "ELEMENT_OVERLAP": "先改构图（重排/删元素），不要靠缩小间距掩盖。",
    "TEXT_TOO_WIDE": "改短文本或加宽框；中文断行由作者语义决定。",
    "PROVENANCE_MISSING": "补齐 source/unit/period/basis，或把该页证据删掉。",
    "UNIT_INCONSISTENT": "同一指标跨页必须同单位：统一口径，或拆成两个指标名。",
    "CHART_INVALID": "图表形状必须来自数据：检查 categories 与 values 长度一致、数值有限。",
    "CHART_MIXED_UNITS": "一张图只画一套量纲：门店数（家）与占比（%）不能同图同轴；拆图或改成两个指标块。",
    "FIELD_UNKNOWN": "字段名必须与 schema 完全一致，否则渲染器静默丢弃。",
    "TOKEN_UNKNOWN": "颜色只接受 #RRGGBB 或 theme.colors 里的 token 名。",
    "ACCENT_OVER_BUDGET": "强调色只在语义落点使用：超过预算就说明处处强调。",
    "COMPOSITION_CONTRACT": "构图与数据形状不符：要么补数据，要么换构图。",
    "LAYOUT_DRIFT": "PPTX 回读与 spec 不一致：这是编译器缺陷，先修编译，不要改稿迁就。",
}


class _Report:
    def __init__(self, policy):
        self.policy = policy
        self.errors: list[dict] = []
        self.warnings: list[dict] = []
        self.stats = {"elements": 0, "text_elements": 0, "charts": 0, "images": 0, "slides": 0,
                      "measured": True}

    def error(self, code: str, detail: str, *, slide=None, element=None, fix: str = ""):
        self.errors.append({"code": code, "slide": slide, "element": element, "detail": detail,
                            "fix": fix or FIX_HINTS.get(code, "")})

    def warn(self, code: str, detail: str, *, slide=None, element=None):
        self.warnings.append({"code": code, "slide": slide, "element": element, "detail": detail})


def check(deck: dict, policy, *, mode: str = "draft") -> dict:
    """返回 {status, errors, warnings, fix_groups, stats, degraded}。"""
    fontbook = FontBook(policy)
    rep = _Report(policy)
    if fontbook.degraded:
        rep.stats["measured"] = False
        rep.warn("MEASUREMENT_DEGRADED",
                 "没有可用字体文件：文本容量是估算值，不得当作真实排版结论")
    _check_deck(deck, policy, rep, mode, fontbook)
    return _finish(rep, deck, mode)


# ──────────────────────────────────────────────────────────────────────────
def _check_deck(deck: dict, policy, rep: _Report, mode: str, fontbook: FontBook) -> None:
    cp = policy.req("canvas")
    cw, ch = float(cp["width"]), float(cp["height"])
    slides = deck.get("slides") or []
    rep.stats["slides"] = len(slides)
    if not slides:
        rep.error("EMPTY_DECK", "deck 没有任何幻灯片：空稿不能发布")
        return
    theme = deck.get("theme") or {}
    colors = (theme.get("colors") or {})
    tokens = set(policy.req("colors.tokens"))
    unknown = sorted(set(colors) - tokens)
    if unknown:
        rep.error("TOKEN_UNKNOWN", f"theme.colors 出现未定义的 token：{unknown}")
    for token in tokens:
        if token not in colors:
            rep.error("TOKEN_UNKNOWN", f"theme.colors 缺 token：{token}（缺失会在产物里静默回落）")
    fonts = theme.get("fonts") or {}
    families = {v for v in fonts.values() if v}
    if len(families) > 2:
        rep.error("FONT_FAMILIES", f"字体家族 {sorted(families)} 超过 2 个：统一的是表达逻辑，不是堆字体")

    ids = [s.get("id") for s in slides]
    if len(set(ids)) != len(ids):
        rep.error("ELEMENT_ID_DUPLICATE", f"slide id 重复：{ids}")

    for slide in slides:
        _check_slide(slide, deck, policy, rep, mode, fontbook, cw, ch, colors, {})
    _check_metric_units((deck.get("meta") or {}).get("metric_units") or {}, rep)


def _check_slide(slide, deck, policy, rep, mode, fontbook, cw, ch, colors, metric_units) -> None:
    sid = slide.get("id")
    elements = slide.get("elements") or []
    if not elements:
        rep.error("EMPTY_SLIDE", "这一页没有任何元素：空页要么删掉，要么补内容", slide=sid)
        return
    rep.stats["elements"] += len(elements)
    identity = slide.get("meta") or {}
    accent_limit = int(policy.req("colors.accent_budget_per_page"))
    accent_count = 0
    max_size = 0.0
    size_owners: list[tuple] = []
    seen_ids: set[str] = set()
    texts: set[str] = set()
    boxes: list[dict] = []
    source_elements: list[dict] = []
    numbers = 0

    for element in elements:
        eid = element.get("id")
        where = {"slide": sid, "element": eid}
        if not eid or eid in seen_ids:
            rep.error("ELEMENT_ID_DUPLICATE", f"元素 id 缺失或重复：{eid!r}", **where)
        seen_ids.add(eid)
        etype = element.get("type")
        if etype not in TYPE_FIELDS:
            rep.error("SCHEMA_FAIL", f"未知元素类型：{etype!r}", **where)
            continue
        rep.stats["text_elements"] += 1 if etype == "text" else 0
        _check_fields(element, etype, rep, where)
        _check_element_colors(element, colors, rep, where)
        if not _check_geometry(element, cw, ch, policy, rep, where):
            continue
        if etype == "chart":
            rep.stats["charts"] += 1
            _check_chart(element, policy, rep, where, identity)
        if etype == "image":
            rep.stats["images"] = rep.stats.get("images", 0) + 1
            _check_image(element, deck, policy, rep, where)
        if etype == "text":
            size = float(element.get("size") or 0)
            _check_text_size(element, size, policy, rep, where)
            if element.get("color") == "accent":
                accent_count += 1
            max_size, size_owners = _track_focus(size, eid, max_size, size_owners)
            text = str(element.get("text") or "")
            # 眉标是定位符（页序），不是内容：它和正文里的「05」相同不构成重复信息。
            if (str(element.get("role") or "") in LOCATOR_ROLES
                    or str(element.get("role") or "").startswith("legend_key")):
                pass                              # 定位符/色标不参与重复判定，也不算「数字证据」
            else:
                if text and text in texts:
                    rep.warn("DUPLICATE_TEXT", f"同一页重复出现相同文字：{text[:24]}", **where)
                texts.add(text)
                numbers += _count_numbers(text)
            if element.get("role") in ("source", "note") and identity.get("source"):
                source_elements.append(element)
        ink = ink_box(element, policy, fontbook)
        ink["id"] = eid
        ink["role"] = element.get("role")
        ink["element"] = element
        boxes.append(ink)
        if etype == "text":
            _check_text_fit(element, ink, policy, rep, where)
    _check_overlaps(boxes, policy, rep, sid)
    _check_focus(size_owners, rep, sid, {b["id"]: b["role"] for b in boxes})
    _check_media_contract(slide, deck, elements, boxes, policy, rep)
    _check_text_over_image(slide, deck, elements, boxes, policy, rep)
    _check_text_contrast(slide, elements, policy, rep, colors)
    _check_content_truncated(slide, rep, sid)
    _check_evidence_laid_out(slide, rep, sid)
    if accent_count > accent_limit:
        rep.error("ACCENT_OVER_BUDGET",
                  f"本页强调色元素 {accent_count} 个 > 预算 {accent_limit}：处处强调 = 没有强调",
                  slide=sid)
    if identity.get("has_numeric_evidence") and mode == "release":
        metrics = identity.get("metrics") or []
        unit_ok = bool(str(identity.get("unit") or "").strip()) or (
            bool(metrics) and all(str(m.get("unit") or "").strip() for m in metrics))
        missing = [k for k in policy.req("verification.release_requires")
                   if not (unit_ok if k == "unit" else str(identity.get(k) or "").strip())]
        if missing:
            rep.error("PROVENANCE_MISSING",
                      f"这是含数字的页，release 要求 {missing} 齐备：数字必须与来源/单位/期间/口径同源",
                      slide=sid)
        bare = [m.get("label") or m.get("value") for m in identity.get("metrics") or []
                if not str(m.get("unit") or "").strip()]
        if bare:
            rep.error("PROVENANCE_MISSING",
                      f"这些指标没有单位：{bare}。补单位（家/万元/%），不要用 deck 级单位兜底", slide=sid)
    if identity.get("source") and not any(b["role"] == "source" for b in boxes):
        rep.error("PROVENANCE_MISSING", "meta 声明了来源，但页面上没有来源元素", slide=sid)
    _check_density(numbers, identity, policy, rep, sid)


def _check_fields(element: dict, etype: str, rep: _Report, where: dict) -> None:
    allowed = COMMON_FIELDS | TYPE_FIELDS[etype]
    unknown = {k for k in element if k not in allowed and not k.startswith("_")}
    known = allowed | COMMON_FIELDS | set().union(*TYPE_FIELDS.values())
    for name in sorted(unknown):
        close = _closest(name, known)
        if close:
            rep.error("FIELD_UNKNOWN",
                      f"字段名 {name!r} 疑似写错，应为 {close!r}：渲染器不消费该名，等于静默丢弃",
                      **where)
        else:
            rep.warn("FIELD_UNDECLARED", f"未知字段 {name!r}（可能是作者注解，已忽略）", **where)
    for name in sorted(REQUIRED_FIELDS[etype]):
        if name not in element:
            rep.error("SCHEMA_FAIL", f"{etype} 元素缺必填字段 {name!r}", **where)


_COLOR_FIELDS = ("color", "fill", "stroke", "text_color", "background")


_IMAGE_COLOR_FIELDS = ("fill", "stroke")     # 图片没有描边/填充语义


def _check_element_colors(element: dict, colors: dict, rep: _Report, where: dict) -> None:
    """颜色只接受 #RRGGBB 或 theme 里定义过的 token。

    写错 token 不会报错，只会在产物里静默回落成默认色——所以必须拦。
    图片的 text_color 是语义提示（dark/light），不是颜色值，单独校验。
    """
    tokens = set(colors)
    if element.get("type") == "image":
        hint = str(element.get("text_color") or "")
        if hint and hint not in ("dark", "light"):
            rep.error("ASSET_DECLARATION",
                      f"text_color={hint!r} 只接受 dark / light（表示图上文字用深色还是浅色）", **where)
        fields = _IMAGE_COLOR_FIELDS
    else:
        fields = _COLOR_FIELDS
    for field in fields:
        value = element.get(field)
        if isinstance(value, dict):
            value = value.get("color") if value.get("type") != "none" else None
        if not isinstance(value, str) or not value:
            continue
        if value.startswith("#"):
            if len(value) != 7:
                rep.error("TOKEN_UNKNOWN", f"{field}={value!r} 不是 #RRGGBB", **where)
            continue
        if value not in tokens:
            rep.error("TOKEN_UNKNOWN",
                      f"{field}={value!r} 不是已定义的 token（{sorted(tokens)}）：写错会静默回落", **where)


def _closest(name: str, known) -> str | None:
    if name in known:
        return None
    lowered = name.lower().replace("-", "_")
    if lowered in known:
        return lowered
    # 前缀/后缀变形：font_size → size、maxlines → max_lines、textColor → text_color
    squashed = lowered.replace("_", "")
    for candidate in known:
        key = candidate.replace("_", "")
        if squashed == key:
            return candidate
        if len(key) >= 4 and (squashed.endswith(key) or squashed.startswith(key)) and len(squashed) <= len(key) + 8:
            return candidate
    best, score = None, 0.0
    import difflib
    for candidate in known:
        ratio = difflib.SequenceMatcher(None, lowered, candidate).ratio()
        if ratio > score:
            best, score = candidate, ratio
    return best if score >= 0.82 else None


def _check_geometry(element, cw, ch, policy, rep, where) -> bool:
    try:
        x, y = float(element["x"]), float(element["y"])
        w, h = float(element["width"]), float(element["height"])
    except (KeyError, TypeError, ValueError):
        rep.error("GEOMETRY_INVALID", f"几何字段缺失或非数值：{ {k: element.get(k) for k in ('x','y','width','height')} }",
                  **where)
        return False
    if not all(map(_finite, (x, y, w, h))):
        rep.error("GEOMETRY_INVALID", "几何字段必须是有限数值（NaN/Inf 会让 pptx 抛错）", **where)
        return False
    is_rule = element.get("type") == "shape" and element.get("shape") in ("line", "arrow")
    if w < 0 or h < 0 or (not is_rule and (w == 0 or h == 0)):
        rep.error("GEOMETRY_INVALID", f"退化几何：width={w} height={h}", **where)
        return False
    if x < 0 or y < 0 or x + w > cw + 0.5 or y + h > ch + 0.5:
        rep.error("OUT_OF_CANVAS", f"元素越出画布：({x:.1f},{y:.1f},{w:.1f},{h:.1f}) 画布 {cw}×{ch}",
                  **where)
        return False
    safe = float(policy.num("thresholds.safe_area_ratio"))
    inset_x, inset_y = cw * (1 - safe) / 2, ch * (1 - safe) / 2
    if not is_rule and (x < inset_x - 1 or y < inset_y - 1 or x + w > cw - inset_x + 1 or y + h > ch - inset_y + 1):
        rep.warn("SAFE_AREA", f"元素贴近画布边缘（安全区 {safe:.0%}）：投影与裁切时容易被吃掉", **where)
    return True


def _finite(value) -> bool:
    try:
        value = float(value)
    except (TypeError, ValueError):
        return False
    return value == value and value not in (float("inf"), float("-inf"))


def _check_text_size(element, size, policy, rep, where) -> None:
    scale = policy.req("typography.scale")
    key = min(scale, key=lambda k: abs(float(scale[k]) - size))
    floor = policy.floor(key)
    if size < floor - 0.01:
        rep.error("TEXT_SIZE_FLOOR",
                  f"字号 {size}px 低于 {key} 档下限 {floor}px：该删内容，不是缩字号", **where)


def _track_focus(size, eid, max_size, owners):
    if size > max_size + 0.01:
        return size, [(eid, size)]
    if abs(size - max_size) <= 0.01:
        return max_size, owners + [(eid, size)]
    return max_size, owners


_SYMMETRIC_ROLES = {"left_value", "right_value"}


def _check_focus(owners, rep, sid, roles=None) -> None:
    """同页只能有一个视觉主角——但一组对称的两端（比较的两侧）不算竞争。
    多指标并排的每个数字也各自是「主角」，它们是同一判断的几个面。"""
    roles = roles or {}
    families = {str(roles.get(eid, eid)).split(".")[0] for eid, _ in owners}
    if len(owners) > 1 and not (len(families) == 1 or families <= _SYMMETRIC_ROLES):
        rep.warn("FOCUS_COMPETITION",
                 f"同页有 {len(owners)} 个元素同为最大字号 {owners[0][1]:.0f}px：一页只能有一个视觉主角",
                 slide=sid)


def _count_numbers(text: str) -> int:
    import re
    return len(re.findall(r"\d+(?:\.\d+)?", text))


def _check_density(numbers: int, identity: dict, policy, rep, sid) -> None:
    target = identity.get("density")
    high = float(policy.num("thresholds.numeric_density_high"))
    if target == "sparse" and numbers > 6:
        rep.warn("DENSITY_OVER_TARGET",
                 f"本页判断为 sparse，但出现了 {numbers} 个数字：密度与判断不一致", slide=sid)


def _check_text_fit(element, ink, policy, rep, where) -> None:
    metrics = ink.get("metrics")
    max_lines = element.get("max_lines")
    if max_lines and metrics.count > int(max_lines):
        rep.error("TEXT_OVERFLOW",
                  f"文本排到 {metrics.count} 行，超过本档上限 {max_lines} 行：" + " / ".join(metrics.lines[:3]),
                  **where)
    if metrics.truncated:
        rep.error("TEXT_OVERFLOW", "文本被行数上限截断：内容丢失", **where)
    frame_w = float(element["width"])
    frame_h = float(element["height"])
    if ink["h"] > frame_h + 0.5:
        rep.error("TEXT_OVERFLOW",
                  f"墨迹高 {ink['h']:.1f}px > 框高 {frame_h:.1f}px（{metrics.count} 行）：删内容或加高框，不缩字号",
                  **where)
    if metrics.widths and max(metrics.widths) > frame_w + 0.5:
        rep.error("TEXT_TOO_WIDE",
                  f"最长一行 {max(metrics.widths):.1f}px > 框宽 {frame_w:.1f}px", **where)
    for issue in orphan_issues(metrics, policy):
        rep.warn(issue.kind.upper(), issue.detail, **where)


def _check_overlaps(boxes, policy, rep, sid) -> None:
    tol = float(policy.num("thresholds.ink_overlap_px"))
    contain = float(policy.num("thresholds.frame_overlap_ratio"))
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a["kind"] == "rule" or b["kind"] == "rule":
                continue
            if _contains(a, b, contain) or _contains(b, a, contain):
                continue
            overlap = _overlap_area(a, b)
            if overlap > tol * tol:
                rep.error("ELEMENT_OVERLAP",
                          f"{a['id']}({a['kind']}) 与 {b['id']}({b['kind']}) 墨迹相交 "
                          f"{overlap:.0f}px²：先重排，不要靠缩间距", slide=sid, element=a["id"])


def _contains(a, b, ratio) -> bool:
    area = a["w"] * a["h"]
    if area <= 0:
        return False
    inter = _overlap_area(a, b)
    return inter >= ratio * min(area, b["w"] * b["h"])


def _overlap_area(a, b) -> float:
    dx = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
    dy = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    return dx * dy if dx > 0 and dy > 0 else 0.0


def _check_chart(element, policy, rep, where, identity) -> None:
    kind = element.get("chart_kind")
    if kind not in CHART_KINDS:
        rep.error("CHART_INVALID", f"未知 chart_kind：{kind!r}（合法值 {sorted(CHART_KINDS)}）", **where)
        return
    categories = element.get("categories") or []
    series = element.get("series") or []
    if not categories or not series:
        rep.error("CHART_INVALID", "图表缺 categories 或 series：空载荷不画图", **where)
        return
    limit = int(policy.num("thresholds.chart_max_categories"))
    if len(categories) > limit:
        rep.error("CHART_INVALID", f"类别 {len(categories)} 个 > 上限 {limit}：改表格或拆分", **where)
    names = [s.get("name") for s in series]
    if any(not str(n or "").strip() for n in names) or len(set(names)) != len(names):
        rep.error("CHART_INVALID", "多序列必须各有非空且唯一的 name", **where)
    for s in series:
        values = s.get("values") or []
        if len(values) != len(categories):
            rep.error("CHART_INVALID",
                      f"序列 {s.get('name')!r} 的 values 数量 {len(values)} ≠ 类别 {len(categories)}", **where)
        for v in values:
            if not _finite(v):
                rep.error("CHART_INVALID", f"序列 {s.get('name')!r} 含非有限值 {v!r}", **where)
    if kind == "donut":
        values = [v for s in series for v in (s.get("values") or [])]
        if any(v < 0 for v in values) or sum(values) <= 0:
            rep.error("CHART_INVALID", "环形图：数值须非负且总和 > 0", **where)
        if len(series) != 1:
            rep.error("CHART_INVALID", "环形图只接受一个序列：多序列环是伪造的整体", **where)
        if identity.get("has_numeric_evidence") and sum(values) and abs(sum(values) - 100) > 0.51:
            rep.warn("DONUT_NOT_PERCENT",
                     f"环形图数值合计 {sum(values):.1f} 不是 100：确认是占比还是量（量不该画环）", **where)
    _check_chart_scales(series, policy, rep, where)
    if kind in NUMERIC_KINDS and len(categories) == 1:
        rep.warn("CHART_SINGLE_CATEGORY", "单类别数值图不如一个数字主张", **where)


def _series_unit(name: str, policy) -> str | None:
    """序列名里的单位（最长优先，避免把「万元」读成「元」）。"""
    tokens = sorted((str(t) for t in policy.req("typography.number_unit.unit_patterns")),
                    key=len, reverse=True)
    low = str(name or "")
    for token in tokens:
        if token and token in low:
            return token
    return None


def _check_chart_scales(series: list[dict], policy, rep, where: dict) -> None:
    """一张图只能画一套量纲。

    两类真实失效：
      * 混量纲：序列名里出现两种单位（门店数「家」与占比「%」）却被画进同一张图——
        读者会拿两条不可换算的线互相读；
      * 量纲悬殊：两序列均值差 100 倍以上，画在同一根轴上会把小的压成零。
    单位词表来自策略（typography.number_unit.unit_patterns），阈值来自
    thresholds.chart_scale_skew_ratio；这里只做判定，不解释业务。
    """
    units = {}
    for item in series:
        unit = _series_unit(item.get("name") or "", policy)
        if unit:
            units.setdefault(unit, []).append(item.get("name"))
    if len(units) > 1:
        detail = "；".join(f"{u}: {', '.join(map(str, names))}" for u, names in sorted(units.items()))
        rep.error("CHART_MIXED_UNITS",
                  f"同一张图混了 {len(units)} 种单位（{detail}）：不可换算的量不能共用一张图或一根轴",
                  **where)
        return
    means = []
    for item in series:
        values = [abs(float(v)) for v in (item.get("values") or []) if isinstance(v, (int, float))]
        if values and sum(values) > 0:
            means.append(sum(values) / len(values))
    if len(means) > 1 and min(means) > 0:
        skew = max(means) / min(means)
        limit = float(policy.num("thresholds.chart_scale_skew_ratio"))
        if skew > limit:
            rep.warn("CHART_SCALE_SKEW",
                     f"序列均值相差 {skew:.0f} 倍（阈值 {limit:.0f}）：小的那条会被压平，"
                     "考虑拆图或改用两个指标块", **where)


def _image_path(element: dict, deck: dict):
    from assets import resolve_source
    base = (deck.get("meta") or {}).get("base_dir") or None
    return resolve_source(element.get("source"), base)


def _check_image(element: dict, deck: dict, policy, rep: _Report, where: dict) -> None:
    """图片的硬检查：缺文件、格式、分辨率、裁切、预算、声明一致性、路径边界。

    这一层不评价「好不好看」，只回答「这张图能不能用、会不会读不清」。
    """
    from assets import AssetError, cover_crop, inspect, required_source_px
    assets = policy.req("assets")
    source = str(element.get("source") or "").strip()
    if not source:
        rep.error("ASSET_MISSING", "image 元素没有 source：图位必须有真实文件", **where)
        return
    path = _image_path(element, deck)
    if path is None:
        rep.error("ASSET_MISSING",
                  f"素材文件不存在：{source!r}（相对 {((deck.get('meta') or {}).get('base_dir') or '当前目录')}）",
                  **where)
        return
    try:
        facts = inspect(path, policy)
    except AssetError as exc:
        rep.error("ASSET_FORMAT", str(exc), **where)
        return
    suffix = path.suffix.lower()
    if suffix not in set(assets["formats"]):
        rep.error("ASSET_FORMAT", f"{suffix} 不在白名单 {assets['formats']}", **where)
    expected = {".jpg": "JPEG", ".jpeg": "JPEG", ".png": "PNG", ".webp": "WEBP"}.get(suffix)
    if expected and facts["format"] != expected:
        rep.error("ASSET_FORMAT",
                  f"后缀是 {suffix}，实际解码格式是 {facts['format']}：改后缀不改格式，PowerPoint 里可能打不开",
                  **where)
    need = required_source_px((element["width"], element["height"]), policy)
    if facts["long_side"] < need:
        rep.error("ASSET_RESOLUTION",
                  f"源图长边 {facts['long_side']}px < 需要 {need}px"
                  f"（放置 {max(element['width'], element['height']):.0f}px × oversample "
                  f"{assets['oversample']}）：放大必糊", **where)
    base = (deck.get("meta") or {}).get("base_dir")
    if base:
        try:
            path.resolve().relative_to(Path(base).expanduser().resolve())
        except ValueError:
            rep.warn("ASSET_OUTSIDE_PROJECT",
                     f"素材 {path} 在 brief 目录之外：交付时要单独确认授权与随包分发", **where)
    fit = element.get("fit") or "cover"
    if fit not in ("cover", "contain"):
        rep.error("ASSET_DECLARATION", f"fit={fit!r} 非法（cover / contain）", **where)
    if fit == "cover":
        crop = cover_crop((facts["width"], facts["height"]), (element["width"], element["height"]))
        limit = float(assets["crop_requires_reason_above"])
        if crop["loss"] > limit and not str(element.get("reason") or "").strip():
            rep.error("ASSET_CROP",
                      f"cover 裁掉 {crop['loss']}× 画面（阈值 {limit}×）却没有写 reason："
                      "先说明为什么可以裁", **where)
    if element.get("function") in ("frame", "emotion") and \
            assets["decorative_requires_reason"] and not str(element.get("reason") or "").strip():
        rep.warn("ASSET_DECLARATION", "装饰性素材没有写职责：说不出「拿掉它损失什么」就是可删的", **where)


def _check_media_contract(slide: dict, deck: dict, elements: list, boxes: list, policy, rep: _Report) -> None:
    """页级素材契约：判断说要图就必须有图；有图就必须在判断里说清职责。"""
    sid = slide.get("id")
    assets = policy.req("assets")
    media = (slide.get("meta") or {}).get("media") or {}
    images = [e for e in elements if e.get("type") == "image"]
    decision = media.get("decision", "none")
    if len(images) > int(assets["max_per_page"]):
        rep.error("ASSET_BUDGET",
                  f"本页 {len(images)} 张图 > 上限 {assets['max_per_page']}：图片不是靠数量建立氛围", slide=sid)
    if decision in ("required", "existing") and not images:
        rep.error("ASSET_MISSING",
                  f"判断（{media.get('source')}）认为本页需要 {media.get('function')} 素材，但页面上没有图片元素："
                  "判断没有落到执行", slide=sid)
    if decision == "none" and images:
        rep.warn("ASSET_DECLARATION", "页面上有图，但素材判断是 none（未声明）：补写职责或删图", slide=sid)
    for image in images:
        declared_role = str(image.get("function") or "")
        if media.get("function") and declared_role and declared_role != media["function"]:
            rep.warn("ASSET_DECLARATION",
                     f"元素 function={declared_role} 与判断 function={media['function']} 不一致", slide=sid)
        # 图片不得侵入来源区（出处文字必须始终可读）
        zone = (slide.get("meta") or {}).get("source_zone")
        if zone and _intersects(image, zone):
            rep.error("SOURCE_ZONE_COLLISION", "图片覆盖了来源区：出处被遮住等于没有出处",
                      slide=sid, element=image.get("id"))


def _check_content_truncated(slide: dict, rep: "_Report", sid: str) -> None:
    """compose 记账了被丢弃的条目：告警，不静默。

    「一页六道工序」遇到五列上限时，旧实现直接 items[:5]，第六道工序在产物里消失，
    而 guard 看不见——页面 PASS，内容却少了一条。丢弃数量由 compose 写进 meta，
    这里只负责把它变成可见的告警。
    """
    dropped = int((slide.get("meta") or {}).get("steps_dropped") or 0)
    if dropped > 0:
        rep.warn("CONTENT_TRUNCATED",
                 f"有 {dropped} 个步骤超出列数上限未排版：删条目或改构图（不静默丢弃）",
                 slide=sid)


def _check_evidence_laid_out(slide: dict, rep: "_Report", sid: str) -> None:
    """作者声明的 metrics 必须在页面上有落位。

    案例 B s04/s10 的回归：作者写了「3 个庄园 / 640 米 / 6 天」，构图选 image_narrative，
    而该配方当时没有指标槽——数字在 plan.json 和 brief 里都在，成片里一个都没有，guard 却 PASS。
    静默丢证据比排版难看严重得多：评审看到的是一页没有证据的图片页。
    """
    declared = (slide.get("meta") or {}).get("metrics") or []
    if not declared:
        return
    laid_out = [e for e in slide.get("elements") or []
                if str(e.get("role") or "").startswith(
                    ("metric_value", "hero_value", "left_value", "right_value"))]
    if not laid_out:
        rep.error("EVIDENCE_NOT_LAID_OUT",
                  f"声明了 {len(declared)} 个指标，页面上没有任何指标元素：换构图或删数据",
                  slide=sid)


def _intersects(element: dict, zone: dict) -> bool:
    dx = min(element["x"] + element["width"], zone["x"] + zone["width"]) - max(element["x"], zone["x"])
    dy = min(element["y"] + element["height"], zone["y"] + zone["height"]) - max(element["y"], zone["y"])
    return dx > 1 and dy > 1


def _overlaps(a: tuple, b: tuple) -> bool:
    """两个像素框是否相交（1px 容差：相邻不算压住）。"""
    dx = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    dy = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return dx > 1 and dy > 1


def _hex_of_token(value, colors: dict) -> str | None:
    """token → #RRGGBB；非法/未知返回 None（未知 token 由 TOKEN_UNKNOWN 单独拦）。"""
    text = str(value or "").strip()
    if text.startswith("#") and len(text) == 7:
        return text.upper()
    return (colors.get(text) or "").upper() or None


def _check_text_contrast(slide: dict, elements: list, policy, rep: _Report, colors: dict) -> None:
    """文字与底板的对比度：用策略阈值的实测值，不靠观感。

    只查「直接落在页面底色上的文字」：压在图上的文字由 `_check_text_over_image` 按像素实测；
    压形状上的文字底板不确定，不在这里猜。结论是告警而非阻断——投影环境未知，
    但对比度不达标的稿子在会议室里就是读不清。
    """
    from assets import contrast_ratio, luma
    rule = policy.req("colors.contrast")
    body_min = float(rule["body_min"])
    large_min = float(rule["large_min"])
    large_px = float(rule["large_size_px"])
    background = _hex_of_token(slide.get("background") or "background", colors)
    if not background:
        return
    bg_luma = luma(tuple(int(background.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)))
    pairing = policy.req("colors.pairing")
    allowed = {str(t) for t in (pairing["light_background"] if bg_luma > 0.5
                                else pairing["dark_background"])}
    covered = [(e["x"], e["y"], e["width"], e["height"]) for e in elements
               if e.get("type") in ("shape", "image", "chart")]
    sid = slide.get("id")
    for element in elements:
        if element.get("type") != "text":
            continue
        if str(element.get("role") or "").startswith("legend_key"):
            continue  # 色标是非语言图形，不适用正文文本对比阈值
        box = (element["x"], element["y"], element["width"], element["height"])
        if any(_overlaps(box, other) for other in covered):
            continue                                  # 底板不确定，交给像素实测那一层
        token = str(element.get("color") or "ink")
        hex_value = _hex_of_token(token, colors)
        if not hex_value:
            continue
        fg_luma = luma(tuple(int(hex_value.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)))
        ratio = contrast_ratio(fg_luma, bg_luma)
        floor = large_min if float(element.get("size") or 0) >= large_px else body_min
        if ratio < floor:
            rep.warn("TEXT_CONTRAST_LOW",
                     f"{element.get('id')}：{token} 落在 {slide.get('background') or 'background'} 上，"
                     f"对比度 {ratio:.2f} < 下限 {floor}（字号 {element.get('size')}）："
                     "换 token、换底板或加分隔", slide=sid, element=element.get("id"))
        elif token in colors and token not in allowed:
            rep.warn("TOKEN_PAIRING",
                     f"{element.get('id')}：{token} 不在 "
                     f"{'light' if bg_luma > 0.5 else 'dark'}_background 的配对表里："
                     "配色是策略，越界要显式声明", slide=sid, element=element.get("id"))


def _check_text_over_image(slide: dict, deck: dict, elements: list, boxes: list, policy, rep: _Report) -> None:
    """文字压图：必须有安全区声明，且实测对比度达标（量的是观众看见的那块）。

    这是「预览好看、投影读不清」最常见的来源，所以用像素实测，不用猜。
    """
    from assets import contrast_ratio, luma, region_stats
    from render import visible_image
    assets = policy.req("assets")
    rule = assets["text_over_image"]
    colors = (deck.get("theme") or {}).get("colors") or {}
    images = [e for e in elements if e.get("type") == "image"]
    if not images:
        return
    sid = slide.get("id")
    for image in images:
        layer = visible_image((image["width"], image["height"]), image.get("source"),
                              image.get("fit") or "cover",
                              (deck.get("meta") or {}).get("base_dir") or None, policy)
        if layer is None:
            continue
        covered = [b for b in boxes if b["kind"] == "text" and b.get("role") not in ("source", "note")
                   and _overlap_ratio(b, image) > 0.15]
        if not covered:
            continue
        if rule["require_safe_area"] and not image.get("safe_area"):
            rep.error("ASSET_SAFE_AREA",
                      f"{len(covered)} 处文字压在图上，但没有声明 safe_area：先定文字落点，再谈构图",
                      slide=sid, element=image.get("id"))
            continue
        area = image.get("safe_area") or {}
        if area and float(area.get("width") or 0) < float(assets["safe_area"]["min_width_ratio"]):
            rep.error("ASSET_SAFE_AREA",
                      f"声明的 safe_area 宽度 {float(area.get('width') or 0):.2f} 小于下限 "
                      f"{assets['safe_area']['min_width_ratio']}：这么窄的安静面装不下标题",
                      slide=sid, element=image.get("id"))
        for box in covered:
            ratio_box = ((box["x"] - image["x"]) / image["width"],
                         (box["y"] - image["y"]) / image["height"],
                         box["w"] / image["width"], box["h"] / image["height"])
            stats = region_stats(layer, ratio_box, policy)
            if not stats:
                continue
            text_color = str(box["element"].get("color") or "ink")
            hex_value = colors.get(text_color, text_color if text_color.startswith("#") else "#14161A")
            text_luma = luma(tuple(int(hex_value.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)))
            delta = abs(stats["mean_luma"] - text_luma)
            contrast = contrast_ratio(stats["mean_luma"], text_luma)
            if delta < rule["min_luma_delta"] or contrast < rule["min_contrast_ratio"]:
                rep.error("ASSET_CONTRAST",
                          f"{box['id']} 压在图上：亮度差 {delta:.2f}（下限 {rule['min_luma_delta']}）"
                          f"、对比度 {contrast:.2f}（下限 {rule['min_contrast_ratio']}）："
                          "换文字色、加遮罩或换图位", slide=sid, element=box["id"], fix=FIX_HINTS["ASSET_CONTRAST"])
            if stats.get("flat_ratio", 1.0) < float(assets["qc"]["min_negative_ratio"]):
                rep.warn("ASSET_QUIET_AREA",
                         f"{box['id']} 所在区域平坦占比 {stats['flat_ratio']:.2f} < 下限 "
                         f"{assets['qc']['min_negative_ratio']}：这不算留白，字会陷进纹理里",
                         slide=sid, element=box["id"])
            if stats["std"] > float(assets["qc"]["busy_std"]):
                rep.warn("ASSET_TEXTURE",
                         f"{box['id']} 所在图像区域纹理过密（std {stats['std']}）：小字会读不清",
                         slide=sid, element=box["id"])


def _overlap_ratio(box: dict, image: dict) -> float:
    area = max(1.0, box["w"] * box["h"])
    dx = min(box["x"] + box["w"], image["x"] + image["width"]) - max(box["x"], image["x"])
    dy = min(box["y"] + box["h"], image["y"] + image["height"]) - max(box["y"], image["y"])
    return (dx * dy / area) if dx > 0 and dy > 0 else 0.0


def _check_metric_units(metric_units, rep) -> None:
    for metric, units in (metric_units or {}).items():
        real = {u for u in units or [] if u}
        if len(real) > 1:
            rep.error("UNIT_INCONSISTENT", f"指标「{metric}」跨页出现多个单位 {sorted(real)}：同一指标必须同单位")


def _finish(rep: _Report, deck: dict, mode: str) -> dict:
    errors = rep.errors
    blocking = set(rep.policy.req("verification.blocking_codes"))
    groups: dict[str, list] = {}
    for item in errors:
        groups.setdefault(item["code"], []).append(item)
    return {
        "schema": "pvad-guard-v1",
        "mode": mode,
        "status": "BLOCK" if any(e["code"] in blocking for e in errors) else "PASS",
        "errors": errors,
        "warnings": rep.warnings,
        "fix_groups": [{"code": code, "count": len(items), "hint": FIX_HINTS.get(code, ""),
                        "items": items[:8]} for code, items in sorted(groups.items())],
        "stats": rep.stats,
        "degraded": not rep.stats["measured"],
        "disclaimer": "guard 只拦硬错误：PASS 不代表构图、层级或排版质量通过（SKILL「四层架构与权威源」「证据层级」）。",
    }
