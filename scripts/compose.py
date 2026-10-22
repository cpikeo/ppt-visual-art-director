#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""compose.py · 把判断落成几何（确定性推导，可被作者覆盖）。

边界：
  * 输入 plan（判断 + 依据）+ brief（原文）。输出 deck spec（元素、坐标、字号、颜色）。
  * 每个元素的 layout_source 记录来源：plan（配方推导）| author（作者声明）。
  * 配方给的是「有理由的候选」：why/fails 写在 policy/composition.yaml，作者声明 composition 时原样生效。
  * 不做视觉审美判断的伪装：几何只是判断的落点，能否成立由 guard + 人工验收决定。
"""
from __future__ import annotations

import copy

from typography import FontBook, NBSP, bind_units, ink_height, wrap

DECK_SCHEMA = "pvad-deck-v1"


def compose_deck(brief: dict, plan: dict, policy, *, fontbook: FontBook | None = None,
                 base_dir=None) -> dict:
    fontbook = fontbook or FontBook(policy)
    canvas = policy.req("canvas")
    theme = _theme(brief, policy)
    slides = []
    for page in plan["pages"]:
        source_slide = brief["slides"][page["index"]]
        slides.append(_compose_slide(source_slide, page, policy, canvas, theme, fontbook, brief,
                                     base_dir))
    return {
        "schema": DECK_SCHEMA,
        "canvas": {"width": canvas["width"], "height": canvas["height"],
                   "ppt_dpi": canvas["ppt_dpi"]},
        "theme": theme,
        "meta": {
            "title": (brief.get("meta") or {}).get("title") or (brief["slides"][0]["title"] if brief["slides"] else ""),
            "audience": brief.get("audience", ""),
            "decision": brief.get("decision", ""),
            "visual_world": plan["deck"]["visual_world"],
        "base_dir": str(base_dir) if base_dir else "",
            "policy_sha256": plan["policy_sha256"],
            "brief_sha256": plan["brief_sha256"],
            "plan_schema": plan["schema"],
            "metric_units": _metric_registry(brief, plan),
        },
        "slides": slides,
    }


def _theme(brief: dict, policy) -> dict:
    colors = policy.colors()
    brand = (brief.get("brand") or {}).get("colors") or {}
    for token, value in brand.items():
        if token in colors:            # 非法 token 由 brief schema 拦住；这里再防一次静默回落
            colors[token] = value.upper()
    fonts = {"cn": policy.req("typography.families.cn"), "latin": policy.req("typography.families.latin")}
    for key, value in ((brief.get("brand") or {}).get("fonts") or {}).items():
        if key in fonts and value:
            fonts[key] = value
    return {"colors": colors, "fonts": fonts,
            "accent_budget": policy.req("colors.accent_budget_per_page")}


def _metric_registry(brief: dict, plan: dict) -> dict:
    """跨页同指标单位登记：同一指标在不同页用不同单位 = 口径事故，交 guard 拦。"""
    registry: dict[str, set] = {}
    for page in plan["pages"]:
        slide = brief["slides"][page["index"]]
        unit = slide.get("unit") or brief.get("unit") or ""
        for metric in page["understanding"]["atoms"]["metrics"]:
            label = (metric.get("label") or "").strip()
            if label:
                registry.setdefault(label, set()).add(metric.get("unit") or unit)
    return {label: sorted(u for u in units) for label, units in registry.items()}


# ──────────────────────────────────────────────────────────────────────────
# 几何
# ──────────────────────────────────────────────────────────────────────────
class Frame:
    def __init__(self, canvas: dict, policy):
        cp = policy.req("canvas")
        self.cw, self.ch = float(canvas["width"]), float(canvas["height"])
        self.m = {k: float(cp["margin"][k]) for k in ("left", "right", "top", "bottom")}
        self.snap_px = float(policy.num("thresholds.snap"))
        self.gutter = float(cp["gutter"])
        self.content = (self.m["left"], self.m["top"],
                        self.cw - self.m["left"] - self.m["right"],
                        self.ch - self.m["top"] - self.m["bottom"])

    def snap(self, value: float) -> float:
        return round(value / self.snap_px) * self.snap_px

    def zone(self, zone) -> tuple[float, float, float, float]:
        x0, y0, w, h = self.content
        return (self.snap(x0 + zone[0] * w), self.snap(y0 + zone[1] * h),
                self.snap(zone[2] * w), self.snap(zone[3] * h))

    def columns(self, zone, count: int, *, gutter: float | None = None):
        g = self.gutter if gutter is None else gutter
        x, y, w, h = self.zone(zone)
        cell = (w - g * (count - 1)) / count
        return [(self.snap(x + i * (cell + g)), y, self.snap(cell), h) for i in range(count)]


def _text_el(eid: str, role: str, text: str, box, *, policy, fontbook: FontBook,
             size_key: str, color: str, bold: bool = False, align: str = "left",
             anchor: str = "top", max_lines: int | None = None, size_px: float | None = None) -> dict:
    x, y, w, h = box
    size = float(size_px or policy.size(size_key))
    lh = policy.line_height(size_key)
    tracking = policy.tracking(size_key)
    limit = max_lines if max_lines is not None else policy.max_lines(size_key)
    metrics = wrap(text, w, size, policy=policy, fontbook=fontbook, bold=bold, tracking_pt=tracking,
                   max_lines=limit)
    needed = ink_height(metrics.count, size, lh, policy)
    # 框高取「实际需要」与「配方给的高度」的较小者：装不下时墨迹会溢出框外，
    # 由 guard 的墨迹相交 / 垂直溢出判据抓住，而不是靠把框撑大来掩盖。
    height = round(min(max(needed, 1.0), h)) if h else round(needed)
    return {
        "id": eid, "type": "text", "role": role,
        "x": x, "y": y, "width": w, "height": round(height),
        "size": size, "line_height": lh, "tracking": tracking,
        "color": color, "bold": bold, "align": align, "anchor": anchor,
        "max_lines": limit, "wrap": True, "text": text,
        "layout_source": "plan",
        "_measured": {"lines": metrics.count, "degraded": metrics.degraded,
                      "max_line_width": round(max(metrics.widths), 1) if metrics.widths else 0.0},
    }


def _shape_el(eid: str, role: str, box, *, shape: str = "line", color: str = "hairline",
              width_px: float = 1.0) -> dict:
    x, y, w, h = box
    return {"id": eid, "type": "shape", "role": role, "shape": shape,
            "x": x, "y": y, "width": max(w, 0.0), "height": max(h, width_px),
            "fill": {"type": "none"}, "stroke": color, "stroke_width": width_px,
            "layout_source": "plan"}


# ──────────────────────────────────────────────────────────────────────────
# 单页
# ──────────────────────────────────────────────────────────────────────────
def _compose_slide(slide: dict, page: dict, policy, canvas: dict, theme: dict,
                   fontbook: FontBook, brief: dict | None = None, base_dir=None) -> dict:
    composition_name = page["decision"]["composition"]
    downgraded = ""
    # 形状跟着证据走：声明为曲线但没有时间读数时，不画空轴——降级并留下理由
    if composition_name == "series_chart" and not slide.get("chart"):
        downgraded = "series_chart→metric_row：缺时间读数（categories/series），不画空轴"
        composition_name = "metric_row"
    if composition_name == "part_whole" and not (slide.get("chart") or page["understanding"]["atoms"]["metrics"]):
        downgraded = "part_whole→metric_row：缺分项数据，不画空占比图"
        composition_name = "metric_row"
    recipe = policy.recipe(composition_name)
    frame = Frame(canvas, policy)
    atoms = copy.deepcopy(page["understanding"]["atoms"])
    claim = page["claim"]
    elements: list[dict] = []
    emphasis_notes: list[str] = []
    dropped_steps = [0]
    accent_used = [0]
    source_text = atoms.get("source") or slide.get("source") or ""

    def add_text_box(role: str, text: str, box, size_key: str, color: str, **kw) -> dict | None:
        """box 已是像素框（由配方槽位或 columns() 算好），不再二次变换。"""
        if not str(text or "").strip():
            return None
        eid = f"{slide['id']}.{role}"
        el = _text_el(eid, role, text, box, policy=policy, fontbook=fontbook,
                      size_key=size_key, color=color, **kw)
        if color == "accent":
            accent_used[0] += 1
        elements.append(el)
        return el

    def add_text(role: str, text: str, zone, size_key: str, color: str, **kw) -> dict | None:
        """zone 是内容区比例坐标；这里统一换算成像素框。"""
        return add_text_box(role, text, frame.zone(zone), size_key, color, **kw)

    slots = {slot["role"]: slot for slot in recipe["slots"]}
    composition = composition_name

    # 通用头部：眉标（eyebrow 由判断派生：3–6 字，只做定位不做内容）
    if "eyebrow" in slots:
        eyebrow = _eyebrow(page, slide)
        add_text("eyebrow", eyebrow, slots["eyebrow"]["zone"], "micro", "muted")

    # 主张
    claim_text = claim["text"] or slide.get("title", "")
    claim_absent = claim["source"] == "absent"
    if "claim" in slots and claim_text:
        # 主张缺席时，标题只是话题标签：用次要档位与墨色，不让它抢主角位
        # 字号档：配方槽位声明优先（尺寸是配方的决定），缺省才退回配方的 ladder
        base_ladder = str(slots["claim"].get("size") or _claim_ladder(composition, policy))
        claim_ladder = base_ladder
        emphasis_applied = False
        if not claim_absent:
            boosted = _emphasised(base_ladder, page["decision"].get("emphasis"), policy)
            if boosted != base_ladder and _emphasis_fits(
                    claim_text, frame.zone(slots["claim"]["zone"]), boosted,
                    policy=policy, fontbook=fontbook,
                    bold=bool(slots["claim"].get("bold", True))):
                claim_ladder, emphasis_applied = boosted, True
            elif boosted != base_ladder:
                # 判断成立但物理上装不下：如实记下来，交给下一轮改内容而不是改字号
                claim_emphasis_note = (f"强调档 {page['decision'].get('emphasis')} 想上移一档到 {boosted}，"
                                       "但主张在这个框里装不下：保持基准档，考虑删内容")
                emphasis_notes.append(claim_emphasis_note)
        add_text("claim", claim_text, slots["claim"]["zone"],
                 "lede" if claim_absent else claim_ladder,
                 "muted" if claim_absent else slots["claim"].get("color", "ink"),
                 bold=not claim_absent, anchor=slots["claim"].get("anchor", "top"),
                 max_lines=int(slots["claim"].get("max_lines")
                               or policy.max_lines(claim_ladder if not claim_absent
                                                   else _claim_ladder(composition, policy))))

    media_el = _compose_media(slots, slide, page, frame, policy, elements, base_dir)
    dropped_steps[0] = _fill_body(composition, elements, slots, page, atoms, frame, policy, fontbook, add_text,
               add_text_box, source_text, accent_used, slide.get("chart"))

    # 来源：与数字同页可读；没有来源就不写假出处
    if "source" in slots and source_text:
        add_text("source", source_text, slots["source"]["zone"], "source", "muted")

    source_box = frame.zone(slots["source"]["zone"]) if "source" in slots else None
    meta = {
        "media": {k: v for k, v in (page["decision"].get("media") or {}).items()},
        "source_zone": ({"x": source_box[0], "y": source_box[1],
                         "width": source_box[2], "height": source_box[3]} if source_box else None),
        "source": source_text,
        "unit": _page_unit(slide, atoms, brief),
        "period": slide.get("period") or (brief or {}).get("period") or "",
        "basis": slide.get("basis") or (brief or {}).get("basis") or "",
        "focus": page["decision"]["focus"],
        "density": page["decision"]["density"],
        "emphasis": page["decision"]["emphasis"],
        "emphasis_applied": bool(emphasis_applied),
        "emphasis_notes": emphasis_notes,
        "rhythm": page["decision"]["rhythm"]["move"],
        "role": page["decision"]["role"],
        "reading_seconds": page["production"]["typography"]["reading_seconds"],
        "open_questions": [q["code"] for q in page.get("open_questions") or []],
        "has_numeric_evidence": bool(atoms.get("metrics")) or bool(slide.get("chart")),
        "claim_absent": claim_absent,
        "metrics": [{"label": m.get("label") or "", "value": m.get("value"), "unit": m.get("unit") or ""}
                    for m in atoms.get("metrics") or []],
        "measured": not fontbook.degraded,
        "steps_dropped": dropped_steps[0],
    }
    return {"id": slide["id"], "index": page["index"],
            "composition": composition,
            "composition_downgraded": downgraded,
            "layout_source": "author" if slide.get("composition") else "plan",
            "composition_why": page["decision"]["composition_why"],
            "accent_elements": accent_used[0],
            "elements": elements,
            "meta": meta}


def _media_brief(element: dict, media: dict, policy) -> dict:
    """给「需要生产」的图配上生产说明（提示词）。

    只对缺失/损坏的素材生成：已有文件的图走 QC 与对账，不需要提示词。
    提示词本身是确定性翻译（prompt.py），这里只负责把判断结论与几何传进去。
    """
    from prompt import build
    return build(policy=policy,
                 function=element.get("function") or "context",
                 role=media.get("role") or "subject",
                 subject=media.get("subject") or "",
                 reason=element.get("reason") or media.get("reason") or "",
                 box=(element["x"], element["y"], element["width"], element["height"]),
                 safe_area=element.get("safe_area"),
                 text_color=element.get("text_color") or "",
                 avoid=media.get("avoid") or ())


def _page_unit(slide: dict, atoms: dict, brief: dict | None) -> str:
    """页级单位：作者声明优先；否则只在「全页指标同单位且都有单位」时取该单位。

    不做「deck 级单位兜底」：把 万元 套到「门店数量」上就是口径事故。
    """
    if slide.get("unit"):
        return slide["unit"]
    metrics = [m for m in (atoms.get("metrics") or []) if str(m.get("value", "")).strip() != ""]
    if metrics:
        units = {m.get("unit") or "" for m in metrics}
        if len(units) == 1 and "" not in units:
            return units.pop()
        return ""
    return (brief or {}).get("unit") or ""


def _claim_ladder(composition: str, policy) -> str:
    """字号档位是配方属性（composition.yaml），不许在这里维护第二份映射。"""
    return str(policy.recipe(composition)["ladder"])


def _emphasis_fits(text: str, box, ladder: str, *, policy, fontbook: FontBook, bold: bool) -> bool:
    """放大一档还装得下吗？装不下就退回基准档——不靠缩字号硬塞，也不靠撑框掩盖。"""
    x, y, w, h = box
    size = float(policy.size(ladder))
    metrics = wrap(text, w, size, policy=policy, fontbook=fontbook, bold=bold,
                   tracking_pt=policy.tracking(ladder), max_lines=policy.max_lines(ladder))
    if metrics.truncated:
        return False
    needed = ink_height(metrics.count, size, policy.line_height(ladder), policy)
    return needed <= (h + 0.5 if h else needed)


def _emphasised(ladder: str, emphasis: str, policy) -> str:
    """强调档 → 字号档上移。级数来自 typography.emphasis_ladder_boost。

    只用「更大」这一件事表示强调：不额外加粗、不换色、不加装饰——
    强调如果同时改三件事，读者看不出是哪一件在起作用，也测不出判断有没有落地。
    """
    boost = int((policy.req("typography.emphasis_ladder_boost") or {}).get(str(emphasis), 0))
    if boost <= 0:
        return ladder
    keys = list(policy.req("typography.scale"))          # 由大到小
    if ladder not in keys:
        return ladder
    return keys[max(0, keys.index(ladder) - boost)]


def _eyebrow(page: dict, slide: dict) -> str:
    """眉标 = 本页证据形状 + 序号的极短标记；不承担信息，只做定位。"""
    index = page["index"] + 1
    return f"{index:02d}"


def _fill_body(composition, elements, slots, page, atoms, frame, policy, fontbook, add_text,
               add_text_box, source_text, accent_used, chart=None):
    metrics = [m for m in atoms["metrics"] if str(m.get("value", "")).strip() != ""]
    if composition == "statement":
        if metrics and "hero_value" in slots:
            hero = metrics[0]
            value = f"{hero['value']}{NBSP if hero.get('unit') else ''}{hero.get('unit') or ''}"
            add_text("hero_value", value, slots["hero_value"]["zone"], "display", "primary", bold=True,
                     anchor="bottom")
            if hero.get("label") and "hero_unit" in slots:
                add_text("hero_unit", hero["label"], slots["hero_unit"]["zone"], "label", "muted")
        body = "\n".join(atoms["body_lines"] or atoms["sentences"][:3])
        if "lede" in slots and body:
            add_text("lede", body, slots["lede"]["zone"], "lede", "muted")
        if "note" in slots:
            notes = "\n".join(atoms["sentences"][:2]) if composition != "editorial_text" else ""
            add_text("note", notes, slots["note"]["zone"], "body", "muted")

    elif composition == "metric_row":
        _metric_cells(elements, slots, metrics, frame, policy, fontbook, add_text_box, accent_used)
        if "note" in slots:
            add_text("note", "\n".join(atoms["body_lines"][:2]), slots["note"]["zone"], "body", "muted")

    elif composition == "two_end":
        _two_end(elements, slots, metrics, frame, policy, fontbook, add_text_box, accent_used)
        if "note" in slots:
            add_text("note", "\n".join(atoms["body_lines"][:2]), slots["note"]["zone"], "body", "muted")

    elif composition == "series_chart":
        if chart:
            elements.append(_chart_el(f"{page['id']}.chart", "chart", frame.zone(slots["chart"]["zone"]),
                                      chart, policy))
        if "takeaway" in slots:
            takeaway = chart.get("takeaway") if chart else ""
            add_text("takeaway", takeaway or "\n".join(atoms["sentences"][:1]),
                     slots["takeaway"]["zone"], "body", "primary")

    elif composition == "part_whole":
        if chart:
            elements.append(_chart_el(f"{page['id']}.chart", "chart", frame.zone(slots["chart"]["zone"]),
                                      chart, policy))
        elif metrics:
            payload = _metrics_as_donut(metrics, policy)
            elements.append(_chart_el(f"{page['id']}.chart", "chart", frame.zone(slots["chart"]["zone"]),
                                      payload, policy))
        if "parts" in slots:
            lines = []
            for m in metrics:
                unit = m.get("unit") or ""
                lines.append(f"{m.get('label') or '分项'}　{m['value']}{NBSP if unit else ''}{unit}")
            add_text("parts", "\n".join(lines), slots["parts"]["zone"], "body", "ink")
        if "note" in slots:
            add_text("note", "\n".join(atoms["body_lines"][:2]), slots["note"]["zone"], "micro", "muted")

    elif composition == "sequence_steps":
        return _steps(elements, slots, atoms, frame, policy, fontbook,
                      add_text_box, accent_used) or 0

    elif composition == "structure_map":
        _groups(elements, slots, atoms, frame, policy, fontbook, add_text_box, accent_used)

    elif composition == "image_narrative":
        body = "\n".join(atoms["body_lines"] or atoms["sentences"][:2])
        if "lede" in slots and body:
            add_text("lede", body, slots["lede"]["zone"], "body", "muted")
        if metrics and "metric_cells" in slots:
            _metric_cells(elements, slots, metrics, frame, policy, fontbook, add_text_box, accent_used)

    elif composition == "editorial_text":
        body = "\n".join(atoms["body_lines"] or atoms["sentences"])
        if "lede" in slots:
            add_text("lede", body, slots["lede"]["zone"], "body", "ink")
        if "side_note" in slots:
            add_text("side_note", "\n".join(atoms["sentences"][-1:]), slots["side_note"]["zone"],
                     "body", "muted")
    return 0


def _compose_media(slots, slide, page, frame, policy, elements, base_dir) -> dict | None:
    """把作者声明的素材落成 image 元素。

    关键纪律：**文件缺失也要写元素**。悄悄不写会让「有图」的稿子在没有图的情况下通过检查；
    写进去之后 guard 才能报 ASSET_MISSING，让问题暴露在交付前。
    """
    from assets import AssetError, cover_crop, inspect, resolve_source
    media = page["decision"].get("media") or {}
    if media.get("decision", "none") == "none" or "media" not in slots:
        return None
    author = slide.get("asset") or {}
    box = frame.zone(slots["media"]["zone"])
    function = media.get("function") or "context"
    if function in ("frame", "emotion"):
        # 全幅背景：文字必须落在声明过的安全区里
        box = (0.0, 0.0, float(frame.cw), float(frame.ch))
    element = {
        "id": f"{slide['id']}.media", "type": "image", "role": "media",
        "x": box[0], "y": box[1], "width": box[2], "height": box[3],
        "source": author.get("path") or "",
        "fit": author.get("fit") or ("cover" if function in ("hero", "frame", "emotion") else "contain"),
        "function": function,
        "reason": media.get("reason") or "",
        "safe_area": author.get("safe_area") or None,
        "text_color": author.get("text_color") or "",
        "layout_source": "author",
    }
    path = resolve_source(element["source"], base_dir)
    if path is None:
        element["_missing"] = True
        element["brief"] = _media_brief(element, media, policy)
        elements.append(element)
        return element
    try:
        facts = inspect(path, policy)
    except AssetError as exc:
        element["_undecodable"] = str(exc)
        element["brief"] = _media_brief(element, media, policy)
        elements.append(element)
        return element
    element["asset_sha256"] = facts["sha256"]
    element["source_px"] = {"width": facts["width"], "height": facts["height"]}
    element["format"] = facts["format"]
    if element["fit"] == "cover":
        crop = cover_crop((facts["width"], facts["height"]), (element["width"], element["height"]))
        element["crop"] = crop["crop"]
        element["crop_loss"] = crop["loss"]
    elements.append(element)
    return element


def _metric_cells(elements, slots, metrics, frame, policy, fontbook, add_text_box, accent_used):
    metrics = metrics[:4]
    if not metrics:
        return
    zone = slots["metric_cells"]["zone"]
    cells = frame.columns(zone, len(metrics))
    size_key = "h2" if len(metrics) <= 3 else "lede"
    for i, (cell, metric) in enumerate(zip(cells, metrics)):
        x, y, w, h = cell
        value_size = policy.size(size_key)
        unit = metric.get("unit") or ""
        value = f"{metric['value']}{NBSP if unit else ''}{unit}"
        color = "ink"
        add_text_box(f"metric_value.{i + 1}", value, (x, y, w, h * 0.46), size_key, color, bold=True,
                 anchor="bottom", max_lines=1)
        if accent_used[0] and color == "accent":
            pass
        add_text_box(f"metric_label.{i + 1}", metric.get("label") or "", (x, y + h * 0.5, w, h * 0.2),
                 "label", "ink")
        add_text_box(f"metric_desc.{i + 1}", metric.get("desc") or metric.get("period") or "",
                 (x, y + h * 0.72, w, h * 0.2), "micro", "muted")


def _two_end(elements, slots, metrics, frame, policy, fontbook, add_text_box, accent_used):
    metrics = metrics[:2]
    if len(metrics) < 2:
        return
    left, right = metrics
    for side, metric, value_key, label_key in (("left", left, "left_value", "left_label"),
                                               ("right", right, "right_value", "right_label")):
        if value_key not in slots:
            continue
        unit = metric.get("unit") or ""
        color = "ink" if side == "left" else "primary"
        add_text_box(f"{value_key}", f"{metric['value']}{NBSP if unit else ''}{unit}",
                 frame.zone(slots[value_key]["zone"]), slots[value_key].get("size", "display"),
                 color, bold=True, anchor="bottom")
        add_text_box(f"{label_key}", metric.get("label") or "", frame.zone(slots[label_key]["zone"]),
                     "label", "muted")
    if "delta" in slots:
        delta = _delta_text(left, right)
        if delta:
            add_text_box("delta", delta, frame.zone(slots["delta"]["zone"]),
                         slots["delta"].get("size", "body"), "accent", bold=True, align="center")
    if "rule" in slots:
        x, y, w, h = frame.zone(slots["rule"]["zone"])
        elements.append({"id": "rule", "type": "shape", "role": "rule", "shape": "line",
                         "x": frame.snap(x), "y": frame.snap(y),
                         "width": 0.0, "height": frame.snap(h or 200.0),
                         "fill": {"type": "none"}, "stroke": "hairline", "stroke_width": 1.0,
                         "layout_source": "plan"})


def _delta_text(left: dict, right: dict) -> str:
    try:
        a = float(str(left["value"]).replace(",", ""))
        b = float(str(right["value"]).replace(",", ""))
    except (TypeError, ValueError):
        return ""
    unit = right.get("unit") or left.get("unit") or ""
    magnitude = abs(b - a)
    pretty = f"{magnitude:.2f}".rstrip("0").rstrip(".") if magnitude % 1 else f"{int(magnitude)}"
    return f"差 {pretty}{NBSP if unit else ''}{unit}"


def _steps(elements, slots, atoms, frame, policy, fontbook, add_text_box, accent_used):
    """横向工序。列数上限是策略（thresholds.max_steps），超出部分必须记账而不是消失。"""
    pool = atoms.get("steps") or [s for s in atoms["sentences"] if s][:4]
    limit = int(policy.num("thresholds.max_steps"))
    items = pool[:limit]
    if not items:
        return 0
    cells = frame.columns(slots["steps"]["zone"], len(items))
    for i, (cell, item) in enumerate(zip(cells, items)):
        x, y, w, h = cell
        add_text_box(f"step_index.{i + 1}", f"{i + 1:02d}", (x, y, w, h * 0.18), "micro", "muted")
        title, desc = _split_item(item)
        add_text_box(f"step_title.{i + 1}", title, (x, y + h * 0.2, w, h * 0.36), "body", "ink", bold=True)
        if desc:
            add_text_box(f"step_desc.{i + 1}", desc, (x, y + h * 0.58, w, h * 0.42), "micro", "muted")
        if i < len(items) - 1:
            elements.append({"id": f"step_rule.{i + 1}", "type": "shape", "role": "step_rule",
                             "shape": "line", "x": frame.snap(x + w + 4), "y": y,
                             "width": 0.0, "height": frame.snap(h * 0.8),
                             "fill": {"type": "none"}, "stroke": "hairline", "stroke_width": 1.0,
                             "layout_source": "plan"})
    return max(0, len(pool) - len(items))


def _groups(elements, slots, atoms, frame, policy, fontbook, add_text_box, accent_used):
    items = atoms["bullets"] or [s for s in atoms["sentences"] if s]
    items = items[:6]
    if not items:
        return
    cells = frame.columns(slots["groups"]["zone"], min(len(items), 3))
    per = (len(items) + len(cells) - 1) // len(cells)
    for i, cell in enumerate(cells):
        x, y, w, h = cell
        chunk = items[i * per:(i + 1) * per]
        if not chunk:
            continue
        title, desc = _split_item(chunk[0])
        add_text_box(f"group_title.{i + 1}", title, (x, y, w, h * 0.22), "body", "ink", bold=True)
        body = "\n".join([desc] + [_split_item(it)[0] for it in chunk[1:]]) if desc else \
            "\n".join(_split_item(it)[0] for it in chunk[1:])
        if body.strip():
            add_text_box(f"group_body.{i + 1}", body, (x, y + h * 0.26, w, h * 0.7), "body", "muted")


def _split_item(item: str) -> tuple[str, str]:
    m = __import__("re").match(r"^([^：:，,。]{1,16})[：:]\s*(.+)$", item.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return item.strip(), ""


def _metrics_as_donut(metrics: list[dict], policy) -> dict:
    values = []
    for m in metrics:
        try:
            values.append(float(str(m["value"]).replace(",", "").rstrip("%％")))
        except (TypeError, ValueError):
            values.append(0.0)
    return {"kind": "donut", "categories": [m.get("label") or f"分项{i + 1}"
                                            for i, m in enumerate(metrics)],
            "series": [{"name": "占比", "values": values}], "derived": "from_metrics"}


def _chart_el(eid: str, role: str, box, chart: dict, policy) -> dict:
    x, y, w, h = box
    categories = list(chart.get("categories") or [])
    series = [{"name": s.get("name") or f"系列{i + 1}", "values": [float(v) for v in s.get("values") or []]}
              for i, s in enumerate(chart.get("series") or [])]
    return {"id": eid, "type": "chart", "role": role, "chart_kind": chart.get("kind") or "column",
            "x": x, "y": y, "width": w, "height": h,
            "categories": categories, "series": series,
            "show_values": len(series) == 1 and len(categories) <= int(policy.num("thresholds.chart_value_label_max")),
            "takeaway": chart.get("takeaway") or "",
            "derived": chart.get("derived", "author"),
            "layout_source": "plan"}
