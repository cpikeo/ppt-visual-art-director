#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""judgment.py · 设计判断协议的可执行部分（判断，不是排版）。

边界（与 SKILL「设计判断协议」 一一对应）：
  * 只产出「判断 + 依据 + 未决问题」，不产出坐标、不写 PPTX。
  * 每一条判断都带 source：
        author   —— 作者显式声明（压过推导）
        rule:<id>—— 由某条判断规则推出（带命中片段）
        absent   —— 判断不出，如实记为未断言/未决，不补编
  * 词表只做提示：命中产生一条待确认问题与依据片段，不产生风格、不直接决定构图。
  * 判断结果的确定性：同样 brief + 同样策略 → 字节级相同的 plan（便于回归与盲评）。

不做的事：不猜内容、不虚构数据关系、不把并列判成因果、不把相加 100% 判成同一分母。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from assets import resolve_source

PLAN_SCHEMA = "pvad-plan-v1"

# ── 文本形状的确定性识别（与词表无关的结构特征） ──────────────────────────
_NUM = re.compile(r"(?<![\d.])(?:\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)(?![\d.])")
_PERCENT = re.compile(r"(\d+(?:\.\d+)?)\s*[%％]")
_YEAR = re.compile(r"(?:19|20)\d{2}\s*(?:年|Q[1-4])?")
_PERIOD = re.compile(r"(?:19|20)\d{2}\s*年?|Q[1-4]|第[一二三四1-4]季度?|\d{1,2}\s*月(?:份)?|H[12]")
_SENT_SPLIT = re.compile(r"[。！？!?；;\n]+")
_CLAUSE_SPLIT = re.compile(r"[，,。；;：:！？!?]")
_VALUE_LINE = re.compile(
    r"^\s*([+\-−]?\d+(?:[.,]\d+)*\s*(?:%|％|‰|亿元|万元|亿美元|万美元|元|美元|公里|千米|米|吨|公斤|千克|"
    r"平方米|立方米|小时|分钟|秒|天|周|个月|年|倍|个百分点|台|人|户|家|件|次)?)\s*$")
_LABEL_VALUE = re.compile(r"^\s*([^：:，。,]{1,18})[：:]\s*(.+?)\s*$")
# 量的形状：数字 + 可选单位/百分号/比值，或「A/B」这类成对计数
_QUANTITY = re.compile(
    r"^[+\-−]?\d+(?:[.,]\d+)*(?:\s*(?:%|％|‰|亿元|万元|亿美元|万美元|元|美元|公里|千米|米|吨|公斤|千克|"
    r"平方米|立方米|小时|分钟|秒|天|周|个月|年|倍|个百分点|台|人|户|家|件|次))?$")
_SHARE_VALUE = re.compile(r"^\s*(?:占(?:比)?\s*)?(\d+(?:\.\d+)?\s*(?:%|％))\s*$")
# 行内写法：「经销 51%」「门店数 128 家」——标签与量在同一行
_INLINE_METRIC = re.compile(
    r"^\s*([^\s：:，。,0-9]{1,16})\s+(\d+(?:[.,]\d+)*\s*(?:%|％|‰|亿元|万元|亿美元|万美元|元|美元|公里|千米|"
    r"米|吨|公斤|千克|平方米|立方米|小时|分钟|秒|天|周|个月|年|倍|个百分点|台|人|户|家|件|次)?)\s*$")
_META_LINE = re.compile(r"^\s*(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{4}\s*年\s*\d{1,2}\s*月|内部|草案|机密|版本|v\d+(?:\.\d+)*)\s*$", re.I)
_SOURCE_LINE = re.compile(r"^\s*(?:资料)?(?:来源|数据来源|注|口径|取数|统计口径|Source)\s*[:：]", re.I)
_BULLET = re.compile(r"^\s*(?:[·•\-–—*+]|\d+[.、)]|\(\d+\))\s+")
_ORDER_MARK = re.compile(r"(?:第[一二三四五六七八九十\d]+步|步骤\s*\d|阶段\s*[一二三四1-4]|^\s*\d+[.、)]\s)")
# 措辞线索全部来自 policy/strategy.yaml 的 judgment.lexicons，不在这里硬编码：
#   claim_cue / assertion_cue / change_cue / imperative_cue
# 这样「判断依赖哪些词」是可审计、可消融的，而不是散落在代码里的常量。

# 判断输出的行动文本（可读的判断结论，不是决策逻辑；不参与分支）


@dataclass
class Trace:
    entries: list = field(default_factory=list)

    def add(self, field_name: str, value, source: str, why: str, span: str = ""):
        self.entries.append({"field": field_name, "value": value, "source": source,
                             "why": why, "span": span})
        return value


# ──────────────────────────────────────────────────────────────────────────
# 基础识别
# ──────────────────────────────────────────────────────────────────────────
def hits(text: str, words) -> list[str]:
    low = (text or "").lower()
    return [w for w in words if w and w.lower() in low]


def numerals(text: str) -> list[dict]:
    out = []
    for m in _NUM.finditer(text or ""):
        raw = m.group(0)
        value = float(raw.replace(",", ""))
        tail = text[m.end(): m.end() + 6]
        unit = ""
        pm = _PERCENT.match(raw)  # 已在 _NUM 之外单独处理百分号
        um = re.match(r"\s*(%|％|‰|亿元|万元|亿美元|万美元|元|美元|公里|千米|米|吨|公斤|千克|平方米|立方米|小时|分钟|秒|天|周|个月|年|倍|个百分点|台|人|户|家|件|次)", tail)
        if um:
            unit = um.group(1)
        for pm2 in _PERCENT.finditer(text):
            if pm2.start() == m.start():
                unit = "%"
        out.append({"raw": raw, "value": value, "unit": unit,
                    "year": bool(_YEAR.fullmatch(raw)) or (unit == "年" and 1900 < value < 2100)})
    return out


def numeric_density(text: str) -> float:
    chars = len(re.sub(r"\s", "", text or ""))
    if chars < 8:
        return 0.0
    count = len(numerals(text))
    return round(count * 100.0 / chars, 2)


def sentences(text: str) -> list[str]:
    return [s.strip() for s in _SENT_SPLIT.split(text or "") if s.strip()]


def parse_atoms(slide: dict, policy) -> dict:
    """把 content 解析成判断用的原子：指标、句子、条目、来源。

    确定性规则，不猜：数值行 + 紧随的短标签行 = 一个指标；「标签：值」直接成指标。
    """
    lines = [ln.strip() for ln in (slide.get("content") or "").split("\n") if ln.strip()]
    metrics: list[dict] = []
    body_lines: list[str] = []
    bullets: list[str] = []
    ordered: list[str] = []
    source = slide.get("source") or ""
    i = 0
    while i < len(lines):
        line = lines[i]
        if _SOURCE_LINE.match(line):
            source = source or _SOURCE_LINE.sub("", line).strip()
            i += 1
            continue
        if _META_LINE.match(line):
            i += 1
            continue
        lv = _LABEL_VALUE.match(line)
        if lv and _QUANTITY.match(lv.group(2).strip()) and len(lv.group(1)) <= 12:
            metrics.append({"value": lv.group(2).strip(), "unit": "", "label": lv.group(1).strip(),
                            "desc": "", "period": _period_of(lv.group(2))})
            i += 1
            continue
        vm = _VALUE_LINE.match(line) or _SHARE_VALUE.match(line)
        if vm:
            raw = (vm.group(1) if vm.re is _SHARE_VALUE else vm.group(1)).strip()
            value, unit = _split_value_unit(raw)
            if unit == "":
                sm = _SHARE_VALUE.match(line)
                if sm:
                    value, unit = sm.group(1).replace("%", "").replace("％", "").strip(), "%"
            label, desc = "", ""
            # 反序优先：「标签 / 占 N%」是中文里最常见的占比写法，标签在值之前
            is_share = bool(_SHARE_VALUE.match(line))
            if is_share and body_lines and _is_short_label(body_lines[-1]):
                label = body_lines.pop()
            # 正序：值行 + 标签行（+ 一行业务描述）
            if not label and i + 1 < len(lines):
                nxt = lines[i + 1]
                if _is_short_label(nxt):
                    label = nxt
                    i += 1
                    if i + 1 < len(lines) and _is_short_label(lines[i + 1]) and not _VALUE_LINE.match(lines[i + 1]):
                        desc = lines[i + 1]
                        i += 1
            if desc and _is_unit_token(desc):
                unit = unit or desc
                desc = ""
            metrics.append({"value": value, "unit": unit, "label": label, "desc": desc,
                            "period": _period_of(raw)})
            i += 1
            continue
        if _BULLET.match(line):
            bullets.append(_BULLET.sub("", line).strip())
            i += 1
            continue
        if _ORDER_MARK.match(line):
            ordered.append(_ORDER_MARK.sub("", _BULLET.sub("", line)).strip("：: "))
            i += 1
            continue
        im = _INLINE_METRIC.match(line)
        if im:
            value, unit = _split_value_unit(im.group(2).strip())
            metrics.append({"value": value, "unit": unit, "label": im.group(1).strip(),
                            "desc": "", "period": _period_of(line)})
            i += 1
            continue
        body_lines.append(line)
        i += 1

    for metric in slide.get("metrics") or []:
        metrics.append({"value": metric.get("value"), "unit": metric.get("unit") or "",
                        "label": metric.get("label") or "", "desc": metric.get("desc") or "",
                        "period": metric.get("period") or ""})

    # 词表在原文上匹配：归一化只服务排版，不能让「第一步」这类线索在切分时消失
    text = (slide.get("content") or "")
    return {
        "text_raw": text,
        "metrics": metrics,
        "body_lines": body_lines,
        "bullets": bullets,
        "steps": ordered + bullets,
        "sentences": sentences("\n".join(body_lines + bullets + ordered)),
        "source": source,
        "text": text,
        "period_labeled": len(set(_PERIOD.findall(text))) >= 2,
        "order_marked": bool(_ORDER_MARK.search(slide.get("content") or "")),
        "ordered_marked": bool(ordered),
    }


_UNIT_TOKEN = re.compile(
    r"^(?:%|％|‰|亿元|万元|亿美元|万美元|元|美元|公里|千米|米|吨|公斤|千克|平方米|立方米|小时|分钟|秒|天|周|"
    r"个月|年|倍|个百分点|台|人|户|家|件|次)$")


def _is_unit_token(text: str) -> bool:
    return bool(_UNIT_TOKEN.match((text or "").strip()))


def _is_short_label(line: str) -> bool:
    text = (line or "").strip()
    if not text or len(text) > 16:
        return False
    if _VALUE_LINE.match(text) or _SHARE_VALUE.match(text) or _SOURCE_LINE.match(text):
        return False
    if text.endswith(("。", "！", "？", "，", "；")):
        return False
    return not bool(re.search(r"[，；]", text))


def _split_value_unit(raw: str) -> tuple[str, str]:
    text = (raw or "").strip().replace("％", "%")
    m = re.match(r"^([+\-−]?\d+(?:[.,]\d+)*)\s*(.*)$", text)
    if not m:
        return raw, ""
    return m.group(1), m.group(2).strip()


def _period_of(text: str) -> str:
    m = _PERIOD.search(text or "")
    return m.group(0) if m else ""


# ──────────────────────────────────────────────────────────────────────────
# 判断
# ──────────────────────────────────────────────────────────────────────────
def judge_deck(brief: dict, policy, base_dir=None) -> dict:
    warnings: list[dict] = []
    trace = Trace()
    lex = policy.req("judgment.lexicons")
    probes = policy.req("judgment.probes")

    audience = brief.get("audience") or ""
    decision = brief.get("decision") or ""
    tension = brief.get("tension") or ""
    for name, value in (("audience", audience), ("decision", decision), ("tension", tension)):
        if not value:
            warnings.append({"code": "DECK_FIELD_EMPTY", "field": name,
                             "detail": f"{name} 未声明：收束页焦点与节奏判断会退化为默认值，请补上"})

    world = brief.get("visual_world") or ""
    if world:
        trace.add("deck.visual_world", world, "author", "作者给出的视觉世界原样生效")
    else:
        derived = _derive_world(brief, policy)
        world = derived["value"]
        trace.add("deck.visual_world", world, "rule:world-neutre", derived["why"])

    max_slides = (brief.get("constraints") or {}).get("max_slides")
    if max_slides and len(brief["slides"]) > max_slides:
        warnings.append({"code": "DECK_TOO_LONG", "field": "slides",
                         "detail": f"页数 {len(brief['slides'])} 超过声明上限 {max_slides}：先删页，不压字号"})

    pages = []
    for slide in brief["slides"]:
        pages.append(_judge_page(slide, brief, policy, lex, probes, len(pages), len(brief["slides"]),
                            base_dir=base_dir))

    # 全篇：高潮页与节奏复查（整稿层判断，不是逐页加总）
    climax = _pick_climax(pages)
    if climax is not None:
        pages[climax]["decision"]["emphasis"] = "lead"
        pages[climax]["decision"]["emphasis_why"] = "整套最强的一击落在这里：证据最实、且直接支撑收束页的决定。"
    for i, page in enumerate(pages):
        page["decision"]["rhythm"] = _rhythm(pages, i)

    structural = [p["understanding"]["evidence"]["value"] for p in pages]
    if not any(v in ("comparison", "series", "part_whole", "number") for v in structural):
        warnings.append({"code": "DECK_NO_EVIDENCE", "field": "slides",
                         "detail": "全稿没有可判定的量化证据：这是一份「纯观点稿」，请确认是否要引用数据"})

    return {
        "schema": PLAN_SCHEMA,
        "policy_sha256": policy.sha256,
        "brief_sha256": brief.get("_brief_sha256", ""),
        "deck": {
            "audience": audience, "decision": decision, "tension": tension,
            "visual_world": world,
            "coherence": {"climax": pages[climax]["id"] if climax is not None else None,
                          "why": "高潮页 = 证据最实且最贴近决定的那一页；其余页给它让路。"},
            "warnings": warnings,
        },
        "pages": pages,
        "trace": {"entries": trace.entries},
        "brief_notes": brief.get("_brief_notes", []),
    }


def _derive_world(brief: dict, policy) -> dict:
    """没有品牌与视觉声明时，用策略里的中性兜底——不按行业给「风格」。

    这段设计语言过去写在 Python 里（改口径要改代码、也不进策略指纹）；现在它在
    strategy.yaml 的 `judgment.default_world`，读不到就报错，不允许静默用另一套说法。
    """
    fallback = policy.req("judgment.default_world")
    return {"value": str(fallback["value"]), "why": str(fallback["why"])}


def _judge_page(slide: dict, brief: dict, policy, lex: dict, probes: dict,
                index: int, total: int, base_dir=None) -> dict:
    trace = Trace()
    atoms = parse_atoms(slide, policy)
    text_all = "\n".join([slide["title"], slide["content"] or ""])
    density = numeric_density(slide["content"] or "")
    nums = numerals(slide["content"] or "")
    non_year = [n for n in nums if not n["year"]]

    # 1. 主张
    claim = _judge_claim(slide, atoms, lex, trace)

    # 2. 证据形状
    evidence = _judge_evidence(slide, atoms, claim, lex, non_year, trace, probes)
    comparison_expression = (
        "ratio" if evidence.get("delta") and hits(claim.get("text") or "", lex.get("comparison_ratio") or [])
        else "difference" if evidence.get("delta") else "none"
    )

    # 3. 角色
    role = _judge_role(slide, index, total, evidence, claim, atoms, lex, trace)

    # 4. 构图
    composition = _judge_composition(slide, evidence, role, atoms, claim, trace)

    # 5. 焦点与密度
    focus = slide.get("focus") or _judge_focus(composition["value"], evidence, atoms, role, policy)
    trace.add("decision.focus", focus, "author" if slide.get("focus") else f"rule:focus-{composition}",
              "第一落点：本页唯一被记住的元素")
    density_target = slide.get("density") or _judge_density(composition, density, policy, atoms)
    trace.add("decision.density", density_target,
              "author" if slide.get("density") else "rule:density-budget",
              f"数字密度 {density}/100 字：密度由内容决定，不靠缩字号解决")

    # 6. 素材职责（图片不是装饰：先回答「拿掉它会损失什么」）
    media = _judge_media(slide, atoms, claim, evidence, composition, lex, probes, policy,
                         base_dir=base_dir)

    # 7. 未决问题
    questions = _questions(slide, atoms, claim, evidence, lex, probes, policy, brief, media)

    must_place = _must_place(composition, evidence, atoms)
    if media["decision"] != "none":
        must_place.append({"role": "media", "type": "image",
                           "note": f"{media['role']} / {media['function']}：{media['reason'] or '作者未写理由'}"})
    typography = {
        "claim_max_lines": policy.max_lines(_ladder_for(composition["value"], policy)),
        "claim_size_key": _ladder_for(composition["value"], policy),
        "reading_seconds": _reading(atoms.get("text_raw") or atoms["text"], policy),
        "bind_number_unit": bool(policy.req("typography.number_unit.bind")),
    }

    return {
        "id": slide["id"],
        "index": index,
        "title": slide["title"],
        "understanding": {
            "atoms": atoms,
            "numerals": nums,
            "non_year_numerals": len(non_year),
            "numeric_density": density,
            "evidence": evidence,
            "comparison": {"delta": evidence.get("delta", False),
                           "ends": evidence.get("ends", 0),
                           "same_unit": evidence.get("same_unit", False),
                           "expression": comparison_expression},
            "denominator": evidence.get("denominator", {"proven": False, "why": "无占比证据"}),
            "sentence_count": len(atoms["sentences"]),
        },
        "claim": claim,
        "decision": {
            "role": role["value"], "role_why": role["why"], "role_source": role["source"],
            "composition": composition["value"], "composition_why": composition["why"],
            "composition_alternatives": composition.get("alternatives", []),
            "focus": focus,
            "density": density_target,
            "media": media,
            "emphasis": "normal",
            "emphasis_why": "",
            "trace_ref": f"{slide['id']}",
        },
        "production": {
            "must_place": must_place,
            "must_not": list(policy.recipe(composition["value"]).get("must_not") or []),
            "typography": typography,
            "declared": {k: slide.get(k) for k in ("role", "evidence", "composition", "focus", "density")
                         if slide.get(k)},
        },
        "open_questions": questions,
        "trace": trace.entries,
    }


def _ladder_for(composition: str, policy) -> str:
    """字号档位属于配方的属性：写在 composition.yaml，不写在这里。"""
    return str(policy.recipe(composition)["ladder"])


def _reading(text: str, policy) -> float:
    from typography import reading_seconds
    return reading_seconds(text, policy)


def _judge_claim(slide: dict, atoms: dict, lex: dict, trace: Trace) -> dict:
    if slide.get("claim") or slide.get("insight"):
        value = slide.get("claim") or slide.get("insight")
        return trace.add("claim.text", value, "author",
                         "作者显式给出本页唯一主张：原话保留，不被正文证据行顶掉", value) and {
            "text": value, "source": "author", "form": _claim_form(value, lex), "why": "作者声明"}
    title = slide["title"].strip()
    if _is_assertion(title, lex):
        return {"text": title, "source": "author", "form": _claim_form(title, lex),
                "why": "标题本身是论断（含判断/比较/数值变化），占据主张位"}
    # 标题是话题标签：从正文找结论形状的句子，找不到就记 absent
    for sentence in atoms["sentences"]:
        if _claim_form(sentence, lex) != "description" and (
                hits(sentence, lex["claim_cue"]) or hits(sentence, lex["change_cue"])):
            return {"text": sentence, "source": "derived", "form": _claim_form(sentence, lex),
                    "why": "标题只是话题标签；正文这一句带结论形状，按判断协议升为主张"}
    return {"text": "", "source": "absent", "form": "none",
            "why": "标题是话题标签、正文没有结论句：本页主张未决，不编造一句「看起来像结论」的话"}


def _is_assertion(title: str, lex: dict) -> bool:
    title = (title or "").strip()
    if not title:
        return False
    if re.search(r"[，,。；;：:！!？?]", title):
        return True
    nums = [n for n in numerals(title) if not n["year"]]
    if nums:
        return True
    if hits(title, lex["claim_cue"]) or hits(title, lex["change_cue"]) or hits(title, lex["assertion_cue"]):
        return True
    return False


def _claim_form(text: str, lex: dict) -> str:
    if hits(text, lex["change_cue"]) or re.search(r"(高|低|超|低于|高于)\s*于?", text):
        return "comparative"
    if re.search(r"\d", text):
        return "numeric"
    if hits(text, lex["imperative_cue"]):
        return "imperative"
    if hits(text, lex["claim_cue"]):
        return "conclusion"
    return "description"


def _judge_evidence(slide, atoms, claim, lex, non_year, trace, probes) -> dict:
    if slide.get("evidence"):
        return {"value": slide["evidence"], "source": "author",
                "why": "作者显式声明证据形状", "spans": [], "delta": False, "ends": 0,
                "same_unit": False, "denominator": {"proven": False, "why": "作者声明未附分母证据"}}
    if slide.get("chart"):
        kind = slide["chart"]["kind"]
        shape = {"donut": "part_whole", "line": "series", "column": "series",
                 "bar": "comparison", "progress": "number"}.get(kind, "comparison")
        return {"value": shape, "source": "author:chart",
                "why": f"作者给出 {kind} 图数据：形状随数据，不另判",
                "spans": [], "delta": False,
                "ends": len(slide["chart"].get("categories") or []), "same_unit": True,
                "denominator": {"proven": shape == "part_whole", "why": "作者显式提供占比数据"}}

    text = " ".join([slide.get("title") or "", atoms.get("text_raw") or atoms["text"]])
    metrics = atoms["metrics"]
    units = {m.get("unit") for m in metrics if m.get("unit")}

    part_cue = hits(text, lex["part_whole"])
    denom_risk = hits(text, lex["denominator_risk"])
    comp_cue = hits(text, lex["comparison_end"])
    negated = hits(text, lex["comparison_negated"])
    series_cue = hits(text, lex["series"])
    seq_cue = hits(text, lex["sequence_strong"])
    struct_cue = hits(text, lex["structure"])

    # ① 占比：必须同时有占比声明与分母可信。
    # 「相加 100%」是分母风险信号——判断做在数值上（分项是否为百分比、合计是否≈100），
    # 而不是只靠「占比」两个字出现没出现。
    share_values = _share_values(metrics)
    sums_to_100 = bool(share_values) and len(share_values) >= 2 and abs(sum(share_values) - 100) <= 0.6
    # 占比形状只认数值证据：分项必须是百分比读数。
    # 「占比」两个字不构成分母——金额/数量分项即使作者写了「占比最高」，
    # 也不能据此画环（合计 3.2 亿元不是 100%，读者会误读为份额）。
    share_candidate = bool(share_values) and len(share_values) >= 2 and (part_cue or sums_to_100)
    part_note = ""
    if part_cue and metrics and not share_candidate:
        part_note = (f"；{part_cue[:2]}只是文字，没有百分比读数："
                     f"不按占比画图，按并排的分项量呈现（口径仍须作者确认）")
    if share_candidate:
        if denom_risk:
            return {"value": "comparison", "source": "rule:part-whole-demoted",
                    "why": f"出现占比形状（{part_cue[:2] or '百分比分项'}）但正文带分母风险信号"
                           f"（{denom_risk[:2]}）：相加为 100% 不证明同一分母，降级为并排指标，不画占比图",
                    "spans": part_cue[:2] + denom_risk[:2], "delta": False,
                    "ends": len(metrics), "same_unit": len(units) <= 1,
                    "denominator": {"proven": False, "why": "分母未证明；同一主体可同时满足多项"},
                    "demoted_from": "part_whole"}
        same = len(units) <= 1
        return {"value": "part_whole", "source": "rule:part-whole",
                "why": (f"占比声明（{part_cue[:2]}）+ 分项共享单位：可作为整体内的分项呈现"
                        if part_cue else "分项都是百分比但正文没有占比声明：按整体分项呈现，口径仍须确认")
                       + ("；仍须确认互斥性" if not same else ""),
                "spans": part_cue[:2], "delta": False, "ends": len(metrics), "same_unit": same,
                "denominator": {"proven": same, "why": "分项同单位且文本给出占比声明"}}

    # ② 两端比较：需要真实存在另一端，且不是「未超过」这类否定式
    if comp_cue and len(metrics) >= 2 and not negated:
        same = len(units) <= 1
        return {"value": "comparison", "source": "rule:comparison-two-end",
                "why": f"命中比较（{comp_cue[:2]}）且文中存在 {len(metrics)} 个可并置的量"
                       + ("、同一单位：差值成立" if same else "、单位不同：只能并排，不能相减") + part_note,
                "spans": comp_cue[:2], "delta": same and len(metrics) == 2, "ends": len(metrics),
                "same_unit": same,
                "denominator": {"proven": False, "why": "比较页不涉及整体分母"}}
    if negated:
        return {"value": "comparison", "source": "rule:comparison-negated",
                "why": f"命中否定式比较（{negated[:2]}）：量存在但胜负被否定，不画差值轴",
                "spans": negated[:2], "delta": False, "ends": len(metrics), "same_unit": len(units) <= 1,
                "denominator": {"proven": False, "why": "不涉及整体分母"}}

    # ③ 时间序列
    periods = {m.get("period") for m in metrics if m.get("period")}
    if series_cue or (len(units) == 1 and len(metrics) >= 3 and len(periods) >= 3):
        return {"value": "series", "source": "rule:series",
                "why": "同一量纲、带期间标记的多点读数：形状是曲线" if series_cue else
                       "同单位 + ≥3 个期间标记：按时间读数处理",
                "spans": (series_cue or [])[:2], "delta": False, "ends": len(metrics),
                "same_unit": True, "denominator": {"proven": False, "why": "不涉及整体分母"}}

    # ④ 先后顺序
    # 两种证据各自成立：作者用 1)/2)/3) 显式编号（结构证据），或正文出现阶段词（语义证据）。
    # 只认阶段词会让「四个节点」这种作者已经编号的页掉进结构图——顺序是作者给的，
    # 不该因为少了一个「首先」就丢掉（案例 A s14 的回归）。
    # 用 order_marked（对原文跑正则）而不是 ordered_marked（解析出的 ordered 列表）：
    # 「1) xxx」会先被 _BULLET 收进 bullets，ordered 列表反而是空的。
    numbered = bool(atoms.get("order_marked")) and len(atoms.get("steps") or []) >= 3
    if (seq_cue or numbered) and (atoms["steps"] or len(metrics) >= 2):
        strong = bool(atoms["steps"]) and atoms["ordered_marked"] if "ordered_marked" in atoms else bool(atoms["steps"])
        if numbered and not seq_cue:
            why = "作者用序号（1/2/3）逐条给出先后：顺序由作者给出，不另判"
        elif strong:
            why = f"有序条目 + 阶段词（{seq_cue[:2]}）：顺序由文字明确给出"
        else:
            why = f"命中阶段词（{seq_cue[:2]}）但条目没有编号：按阶段处理，但不得自造序号"
        return {"value": "sequence", "source": "rule:sequence-strong" if (strong or numbered) else "rule:sequence-weak",
                "why": why,
                "spans": seq_cue[:2], "delta": False,
                "ends": len(atoms["steps"]) or len(metrics), "same_unit": False,
                "denominator": {"proven": False, "why": "不涉及整体分母"},
                "order_marked": strong}

    # ⑤ 结构/归属
    if struct_cue and (atoms["bullets"] or len(atoms["sentences"]) >= 2):
        return {"value": "structure", "source": "rule:structure",
                "why": f"命中结构词（{struct_cue[:2]}）：先按分组与对齐组织，只有明确写出归属才加层级",
                "spans": struct_cue[:2], "delta": False, "ends": len(atoms["bullets"]),
                "same_unit": False, "denominator": {"proven": False, "why": "不涉及整体分母"}}

    # ⑥ 并排指标（多指标、无差值）
    if len(metrics) >= 2:
        return {"value": "comparison", "source": "rule:metric-row",
                "why": "多个量并排但没有可相减的两端：只并排，不造差值轴" + part_note,
                "spans": [], "delta": False, "ends": len(metrics), "same_unit": len(units) <= 1,
                "denominator": {"proven": False, "why": "不涉及整体分母"}}

    # ⑦ 单一量
    if len(metrics) == 1:
        return {"value": "number", "source": "rule:single-number",
                "why": "唯一量化证据：数字本身就是焦点", "spans": [],
                "delta": False, "ends": 1, "same_unit": True,
                "denominator": {"proven": False, "why": "不涉及整体分母"}}

    # ⑧ 纯观点 / 无证据
    if claim["source"] != "absent" and atoms["sentences"]:
        return {"value": "proposition", "source": "rule:proposition",
                "why": "有主张、没有量化证据：这里是观点页，不硬凑图表",
                "spans": [], "delta": False, "ends": 0, "same_unit": False,
                "denominator": {"proven": False, "why": "无占比证据"}}
    return {"value": "none", "source": "rule:no-evidence",
            "why": "既无可判定的证据形状，也没有可复述的主张：本页证据未决",
            "spans": [], "delta": False, "ends": 0, "same_unit": False,
            "denominator": {"proven": False, "why": "无占比证据"}}


def _share_values(metrics: list[dict]) -> list[float]:
    """分项里所有「百分比」读数（%）——用于判断它们是否真的构成一个整体。"""
    values = []
    for m in metrics:
        unit = str(m.get("unit") or "").strip()
        raw = str(m.get("value") or "").replace(",", "").replace("％", "%").strip()
        if unit in ("%", "％", "") and raw.endswith("%"):
            raw = raw[:-1]
        if unit not in ("%", "％") and not str(m.get("value", "")).strip().endswith(("%", "％")):
            continue
        try:
            values.append(float(raw.rstrip("%").strip()))
        except ValueError:
            continue
    return values


def _judge_role(slide, index: int, total: int, evidence: dict, claim: dict, atoms: dict,
                lex: dict, trace: Trace) -> dict:
    if slide.get("role"):
        return {"value": slide["role"], "source": "author", "why": "作者显式声明页面角色"}
    kind = evidence["value"]
    ask = hits(" ".join([slide.get("title") or "", atoms["text"]]), lex["ask"])
    if index == 0:
        return {"value": "establish", "source": "rule:cover", "why": "首屏建立整稿主张与语境"}
    if index == total - 1 and (ask or claim.get("form") == "imperative"):
        return {"value": "summarize", "source": "rule:closing-ask",
                "why": "末页承载决定：观众看完要做的动作在这里收束"}
    if kind == "comparison":
        return {"value": "compare", "source": "rule:role-compare", "why": "证据形态是比较：页面任务是让胜负可读"}
    if kind in ("series", "part_whole"):
        return {"value": "prove", "source": "rule:role-prove", "why": "数据形状承担证明：页面任务是把结论坐实"}
    if kind in ("sequence", "structure"):
        return {"value": "explain", "source": "rule:role-explain", "why": "内容要求解释关系或过程"}
    if claim.get("form") == "imperative" or ask:
        return {"value": "persuade", "source": "rule:role-persuade", "why": "主张是行动要求：页面任务是推动决定"}
    if index == total - 1:
        return {"value": "summarize", "source": "rule:closing", "why": "末页收束，不再引入新证据"}
    return {"value": "explain", "source": "rule:role-default", "why": "无更强信号：本页承担解释"}


def _judge_composition(slide, evidence: dict, role: dict, atoms: dict, claim: dict, trace: Trace) -> dict:
    if slide.get("composition"):
        return {"value": slide["composition"], "source": "author",
                "why": "作者显式声明构图（原样生效）", "alternatives": []}
    if slide.get("asset"):
        function = (slide["asset"].get("function") or "context")
        if function in ("hero", "context", "proof"):
            return {"value": "image_narrative", "source": "rule:media-narrative",
                    "why": f"作者声明了 {function} 素材：视觉主体占主区，文字退到安全区，"
                           "先验证可读性再谈构图",
                    "alternatives": [{"composition": "statement",
                                      "rejected_because": "有主体素材时纯文字页浪费了唯一的视觉证据"}]}
        return {"value": "image_narrative", "source": "rule:media-frame",
                "why": "作者声明全幅背景素材：文字必须落在安全区，且实测对比度达标", "alternatives": []}
    kind = evidence["value"]
    alternatives = []
    if kind == "part_whole":
        value = "part_whole"
        why = "分母与互斥性成立时，占比形状比并排数字更快解释构成"
    elif kind == "comparison" and evidence.get("delta"):
        value = "two_end"
        why = "只有两端且同单位：差值本身是结论，把两端拉开"
        alternatives = [{"composition": "metric_row", "rejected_because": "差值成立时并排会削弱胜负"}]
    elif kind == "comparison":
        value = "metric_row"
        why = "多指标并排但不构成可相减的两端：并排呈现量级，不造差值轴"
        if evidence.get("demoted_from"):
            why += "（由占比图降级：分母未证明）"
            alternatives = [{"composition": "part_whole", "rejected_because": evidence.get("why", "")}]
    elif kind == "series" and slide.get("chart"):
        value = "series_chart"
        why = "作者提供的时间读数：图占主区，结论句在上方指路"
    elif kind == "series":
        value = "series_chart"
        why = "同量纲的时间读数：曲线承载趋势，结论句在上方指路"
    elif kind == "sequence":
        value = "sequence_steps"
        why = "阶段有先后：横向推进让「下一步」成为阅读终点"
    elif kind == "structure":
        value = "structure_map"
        why = "结构优先：先分组与对齐，容器与层级只在明确给归属时出现"
    elif kind in ("proposition", "none"):
        if role["value"] == "summarize":
            value = "statement"
            why = "收束页只留一个决定：把要观众做的事写成一句话，其余退场"
        elif len(atoms["sentences"]) >= 3 and role["value"] != "establish":
            value = "editorial_text"
            why = "观点需要被读完：分栏正文形成节奏，避免被扫成碎片"
        else:
            value = "statement"
            why = "一页一个第一落点：大字主张压住版面，证据退到次要位置"
    else:  # number
        value = "statement"
        why = "唯一量化证据：让数字成为视觉主角，其余退后"
    return {"value": value, "source": "rule:composition", "why": why, "alternatives": alternatives}


def _judge_media(slide: dict, atoms: dict, claim: dict, evidence: dict, composition: dict,
                 lex: dict, probes: dict, policy, base_dir=None) -> dict:
    """素材判断：要不要图、图承担什么、拿掉会损失什么。

    纪律（SKILL「素材艺术指导与合规」）：
      * 出图由作者声明决定；词表只提示，未声明就 decision=none，绝不自动出图。
      * 声明了就必须说明职责（role + function）；说不出理由的装饰性素材是待确认问题。
      * 生成图不充当事实证词：proof 职能要求作者给出来源，否则记未断言。
    """
    roles = policy.enum("media_role")
    functions = policy.enum("media_function")
    declared = slide.get("asset") or None
    if declared:
        role = (declared.get("function") or "context").strip()
        function = role if role in functions else "context"
        role_name = "background" if function in ("frame", "emotion") else "subject"
        if role_name not in roles:
            role_name = "subject"
        reason = (declared.get("reason") or "").strip()
        # 相对路径相对 brief 所在目录解析——用 cwd 判断会把「文件其实在」误判成要出图
        decision = "existing" if resolve_source(declared.get("path") or "", base_dir) else "required"
        why = (f"作者声明素材（function={function}）：原样生效；"
               + ("素材文件已存在，按既有素材使用" if decision == "existing" else "按出图需求登记，不替外部生成器出图"))
        return {"decision": decision, "role": role_name, "function": function,
                "reason": reason, "source": "author", "why": why,
                "path": declared.get("path") or "", "fit": declared.get("fit") or "",
                "safe_area": declared.get("safe_area") or None,
                "text_color": declared.get("text_color") or "",
                "subject": (declared.get("subject") or "").strip(),
                "avoid": list(declared.get("avoid") or [])}

    words = hits(" ".join([slide.get("title") or "", atoms.get("text_raw") or atoms.get("text") or ""]),
                 lex.get("visual_subject") or [])
    if words:
        # 词表只提示：不产生图片，只产生一条待确认问题（与判断陷阱纪律一致）
        return {"decision": "none", "role": "", "function": "", "reason": "",
                "source": "rule:visual-subject-hint",
                "why": f"内容涉及可被看见的对象（{words[:3]}），但作者没有声明素材："
                       "按「未声明不出图」处理，仅记一条待确认问题",
                "path": "", "fit": "", "safe_area": None, "text_color": "",
                "subject": "", "avoid": [], "hint": words[:3]}
    return {"decision": "none", "role": "", "function": "", "reason": "",
            "source": "rule:no-media",
            "why": "作者未声明素材，内容也不依赖「看见」：纯排版页，不出图",
            "path": "", "fit": "", "safe_area": None, "text_color": "",
            "subject": "", "avoid": []}


def _judge_focus(composition: str, evidence: dict, atoms: dict, role: dict, policy) -> str:
    """第一落点：配方给的默认值（策略）→ 只有「收束页无量化证据」这一条判断能改它。"""
    if role["value"] == "summarize" and evidence["value"] in ("proposition", "none"):
        return "decision"
    if evidence["value"] == "number":
        return "metric"
    return str(policy.recipe(composition)["focus"])


def _judge_density(composition: str, density: float, policy, atoms: dict) -> str:
    high = float(policy.num("thresholds.numeric_density_high"))
    low = float(policy.num("thresholds.numeric_density_low"))
    if composition == "editorial_text":
        return "balanced"
    if density >= high or len(atoms["metrics"]) >= 4:
        return "dense"
    if density <= low and len(atoms["metrics"]) <= 1:
        return "sparse"
    return "balanced"


def _must_place(composition: str, evidence: dict, atoms: dict) -> list[dict]:
    base = [{"role": "claim", "type": "text", "note": "主张必须落在第一落点"}]
    if composition == "metric_row":
        base.append({"role": "metric_cells", "type": "kpi",
                     "note": f"{len(atoms['metrics'])} 个量各自带标签与单位，互不换算"})
    elif composition == "two_end":
        base.append({"role": "two_ends", "type": "text", "note": "两端 + 差值，各自标注口径"})
    elif composition == "series_chart":
        base.append({"role": "chart", "type": "chart", "note": "同一量纲的时间读数；类别 ≤ 上限"})
    elif composition == "part_whole":
        base.append({"role": "chart", "type": "chart", "note": "占比图必须与口径说明同页"})
    elif composition == "sequence_steps":
        base.append({"role": "steps", "type": "steps", "note": "顺序由文字给出，序号跟随"})
    elif composition == "structure_map":
        base.append({"role": "groups", "type": "groups", "note": "分组与对齐；无依据不画连线"})
    elif composition == "statement" and evidence["value"] == "number":
        base.append({"role": "hero_value", "type": "text", "note": "数字作焦点时必须带单位与口径"})
    if atoms["source"]:
        base.append({"role": "source", "type": "text", "note": "来源与期间同页可读"})
    return base


def _questions(slide, atoms, claim, evidence, lex, probes, policy, brief, media=None) -> list[dict]:
    text = " ".join([slide.get("title") or "", atoms.get("text_raw") or atoms["text"]])
    out: list[dict] = []
    for name, probe in probes.items():
        words = lex.get(name) or []
        if hits(text, words):
            # 只有命中确实影响本页判断的词表才追问，避免把每个词都变成噪声
            relevant = {
                "part_whole": evidence["value"] in ("part_whole", "comparison"),
                "comparison_end": evidence["value"] == "comparison",
                "structure": evidence["value"] == "structure",
                "conditional": bool(hits(text, lex["conditional"])),
                "causal_risk": bool(hits(text, lex["causal_risk"])),
                "correlation_cue": bool(hits(text, lex["correlation_cue"])),
            }.get(name, True)
            if relevant:
                # 素材提示使用稳定代码 ASSET_HINT；不要再在 media 分支生成同义副本。
                code = "ASSET_HINT" if name == "visual_subject" else f"PROBE_{name.upper()}"
                out.append({"code": code, "question": probe,
                            "why": f"命中词表 {name}: {hits(text, words)[:3]}"})
    if claim["source"] == "absent":
        out.append({"code": "CLAIM_ABSENT",
                    "question": "这一页要让观众记住哪一句话？标题目前只是话题标签，正文没有结论句。",
                    "why": "主张未决：不编造主张，交由作者裁决"})
    numeric = [m for m in atoms["metrics"] if m.get("value") not in (None, "")]
    missing = [k for k in policy.req("verification.release_requires")
               if not ((slide.get(k) or "").strip() or (brief.get(k) or "").strip())]
    if numeric and missing:
        out.append({"code": "PROVENANCE_MISSING",
                    "question": f"这一页的数字缺 {missing}：请在页内或 brief 级补齐，否则 release 会被拦。",
                    "why": "数字必须与来源、单位、期间、口径同源"})
    if media:
        if media.get("hint") and media["decision"] == "none" and not any(
                q["code"] == "ASSET_HINT" for q in out):
            out.append({"code": "ASSET_HINT", "question": probes.get("visual_subject", ""),
                        "why": f"命中 visual_subject: {media['hint']}（只提示，不自动出图）"})
        if media["decision"] != "none" and not media.get("reason"):
            out.append({"code": "ASSET_REASON_MISSING",
                        "question": "这张素材替观众理解了什么？拿掉它，理解会在哪里下降？",
                        "why": "素材职责未写：没有职责的素材就是装饰"})
        if media.get("function") == "proof" and not (slide.get("source") or brief.get("source")):
            out.append({"code": "ASSET_PROOF_UNSOURCED",
                        "question": "这张图作为证据使用，它的来源与拍摄/制作口径是什么？生成图不能充当事实证词。",
                        "why": "proof 职能要求可核来源"})
        if media["decision"] != "none" and evidence["value"] in ("series", "part_whole", "comparison"):
            out.append({"code": "ASSET_VS_EVIDENCE",
                        "question": "这页已经有量化证据，图片是在承担构图还是只是填空？",
                        "why": "数字页的主要视觉应是数据形状，图片不应抢证据"})
    if evidence["value"] == "none" and not atoms["sentences"]:
        out.append({"code": "EMPTY_PAGE", "question": "这一页没有证据也没有观点，是否应该删除？",
                    "why": "空页占用了整套节奏，却没有信息增量"})
    return out


def _rhythm(pages: list, index: int) -> dict:
    page = pages[index]
    if index == 0:
        return {"move": "open", "previous": None, "why": "首屏建立语境，不承担证明"}
    prev = pages[index - 1]
    cur_c, prev_c = page["decision"]["composition"], prev["decision"]["composition"]
    cur_role, prev_role = page["decision"]["role"], prev["decision"]["role"]
    cur_d, prev_d = page["decision"]["density"], prev["decision"]["density"]
    if index == len(pages) - 1:
        return {"move": "close", "previous": prev["id"], "why": "末页收束到决定，视觉分量让位给一句话"}
    if prev_d == "dense" and cur_d == "sparse":
        return {"move": "breathe", "previous": prev["id"], "why": "前面刚给过密度，这一页留呼吸，避免连续压迫"}
    if cur_c == prev_c and cur_role == prev_role:
        return {"move": "hold", "previous": prev["id"],
                "why": "同一论证的连续展开：保持同一构图骨架，只换证据，便于横向比较"}
    if cur_role != prev_role:
        return {"move": "turn", "previous": prev["id"], "why": "页面任务发生变化，阅读路径跟着转向"}
    return {"move": "build", "previous": prev["id"], "why": "同一任务内的推进：构图变化须有叙事理由"}


def _pick_climax(pages: list):
    """高潮页：证据最实 + 最贴近决定。确定性打分，不随机。"""
    best, best_score = None, -1.0
    for i, page in enumerate(pages):
        ev = page["understanding"]["evidence"]["value"]
        if ev not in ("comparison", "series", "part_whole", "number"):
            continue
        score = {"comparison": 2.0, "part_whole": 1.8, "series": 1.6, "number": 1.2}[ev]
        score += 0.6 if page["understanding"]["comparison"]["delta"] else 0.0
        score += min(1.0, page["understanding"]["numeric_density"] / 3.0)
        score += 0.4 if page["claim"]["source"] != "absent" else 0.0
        score -= abs(i - (len(pages) - 2)) * 0.05   # 靠近收束页略加分
        if score > best_score:
            best, best_score = i, score
    return best
