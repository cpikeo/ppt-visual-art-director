#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""prompt.py · 把「这一页需要什么图」翻译成可交付的生产说明（确定性，不做判断）。

四层边界里的位置：

  判断（要不要图、图承担什么、缺了就丢什么）  → judgment / SKILL「素材艺术指导与合规」
  语言（用哪些词、什么顺序、反向词清单）      → policy/prompts.yaml
  翻译（本文）：把已结算的结论拼成 prompt/negative/ratio/target_px  → 无决策权
  契约与证据：schema/prompts.schema.json + tests + evals/prompt_eval.py

本模块**不**决定：该不该出图、用什么介质风格、留白好不好看。
它只回答四个可被摄影师或图像模型直接执行的问题：拍什么、为什么拍、给谁让位、交付多大。

两条从上游实测里学到的硬纪律（写在这里就不用每次重犯）：
  1. 提示词里**不写坐标**。写了 `x 6% y 8% w 34% h 78%`，模型会把留白画成硬边平色面板；
     安全区只以「哪一侧安静、什么调子」表达，矩形属于 QC。
  2. 反向词只排除**生产失效**（AI 伪影、意外边框、意外平板、意外接缝），
     不排除风格与题材——「廉价感」是评审判断，不是可配置策略。
"""
from __future__ import annotations

import hashlib
import json
import math
import re
from math import gcd

SCHEMA = "pvad-prompt-v1"

# 提示词里出现这些字符组合就说明有人在写坐标（百分比 / 千分位 / 坐标对）
_COORD_HINTS = re.compile(r"\d\s*%|\d+\s*[,，]\s*\d+\s*(?:px|像素)|x\s*[:=]\s*0?\.\d")


class PromptError(ValueError):
    """策略片段缺失或引用不一致（这是引擎缺陷，不是稿子问题）。"""


def _phrases(policy) -> dict:
    prompts = getattr(policy, "prompts", None)
    if not isinstance(prompts, dict) or not prompts:
        raise PromptError("策略缺少 prompts.yaml（提示词层不可静默降级）")
    return prompts


def _dedup(items) -> list[str]:
    seen, out = set(), []
    for item in items:
        text = str(item or "").strip()
        key = text.lower().rstrip(".")
        if text and key not in seen:
            seen.add(key)
            out.append(text)
    return out


def quiet_side(rect: dict | None) -> str | None:
    """安全区落在哪一侧——只看矩形中心偏向哪个方向，不猜作者意图。"""
    if not rect:
        return None
    try:
        x = float(rect.get("x", 0.0))
        y = float(rect.get("y", 0.0))
        w = float(rect.get("width", 0.0))
        h = float(rect.get("height", 0.0))
    except (TypeError, ValueError):
        return None
    if w <= 0 or h <= 0:
        return None
    cx, cy = x + w / 2.0, y + h / 2.0
    # 偏离中心的量：横纵取大者；都不足 1/8 视为「中间」（整块留白需要通篇安静）
    dx, dy = cx - 0.5, cy - 0.5
    if abs(dx) < 0.125 and abs(dy) < 0.125:
        return "center"
    return ("left" if dx < 0 else "right") if abs(dx) >= abs(dy) else ("top" if dy < 0 else "bottom")


def ratio_of(width: float, height: float, policy=None) -> str:
    """像素框 → 出图端能照做的画幅。

    先吸附到策略里的常用比例（容差内），吸附不上才给最简比例——
    496×624 的最简比例是 31:39，没人会按 31:39 出图，但 4:5 就是标准画幅。
    """
    try:
        w, h = float(width), float(height)
    except (TypeError, ValueError):
        return "1:1"
    if w <= 0 or h <= 0:
        return "1:1"
    presets = (getattr(policy, "prompts", {}) or {}).get("ratio") or {}
    target = w / h
    tolerance = float(presets.get("tolerance") or 0)
    best = None
    for pair in presets.get("presets") or []:
        if not isinstance(pair, (list, tuple)) or len(pair) != 2 or not pair[0] or not pair[1]:
            continue
        value = float(pair[0]) / float(pair[1])
        gap = abs(value - target) / target
        if (best is None or gap < best[0]) and gap <= tolerance:
            best = (gap, f"{int(pair[0])}:{int(pair[1])}")
    if best:
        return best[1]
    iw, ih = int(round(w)), int(round(h))
    g = gcd(iw, ih) or 1
    return f"{iw // g}:{ih // g}"


def target_px(width: float, height: float, policy) -> dict:
    """最小源图尺寸：放置尺寸 × oversample，且长边不低于策略下限（放大必糊）。"""
    over = float(policy.req("assets.oversample"))
    floor = int(policy.req("assets.min_long_side_px"))
    w = int(math.ceil(float(width) * over))
    h = int(math.ceil(float(height) * over))
    long_side = max(w, h)
    if long_side < floor:
        scale = floor / long_side if long_side else 1.0
        w, h = int(math.ceil(w * scale)), int(math.ceil(h * scale))
    return {"width": w, "height": h}


def fingerprint(brief: dict) -> str:
    """同一份说明 → 同一指纹：作者能判断「这张图要不要重出」。"""
    payload = {k: brief.get(k) for k in ("prompt", "negative", "ratio", "target_px")}
    blob = json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return "prompt-" + hashlib.sha256(blob).hexdigest()[:16]


def build(*, policy, function: str, role: str, subject: str = "", reason: str = "",
          box: tuple | list, safe_area: dict | None = None, text_color: str = "",
          avoid=(), crop_loss: float | None = None) -> dict:
    """拼装生产说明。所有输入都已由上游判定过：本函数不再改判任何一项。

    返回 dict(prompt, negative, ratio, target_px, safe_area, quiet_side,
              clauses[], fingerprint, issues[])。clauses 让审计能回答
    「这句话是哪条策略生成的」，而不是对着一整段字符串猜。
    """
    table = _phrases(policy)
    phrases = table["phrases"]
    order = list(table["clause_order"])
    width, height = float(box[2]), float(box[3])
    ratio = ratio_of(width, height, policy)
    area = dict(safe_area) if isinstance(safe_area, dict) and safe_area else {}
    side = quiet_side(area)
    issues: list[dict] = []

    if role not in phrases["role"]:
        issues.append({"code": "PROMPT_ROLE_UNKNOWN",
                       "detail": f"role={role!r} 不在策略词表 {sorted(phrases['role'])}"})
        role = "support"
    if function not in phrases["function"]:
        issues.append({"code": "PROMPT_FUNCTION_UNKNOWN",
                       "detail": f"function={function!r} 不在策略词表 {sorted(phrases['function'])}"})
        function = "context"
    if not str(subject or "").strip():
        # 说不清画面里出现什么，就没有可执行的生产说明——这是缺口，不是可选项
        issues.append({"code": "PROMPT_SUBJECT_MISSING",
                       "detail": "brief 没有声明 asset.subject：提示词无法回答「拍什么」"})
    if area and text_color not in ("dark", "light"):
        issues.append({"code": "PROMPT_TONE_MISSING",
                       "detail": "声明了 safe_area 但没写 text_color=dark|light：留白调子无法指定"})
    if str(function) == "proof" and str(role) == "background":
        issues.append({"code": "PROMPT_ROLE_CONFLICT",
                       "detail": "function=proof 需要可被检视的主体，role=background 承诺没有主体：二者不能同时成立"})

    slots: dict[str, list[str]] = {}
    slots["role"] = list(phrases["role"][role])
    slots["subject"] = [str(subject).strip()] if str(subject or "").strip() else []
    slots["function"] = [phrases["function"][function]]
    slots["framing"] = [phrases["framing"].format(ratio=ratio)]
    if side:
        slots["quiet_area"] = [phrases["quiet_area"][side], phrases["quiet_no_panel"]]
        if text_color in ("dark", "light"):
            slots["tone"] = [phrases["tone"][text_color]]
    if crop_loss is not None and crop_loss > float(policy.req("assets.crop_requires_reason_above")):
        slots["crop_safety"] = [phrases["crop_safety"]]
    slots["tail"] = list(phrases.get("tail") or table.get("tail") or [])
    slots["ratio"] = [phrases["ratio_label"].format(ratio=ratio)]

    clauses: list[dict] = []
    for slot in order:
        for text in _dedup(slots.get(slot) or []):
            clauses.append({"slot": slot, "text": text})
    prompt = str(table.get("clause_join") or "; ").join(c["text"] for c in clauses)

    if _COORD_HINTS.search(prompt):
        # 坐标进提示词会把留白画成面板；这条守卫挡的是「以后有人好心加回去」
        issues.append({"code": "PROMPT_COORDINATES",
                       "detail": "提示词里出现了像素/百分比坐标：留白只描述哪一侧安静，矩形属于 QC"})
    limit = int(table["limits"]["max_chars"])
    if len(prompt) > limit:
        issues.append({"code": "PROMPT_TOO_LONG",
                       "detail": f"提示词 {len(prompt)} 字符超过上限 {limit}：在堆砌，砍掉形容词"})

    avoid_list = [avoid] if isinstance(avoid, str) else list(avoid or ())
    negatives = (list(table["negatives"]["base"])
                 + list(table["negatives"]["production"])
                 + list((table["negatives"]["by_function"] or {}).get(function) or [])
                 + [str(x).strip() for x in avoid_list])
    negatives = [n for n in _dedup(negatives) if n]
    cap = int(table["limits"]["max_negative_terms"])
    if len(negatives) > cap:
        issues.append({"code": "PROMPT_NEGATIVE_BLOAT",
                       "detail": f"反向词 {len(negatives)} 条超过上限 {cap}：互相稀释，删掉同义项"})
    join = str(table["negatives"]["join"])
    negative = join.join(
        term if term.lower().startswith("no ") else f"no {term}" for term in negatives)

    brief = {
        "schema": SCHEMA,
        "function": function,
        "role": role,
        "reason": str(reason or "").strip(),
        "prompt": prompt,
        "negative": negative,
        "clauses": clauses,
        "ratio": ratio,
        "target_px": target_px(width, height, policy),
        "safe_area": area or None,
        "quiet_side": side,
        "text_color": text_color if text_color in ("dark", "light") else None,
        "issues": issues,
    }
    brief["fingerprint"] = fingerprint(brief)
    return brief


# ── 生产清单：给外部生成器 / 摄影师的一页纸 ────────────────────────────────
def manifest(deck: dict, policy) -> dict:
    """从 deck 里取出「还需要生产的图」，连同生产说明一起序列化。

    只列需要的：文件已在的图走 QC 与对账，不再出提示词（提示词是给生产的说明，
    不是给审阅的作文）。
    """
    entries: list[dict] = []
    for slide in deck.get("slides") or []:
        for element in slide.get("elements") or []:
            if element.get("type") != "image":
                continue
            status = "to-produce" if element.get("_missing") else ("broken" if element.get("_undecodable") else "ready")
            brief = element.get("brief") or {}
            entries.append({
                "id": element.get("id"),
                "slide": slide.get("id"),
                "status": status,
                "expected_path": element.get("source") or "",
                "function": element.get("function") or brief.get("function"),
                "reason": element.get("reason") or "",
                "ratio": brief.get("ratio"),
                "target_px": brief.get("target_px"),
                "safe_area": element.get("safe_area") or None,
                "prompt": brief.get("prompt"),
                "negative": brief.get("negative"),
                "fingerprint": brief.get("fingerprint"),
                "issues": brief.get("issues") or [],
                "note": element.get("_undecodable") or ("文件缺失，按本说明生产后重建" if element.get("_missing") else ""),
            })
    to_produce = [e for e in entries if e["status"] != "ready"]
    return {"schema": "pvad-media-brief-v1",
            "deck_sha256": deck.get("meta", {}).get("deck_sha256") or "",
            "policy_sha256": deck.get("meta", {}).get("policy_sha256") or "",
            "images": len(entries),
            "to_produce": len(to_produce),
            "entries": entries,
            "disclaimer": "提示词只规定「拍什么/为什么拍/给谁让位/交付多大」，"
                          "不保证出图好看；图片质量由 QC 硬门与人工评审分别判定（SKILL「素材艺术指导与合规」「证据层级」）。"}
