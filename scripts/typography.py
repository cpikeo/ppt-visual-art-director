#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""typography.py · 中文排版的真实度量与断行（确定性，无审美判断）。

职责边界：
  * 只回答「这段文字在给定字号/框宽下会有几行、墨迹多高、哪里不能断」。
  * 不决定字号、不选字体家族、不判断该删哪句 —— 那是判断协议与作者的职责。
  * 度量优先用真实字体文件（PIL 字形前进宽度）；没有字体文件时退化为估算，
    并在返回值里标 degraded，调用方必须把它写进报告（不得把估算当测量）。

关键纪律：
  * 数字与单位绑定（U+00A0），「186 公里」不在单位前断开。
  * 中文禁则：行首禁标点、行尾禁开引号；用「推」而不是「拉」，避免为避标点写出超框行。
  * 墨迹跨度按锚点计算，不用框高冒充墨迹高。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from functools import lru_cache
from pathlib import Path

NBSP = "\u00a0"
_LATIN_WORD = re.compile(r"[A-Za-z0-9][A-Za-z0-9.,%+\-_/:'′″]*")
_NUMBER = re.compile(r"[+\-−]?\d+(?:[.,]\d+)*")
_CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef]")
_HAS_CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
_BULLET_PREFIX = re.compile(r"^\s*(?:[·•\-–—*+]|\d+[.、)]|\(\d+\))\s*")


class FontUnavailable(RuntimeError):
    pass


@dataclass(frozen=True)
class Metrics:
    """一次断行的结果。lines 是最终行；degraded=True 表示没有真实字体度量。"""
    lines: tuple
    widths: tuple
    size: float
    line_height: float
    degraded: bool = False
    truncated: bool = False

    @property
    def count(self) -> int:
        return len(self.lines)


@dataclass
class WrapIssue:
    kind: str            # orphan_line | forbidden_line_start | forbidden_line_end | overflow
    line_index: int
    detail: str
    severity: str = "warn"   # warn | error

    def as_dict(self) -> dict:
        return asdict(self)


class FontBook:
    """按策略解析字体文件；同一进程内缓存。"""

    def __init__(self, policy):
        self.policy = policy
        self.degraded = False
        self.resolved: dict[str, str | None] = {}
        for role in ("cn", "cn_bold", "latin"):
            self.resolved[role] = self._resolve(role)
        if self.resolved["cn"] is None and self.resolved["latin"] is None:
            self.degraded = True

    def _resolve(self, role: str) -> str | None:
        for candidate in self.policy.req(f"typography.font_files.{role}") or []:
            path = Path(candidate).expanduser()
            if path.is_file():
                return str(path)
        return None

    @lru_cache(maxsize=256)
    def _font(self, role: str, size: int, bold: bool):
        from PIL import ImageFont
        key = "cn_bold" if (bold and self.resolved["cn_bold"]) else ("cn" if _HAS_CJK.match("中") else role)
        if role == "latin" and not bold:
            path = self.resolved["latin"] or self.resolved["cn"]
            index = 0
        else:
            path = self.resolved["cn_bold"] if (bold and self.resolved["cn_bold"]) else self.resolved["cn"]
            index = int(self.policy.req("typography.ttc_index")) if path and path.endswith(".ttc") else 0
            if path is None:
                path = self.resolved["latin"]
                index = 0
        if path is None:
            raise FontUnavailable("没有可用的字体文件")
        return ImageFont.truetype(path, max(1, int(round(size))), index=index)

    def measure(self, text: str, size: float, *, bold: bool = False,
                tracking_pt: float = 0.0, latin: bool = False) -> float:
        """返回文字前进宽度（px）。tracking 为字距（pt），按 ppt_dpi 换算。"""
        if not text:
            return 0.0
        px_per_pt = self.policy.num("canvas.ppt_dpi") / 72.0
        tracking_px = tracking_pt * px_per_pt
        try:
            font = self._font("latin" if latin else "cn", size, bold)
            base = float(font.getlength(text))
        except (FontUnavailable, OSError):
            self.degraded = True
            base = _estimate(text, size)
        if tracking_px:
            base += tracking_px * max(0, len(text) - 1)
        return base


def _estimate(text: str, size: float) -> float:
    """无字体时的保守估计：中日韩全宽、拉丁 0.58 宽。"""
    width = 0.0
    for ch in text:
        if _HAS_CJK.match(ch) or ch in "，。、；：？！（）「」『』【】《》％":
            width += size
        elif ch == " ":
            width += size * 0.32
        else:
            width += size * 0.58
    return width


# ──────────────────────────────────────────────────────────────────────────
# 数字与单位绑定
# ──────────────────────────────────────────────────────────────────────────
_UNIT_RE_CACHE: dict = {}


def bind_units(text: str, policy) -> str:
    """在数字与单位之间插入 NBSP，保证断行时不把「186 / 公里」拆开。"""
    if not policy.req("typography.number_unit.bind"):
        return text
    units = policy.req("typography.number_unit.unit_patterns") or []
    key = tuple(units)
    pattern = _UNIT_RE_CACHE.get(key)
    if pattern is None:
        ordered = sorted((u for u in units if u), key=len, reverse=True)
        pattern = re.compile(r"(?<=\d)[ \u00a0]?(" + "|".join(re.escape(u) for u in ordered) + r")")
        _UNIT_RE_CACHE[key] = pattern
    return pattern.sub(lambda m: NBSP + m.group(1), text)


# ──────────────────────────────────────────────────────────────────────────
# 断行
# ──────────────────────────────────────────────────────────────────────────
_BREAK_AFTER_SPACE = set(" \u00a0")


def _unit_tokens(policy) -> tuple:
    units = [u for u in (policy.req("typography.number_unit.unit_patterns") or []) if u]
    return tuple(sorted(set(units), key=len, reverse=True))


def tokenize(text: str, policy=None) -> list[str]:
    """切成不可再分的排印单元。

    中文可以逐字断（这是中文排版的正常规则），但两类东西不能被拆开：
      * 数字与它的单位（「186 公里」不是「186 / 公里」）；
      * 多字单位本身（「万元」不是「万 / 元」）。
    这两条是中文排版的硬约束，靠 token 化保证，而不是靠事后修补。
    """
    units = _unit_tokens(policy) if policy is not None else ()
    tokens: list[str] = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch in " \u00a0":
            if tokens:
                tokens[-1] += ch
            i += 1
            continue
        matched = None
        m = _NUMBER.match(text, i)
        if m:
            # 数字后面紧跟（可隔一个空格/NBSP）白名单单位时，合成一个不可分单元
            matched = m.group(0)
            rest = text[m.end():]
            gap = " " if rest[:1] in (" ", "\u00a0") else ""
            for unit in units:
                if rest[len(gap):].startswith(unit):
                    matched += gap + unit
                    break
        else:
            for unit in units:
                if text.startswith(unit, i):
                    matched = unit
                    break
        if matched is None:
            m = _LATIN_WORD.match(text, i)
            if m:
                matched = m.group(0)
            else:
                j = i + 1
                # 连续的同族字符（全角标点串）合并，避免把「……」拆成四点
                while j < n and text[j] == ch and ch in "….·-—":
                    j += 1
                matched = text[i:j]
        tokens.append(matched)
        i += len(matched)
    return tokens


def wrap(text: str, width: float, size: float, *, policy, fontbook: FontBook,
         bold: bool = False, tracking_pt: float = 0.0, max_lines: int | None = None) -> Metrics:
    """按真实字宽断行；尊重作者显式换行（\\n），不重排作者的语义断行。"""
    size = float(size)
    lh = float(policy.line_height(_ladder_key(policy, size)))
    if width <= 0 or size <= 0:
        return Metrics(lines=(text,), widths=(0.0,), size=size, line_height=lh,
                       degraded=fontbook.degraded)
    forbidden_start = set(policy.req("typography.orphan.forbidden_line_start") or [])
    forbidden_end = set(policy.req("typography.orphan.forbidden_line_end") or [])
    ratio = float(policy.num("thresholds.line_width_ratio"))
    usable = width * ratio

    out: list[str] = []
    for hard_line in str(text).split("\n"):
        if not hard_line:
            out.append("")
            continue
        tokens = tokenize(hard_line, policy)
        line_tokens: list[str] = []
        for token in tokens:
            candidate = "".join(line_tokens) + token
            if line_tokens and fontbook.measure(candidate.rstrip(), size, bold=bold,
                                                tracking_pt=tracking_pt) > usable:
                # 禁则：若下一行将以禁首字符开头，把当前行末单元一起推下去
                if token and token[0] in forbidden_start and len(line_tokens) > 1:
                    moved = line_tokens.pop()
                    out.append("".join(line_tokens).rstrip())
                    line_tokens = [moved, token]
                    continue
                # 回退断点：优先在标点/空格后断，避免把词从中间切开
                cut = _preferred_break(line_tokens, policy)
                if 0 < cut < len(line_tokens):
                    out.append("".join(line_tokens[:cut]).rstrip())
                    line_tokens = line_tokens[cut:]
                else:
                    out.append("".join(line_tokens).rstrip())
                    line_tokens = [token]
                    continue
                line_tokens.append(token)
            else:
                line_tokens.append(token)
        if line_tokens:
            out.append("".join(line_tokens).rstrip())

    # 行尾禁则：行末是开引号/开括号时，把它推到下一行（除非下一行放不下）
    fixed: list[str] = []
    for idx, line in enumerate(out):
        if line and line[-1] in forbidden_end and idx + 1 < len(out):
            nxt = line[-1] + out[idx + 1]
            if fontbook.measure(nxt, size, bold=bold, tracking_pt=tracking_pt) <= usable:
                fixed.append(line[:-1].rstrip())
                out[idx + 1] = nxt
                continue
        fixed.append(line)

    widths = tuple(fontbook.measure(l, size, bold=bold, tracking_pt=tracking_pt) for l in fixed)
    truncated = False
    if max_lines is not None and len(fixed) > max_lines:
        fixed = fixed[:max_lines]
        widths = widths[:max_lines]
        truncated = True
    return Metrics(lines=tuple(fixed), widths=widths, size=size, line_height=lh,
                   degraded=fontbook.degraded, truncated=truncated)


def _preferred_break(tokens: list[str], policy) -> int:
    """在已经排满的一行里找最佳断点：优先标点之后，其次空格之后，且至少保留策略要求的行宽比例。

    只影响「哪里断」，不影响「能不能断」：没有合适断点时退回逐字断（中文的正常规则）。
    哪些标点可以收行、行要排到多满，都是可调策略（typography.line_break.*），不是代码常量。
    """
    break_after = set(str(policy.req("typography.line_break.after_punct")))
    limit = max(1, int(len(tokens) * float(policy.req("typography.line_break.min_line_fill"))))
    punct = space = 0
    for i in range(1, len(tokens) + 1):
        if i < limit:
            continue
        head = tokens[i - 1]
        if not head:
            continue
        if head[-1] in break_after:
            punct = i
        elif head[-1] in _BREAK_AFTER_SPACE:
            space = i
    return punct or space


def _ladder_key(policy, size: float) -> str:
    """把自由字号映射回最接近的语义档（用于取行高/字距/行数上限）。"""
    scale = policy.req("typography.scale")
    return min(scale, key=lambda k: abs(float(scale[k]) - float(size)))


# ──────────────────────────────────────────────────────────────────────────
# 墨迹几何
# ──────────────────────────────────────────────────────────────────────────
def ink_height(lines: int, size: float, line_height: float, policy) -> float:
    if lines <= 0:
        return 0.0
    floor = float(policy.req("typography.ink_height_floor"))
    return (lines - 1) * size * line_height + size * max(line_height, floor)


def ink_top(frame_y: float, frame_h: float, lines: int, size: float,
            line_height: float, anchor: str, policy) -> float:
    height = ink_height(lines, size, line_height, policy)
    if anchor == "bottom":
        return frame_y + frame_h - height
    if anchor == "middle":
        return frame_y + (frame_h - height) / 2.0
    return frame_y


def ink_box(element: dict, policy, fontbook: FontBook) -> dict:
    """元素的实际墨迹框（文本用度量，图形/图表用外框）。返回 {x,y,w,h,kind,lines}。"""
    x, y = float(element["x"]), float(element["y"])
    w, h = float(element["width"]), float(element["height"])
    etype = element["type"]
    if etype != "text":
        if etype == "shape" and element.get("shape") in ("line", "arrow"):
            return {"x": min(x, x + w), "y": min(y, y + h),
                    "w": abs(w) or 1.0, "h": abs(h) or 1.0, "kind": "rule", "lines": 0}
        return {"x": x, "y": y, "w": w, "h": h, "kind": etype, "lines": 0}
    size = float(element.get("size") or policy.size("body"))
    metrics = wrap(str(element.get("text", "")), w, size, policy=policy, fontbook=fontbook,
                   bold=bool(element.get("bold")), tracking_pt=float(element.get("tracking", 0.0) or 0.0))
    lines = metrics.count
    lh = float(element.get("line_height") or metrics.line_height)
    top = ink_top(y, h, lines, size, lh, str(element.get("anchor") or "top"), policy)
    width = min(w, max(metrics.widths) if metrics.widths else 0.0) if metrics.widths else 0.0
    return {"x": x, "y": top, "w": max(width, 1.0), "h": max(ink_height(lines, size, lh, policy), 1.0),
            "kind": "text", "lines": lines, "metrics": metrics}


def orphan_issues(metrics: Metrics, policy) -> list[WrapIssue]:
    """孤字与禁则复查（断行后仍可能出现的形态）。"""
    issues: list[WrapIssue] = []
    minimum = int(policy.req("typography.orphan.min_last_line_chars"))
    if metrics.count > 1 and len(metrics.lines[-1].strip()) < minimum:
        issues.append(WrapIssue("orphan_line", metrics.count - 1,
                                f"末行只有 {len(metrics.lines[-1].strip())} 字：删字或改写，不靠缩字号", "warn"))
    starts = set(policy.req("typography.orphan.forbidden_line_start") or [])
    ends = set(policy.req("typography.orphan.forbidden_line_end") or [])
    for i, line in enumerate(metrics.lines):
        if i and line and line[0] in starts:
            issues.append(WrapIssue("forbidden_line_start", i, f"行首出现禁则字符「{line[0]}」", "warn"))
        if line and line[-1] in ends:
            issues.append(WrapIssue("forbidden_line_end", i, f"行尾出现禁则字符「{line[-1]}」", "warn"))
    return issues


def reading_seconds(text: str, policy) -> float:
    """讲稿时长估算：按字数/语速，只用于节奏提示，不是硬门。"""
    chars = len(re.sub(r"\s", "", text or ""))
    return round(chars / float(policy.num("thresholds.words_per_second")), 1)
