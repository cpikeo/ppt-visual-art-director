#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""typography_eval.py · 中文排版的可复现对照（候选 vs 朴素实现）。

为什么要有这个对照：
  「中文排版更好」如果只是一句主张，就无法回归。这里把它拆成三个机器可判的事实，
  在同一语料、同一字体度量、同一宽度集合上同时跑两条实现：

    baseline  朴素实现：逐字贪心断行，不断行空格、不注意禁则、不在标点处优先断
              （这是「有断行就算排版」的常见做法）
    candidate 本技能包：数字与单位绑定 + 禁则 + 标点优先断点

  三个指标都是「越低越好」：
    unit_split        数字与单位被拆到两行的比例
    forbidden_start   以禁则字符（，。）开头的行占比
    char_wrap         没有任何标点/空格可依、纯逐字断开的行占比（反映可读性）

局限：这是度量模型里的排版，不是目标 PowerPoint 的真实断行；字体回退与
真实换行仍须在 PowerPoint 中复核（SKILL「中文排版」「证据层级」）。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _path in (str(ROOT / "scripts"), str(ROOT / "tests")):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from policy import load as load_policy  # noqa: E402
from typography import FontBook, bind_units, wrap  # noqa: E402

CORPUS = [
    "单店月均产出为 186 公里，需要复核口径。",
    "门店数量增长 18%，单店月均产出下降 21%。",
    "经销渠道的单店产出高出直营门店 2.4 倍。",
    "把 Q1 新增预算的 70% 转向经销渠道，用同样的钱换更高的单店产出。",
    "三项相加为 100%，但一个客户可以同时在三个渠道下单。",
    "2026 Q1 到 Q3，经销渠道收入占比从 34% 升到 51%。",
    "直营门店的边际产出已经低于经销渠道，继续加店不再带来增量。",
    "交付平均 42 天，客户要求的周期是 30 天，差距需要靠并行评审补上。",
    "如果 Q1 继续扩张，单店产出会回升；但这一点还没有被数据证明。",
    "门店数量和好评率同时上升，不能据此推断服务改进带来了增长。",
    "会员业务占比 40%，零售 35%，服务 25%，三者共用同一分母。",
    "供应链成本每公斤下降 1.2 元，全年可节省约 430 万元。",
]
WIDTHS = (180, 220, 280, 360, 480)
SIZES = (18, 22, 30)


def naive_wrap(text: str, width: float, size: float, fontbook: FontBook, policy) -> list[str]:
    """朴素实现：逐字贪心，不绑定单位、不处理禁则、不挑断点。"""
    lines, line = [], ""
    for ch in text:
        if fontbook.measure(line + ch, size) > width:
            lines.append(line)
            line = ch
        else:
            line += ch
    if line:
        lines.append(line)
    return lines


def measure(lines: list[str], policy) -> dict:
    forbidden = set(policy.req("typography.orphan.forbidden_line_start") or [])
    units = sorted({u for u in policy.req("typography.number_unit.unit_patterns") or [] if u},
                   key=len, reverse=True)
    unit_split = 0
    for prev, nxt in zip(lines, lines[1:]):
        tail = prev.rstrip()
        head = nxt.lstrip()
        if any(ch.isdigit() for ch in tail[-1:]) or tail.endswith((" ", "\u00a0")):
            if any(head.startswith(u) for u in units):
                unit_split += 1
    forbidden_start = sum(1 for line in lines[1:] if line[:1] and line[0] in forbidden)
    char_wrap = 0
    for prev, nxt in zip(lines, lines[1:]):
        tail = prev.rstrip()
        head = nxt.lstrip()
        if tail[-1:] in "，。、；：！？,.;:!?）】」\u00a0 ":
            continue
        if any(head.startswith(u) for u in units):
            continue
        char_wrap += 1
    breaks = max(1, len(lines) - 1)
    return {"unit_split": unit_split, "forbidden_start": forbidden_start,
            "char_wrap": char_wrap, "breaks": breaks}


def run(policy, fontbook) -> dict:
    totals = {"baseline": {"unit_split": 0, "forbidden_start": 0, "char_wrap": 0, "breaks": 0},
              "candidate": {"unit_split": 0, "forbidden_start": 0, "char_wrap": 0, "breaks": 0}}
    samples = 0
    for sentence in CORPUS:
        bound = bind_units(sentence, policy)
        for width in WIDTHS:
            for size in SIZES:
                samples += 1
                base_lines = naive_wrap(sentence, width, size, fontbook, policy)
                cand_lines = wrap(bound, width, size, policy=policy, fontbook=fontbook).lines
                for key, lines in (("baseline", base_lines), ("candidate", cand_lines)):
                    for metric, value in measure(list(lines), policy).items():
                        totals[key][metric] += value
    result = {"samples": samples, "corpus": len(CORPUS), "widths": list(WIDTHS), "sizes": list(SIZES)}
    for key, values in totals.items():
        result[key] = {
            "unit_split_pairs": values["unit_split"],
            "unit_split_rate": round(values["unit_split"] / max(1, values["breaks"]), 4),
            "forbidden_start_lines": values["forbidden_start"],
            "forbidden_start_rate": round(values["forbidden_start"] / max(1, values["breaks"]), 4),
            "char_wrap_rate": round(values["char_wrap"] / max(1, values["breaks"]), 4),
        }
    result["delta"] = {
        metric: round(result["candidate"][f"{metric}_rate"] - result["baseline"][f"{metric}_rate"], 4)
        for metric in ("unit_split", "forbidden_start", "char_wrap")
    }
    result["note"] = "度量模型结果，不等于目标 PowerPoint 的真实断行；见 SKILL「中文排版」「证据层级」。"
    return result


def main(argv=None) -> int:
    policy = load_policy()
    fontbook = FontBook(policy)
    if fontbook.degraded:
        print("警告：没有可用字体文件，本对照结果不可采信", file=sys.stderr)
        return 2
    report = run(policy, fontbook)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    out = ROOT / "evals" / "reports" / "typography-baseline.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
