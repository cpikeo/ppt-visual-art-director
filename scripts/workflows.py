#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Derived workflow views: minimal THINK handoff and modification impact scope.

These are views over plan/deck truth, not new authority sources.
"""
from __future__ import annotations

import hashlib
import json


def _digest(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()


def handoff_markdown(plan: dict) -> str:
    pages = plan.get("pages") or []
    deck = plan.get("deck") or {}
    lines = [
        "# THINK → BUILD 最小交接",
        "",
        "> 本文件从 plan 确定性派生，不是第二份机器真值；plan 改动后必须重新生成。",
        "",
        f"- 目标决定：{deck.get('decision') or '未决'}",
        f"- 整稿主张 / 张力：{deck.get('tension') or '未决'}",
        f"- 视觉概念：{deck.get('visual_world') or '未决'}",
        f"- Plan digest：`{_digest(plan)}`",
        "",
    ]
    for page in pages:
        claim = page.get("claim") or {}
        evidence = (page.get("understanding") or {}).get("evidence") or {}
        decision = page.get("decision") or {}
        production = page.get("production") or {}
        questions = page.get("open_questions") or []
        question_lines = [f"{q.get('code')}: {q.get('question')}" for q in questions]
        lines += [
            f"## {page.get('id')}｜{page.get('title')}",
            "",
            f"- 页面任务：{decision.get('role', '未决')}；节奏={(decision.get('rhythm') or {}).get('move', '未决')}",
            f"- 核心主张：{claim.get('text') or '未决'}（{claim.get('source', 'absent')}）",
            f"- 证据状态：{evidence.get('value', 'none')}；{evidence.get('why', '未说明')}",
            f"- 主导焦点：{decision.get('focus', '未决')}",
            f"- 构图理由：{decision.get('composition', '未决')}；{decision.get('composition_why', '未说明')}",
            f"- 必须保留：{_join(production.get('must_place'))}",
            f"- 禁止推断 / 禁止执行：{_join(production.get('must_not'))}",
            f"- 未决事项：{_join(question_lines)}",
            "- 验收条件：主张—证据—焦点一致；plan—PPTX 对象—实际渲染一致；项目强制门有证据。",
            "- 失败回退：事实/关系问题回 THINK；构图/编译问题回 BUILD；目标环境差异回受影响实现。",
            "",
        ]
    return "\n".join(lines)


def _join(values) -> str:
    values = [str(value) for value in (values or []) if str(value).strip()]
    return "；".join(values) if values else "无"


def impact(before: dict, after: dict) -> dict:
    """Conservative verification scope for author-made deck changes."""
    old = {slide.get("id"): slide for slide in before.get("slides", [])}
    new = {slide.get("id"): slide for slide in after.get("slides", [])}
    old_order = [slide.get("id") for slide in before.get("slides", [])]
    new_order = [slide.get("id") for slide in after.get("slides", [])]
    added = [sid for sid in new_order if sid not in old]
    removed = [sid for sid in old_order if sid not in new]
    changed = [sid for sid in new_order if sid in old and _digest(old[sid]) != _digest(new[sid])]
    direct = set(added + removed + changed)
    adjacent = set()
    for order in (old_order, new_order):
        for index, sid in enumerate(order):
            if sid in direct:
                if index > 0:
                    adjacent.add(order[index - 1])
                if index + 1 < len(order):
                    adjacent.add(order[index + 1])
    compositions = {((old.get(sid) or new.get(sid) or {}).get("composition")) for sid in direct}
    same_class = {sid for sid, slide in {**old, **new}.items()
                  if slide.get("composition") in compositions and slide.get("composition")}
    scope = direct | adjacent | same_class
    return {
        "schema": "pvad-impact-v1",
        "before_digest": _digest(before),
        "after_digest": _digest(after),
        "added": added,
        "removed": removed,
        "changed": changed,
        "adjacent_review": sorted(adjacent - direct),
        "same_composition_review": sorted(same_class - direct),
        "verification_scope": sorted(scope),
        "required_checks": [
            "schema_and_guard_for_entire_deck",
            "plan_object_render_for_changed_pages",
            "narrative_review_for_changed_and_adjacent_pages",
            "same_root_cause_review_for_same_composition",
        ],
        "limitation": "本结果比较 deck JSON；不能从任意既有 PPTX 恢复作者主张、证据状态或设计意图。",
    }
