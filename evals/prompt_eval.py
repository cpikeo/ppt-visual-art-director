#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""prompt_eval.py · 提示词层对照（说明是否可执行，不是图是否好看）。

三层结论分开写，绝不互相代替：
  L1 说明完备性  本文件：每份说明必须满足的可判定条款（拍什么 / 为什么 / 给谁让位 / 交付多大）
  L2 对照        同一批意图，用「朴素提示词」（只写画面里有什么）当基线
  L3 图像质量    不在这里。本包不调用图像生成；出图好不好看只有人工盲评（SKILL「证据层级」）

基线是什么：一句「画面里有什么 + clean composition」。这是大多数人交给图像模型的东西，
也正是本包要证明不够的东西——它说不清职责、让位、交付尺寸，反向词只有污染项。

用法：
    python evals/prompt_eval.py                 # 候选 vs 基线
    python evals/prompt_eval.py --report out.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from policy import load as load_policy  # noqa: E402
from prompt import build, quiet_side, ratio_of  # noqa: E402

CASES = ROOT / "evals" / "prompt_cases.yaml"


def load_cases(path: Path = CASES) -> list[dict]:
    import yaml
    return (yaml.safe_load(path.read_text(encoding="utf-8")) or {}).get("cases") or []


def naive(subject: str, box) -> str:
    """基线：把「画面里有什么」当提示词。用来证明「只写主体」为什么不够。"""
    return f"{subject or 'a relevant photo'}; clean composition" if box else (subject or "")


def checks(brief: dict, prompt: str, negative: str, given: dict, policy) -> dict:
    """把一份说明翻译成一组可判定的条款。缺哪条就是缺哪条，不做加权。"""
    phrases = policy.prompts["phrases"]
    ratio = brief["ratio"]
    expected = ratio_of(given["box"][2], given["box"][3], policy)
    result = {
        "subject_stated": bool(given.get("subject")) and given["subject"].strip() in prompt,
        "duty_stated": phrases["function"].get(brief.get("function") or "") in prompt,
        "role_stated": any(line in prompt for line in phrases["role"][brief.get("role") or "subject"]),
        "role_background": phrases["role"]["background"][0] in prompt,
        "quiet_area_stated": bool(brief.get("quiet_side"))
                             and phrases["quiet_area"][brief["quiet_side"]] in prompt,
        "quiet_absent": not brief.get("quiet_side") and "stays quiet" not in prompt,
        "quiet_right": quiet_side(given.get("safe_area")) == "right",
        "no_panel_guard": phrases["quiet_no_panel"] in prompt,
        "tone_stated": bool(brief.get("text_color")) and phrases["tone"][brief["text_color"]] in prompt,
        "no_coordinates": "%" not in prompt and " px" not in prompt,
        "negative_production": "no AI artifacts" in negative,
        "negative_proof": "no dramatization" in negative,
        "proof_plainly": "readable for verification" in prompt,
        "author_avoid": all(f"no {term}" in negative for term in given.get("avoid") or []),
        # 「可用」= 与几何一致、且带着目标像素（生产端能照着设画幅与分辨率）；
        # 不要求一定是常用比例——300×640 的真实图位最简比例是 15:32，如实给出比假装 9:16 有用
        "ratio_usable": ratio == expected and ":" in ratio and bool(brief.get("target_px")),
        "ratio_snapped": ratio in {":".join(str(v) for v in pair)
                                   for pair in policy.prompts["ratio"]["presets"]},
        "target_px_floor": max(brief["target_px"].values()) >= policy.req("assets.min_long_side_px"),
        "no_style_words": not any(
            word in prompt.lower() for word in policy.prompts["style_words_banned"]),
        "subject_missing_reported": "PROMPT_SUBJECT_MISSING" in {i["code"] for i in brief["issues"]},
        "no_fabricated_subject": ("a relevant" not in prompt and "placeholder" not in prompt),
    }
    return result


def run_case(case: dict, policy, *, baseline: bool = False) -> dict:
    given = case["given"]
    brief = build(policy=policy, function=given.get("function", "hero"),
                  role=given.get("role", "subject"), subject=given.get("subject", ""),
                  reason=given.get("reason", ""), box=given["box"],
                  safe_area=given.get("safe_area"), text_color=given.get("text_color", ""),
                  avoid=given.get("avoid") or ())
    if baseline:
        prompt = naive(given.get("subject", ""), given.get("box"))
        negative = "no text, no logo"
        observed = checks(brief, prompt, negative, given, policy)
        # 基线没有说明对象：只保留「主体是否写进提示词」这类可比条款
        observed = {k: observed[k] for k in observed if k in {
            "subject_stated", "no_coordinates", "target_px_floor", "no_style_words"}}
        observed["target_px_floor"] = False       # 朴素提示词不声明交付尺寸
    else:
        observed = checks(brief, brief["prompt"], brief["negative"], given, policy)
    expect = [e for e in case["expect"]]
    failures = [e for e in expect if not observed.get(e)]
    return {"id": case["id"], "name": case["name"], "expect": expect,
            "failures": failures, "passed": not failures,
            "prompt": None if baseline else brief["prompt"],
            "fingerprint": None if baseline else brief["fingerprint"]}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="提示词层对照")
    parser.add_argument("--report")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument("--show", action="store_true", help="打印候选提示词")
    args = parser.parse_args(argv)

    policy = load_policy()
    cases = load_cases()
    candidate = [run_case(c, policy) for c in cases]
    base = [run_case(c, policy, baseline=True) for c in cases]

    def rate(rows) -> tuple[int, int]:
        return sum(1 for r in rows if r["passed"]), len(rows)

    passed, total = rate(candidate)
    b_passed, b_total = rate(base)
    clause_total = sum(len(r["expect"]) for r in candidate)
    clause_ok = sum(len(r["expect"]) - len(r["failures"]) for r in candidate)
    # 基线只对可比条款计数，避免拿「它没有这套词表」当成绩
    comparable = {"subject_stated", "no_coordinates", "target_px_floor", "no_style_words"}

    def base_clauses(rows):
        total = ok = 0
        for r in rows:
            keys = [e for e in r["expect"] if e in comparable]
            total += len(keys)
            ok += len([k for k in keys if k not in r["failures"]])
        return ok, total

    base_clause_ok, base_clause_total = base_clauses(base)

    report = {
        "schema": "pvad-prompt-eval-v1",
        "policy_sha256": policy.sha256,
        "cases": len(cases),
        "candidate": {"passed": passed, "total": total,
                      "clauses_ok": clause_ok, "clauses": clause_total},
        "baseline": {"passed": b_passed, "total": b_total,
                     "clauses_ok": base_clause_ok, "clauses": base_clause_total,
                     "what": "朴素提示词：画面里有什么 + clean composition"},
        "rows": candidate,
        "note": ("只度量说明是否可执行（拍什么/为什么/给谁让位/交付多大），"
                 "不度量出图质量。没有调用任何图像模型：本包不替外部生成器出图。"),
    }
    if args.report:
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=1) + "\n",
                                     encoding="utf-8")
    if not args.quiet:
        print(f"提示词说明完备性  {passed}/{total} 用例通过；条款 {clause_ok}/{clause_total}")
        print(f"朴素基线          {b_passed}/{b_total} 用例通过；条款 {base_clause_ok}/{base_clause_total}")
        for row in candidate:
            if row["failures"]:
                print(f"  - {row['id']} {row['name']}: 缺 {row['failures']}")
            if args.show and row["prompt"]:
                print(f"  · {row['id']}: {row['prompt']}")
        print("注：这只说明「说明写得全不全」，不代表出图好看；图像质量需人工盲评（SKILL「证据层级」）。")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
