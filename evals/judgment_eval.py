#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""judgment_eval.py · 判断基准（离散判断，不含美感）。

三层结论分开写，绝不互相代替：
  L1 判断准确率   本文件：cases.yaml 的离散判断命中率 + 误判明细
  L2 消融覆盖率   --ablate：逐条去掉词表/规则，看哪些用例会翻转
                  （翻转 = 该规则在当前用例集上是被测到的；不翻转 ≠ 无用，只说明没测到）
  L3 视觉质量    不在这里。见 evals/blind_pack.py 与 SKILL §6：人工盲评，未跑就写未跑。

用法：
    python evals/judgment_eval.py                       # 打基准
    python evals/judgment_eval.py --report out.json     # 存报告
    python evals/judgment_eval.py --baseline            # 与「无判断基线」对照
    python evals/judgment_eval.py --ablate              # 词表消融
"""
from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
_TESTS = ROOT / "tests"
for _path in (str(SCRIPTS), str(_TESTS)):
    if _path not in sys.path:
        sys.path.insert(0, _path)

import support  # noqa: E402  (tests/support.py 提供 brief 写入)
from policy import Policy, load as load_policy  # noqa: E402
from brief import load as load_brief  # noqa: E402
from judgment import judge_deck  # noqa: E402

SCHEMA_DIR = ROOT / "schema"
CASES = ROOT / "evals" / "cases.yaml"
ABLATABLE = ("comparison_end", "series", "part_whole", "denominator_risk", "sequence_strong",
             "structure", "claim_cue", "change_cue", "imperative_cue", "ask",
             "causal_risk", "correlation_cue", "conditional", "visual_subject")


def load_cases(path: Path = CASES) -> dict:
    import yaml
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def run_case(case: dict, policy: Policy, *, baseline: bool = False) -> dict:
    deck = dict(load_cases()["deck"])
    deck.update(case.get("deck_override") or {})
    slides = case.get("slides") or [case["slide"]]
    brief_text = _to_yaml({**deck, "slides": slides})
    tmp = support.tmpdir()
    brief, _ = load_brief(support.write_brief(tmp, text=brief_text), policy, SCHEMA_DIR)
    plan = judge_deck(brief, policy)
    pages = plan["pages"]
    if baseline:
        return _naive(pages)
    index = 0
    page = pages[index]
    return {
        "claim_source": page["claim"]["source"],
        "claim_text": page["claim"]["text"],
        "claim_form": page["claim"]["form"],
        "role": page["decision"]["role"],
        "evidence": page["understanding"]["evidence"]["value"],
        "composition": page["decision"]["composition"],
        "delta": page["understanding"]["comparison"]["delta"],
        "denominator_proven": page["understanding"]["denominator"]["proven"],
        "questions": [q["code"] for q in page["open_questions"]],
        "media_decision": page["decision"]["media"]["decision"],
        "media_source": page["decision"]["media"]["source"],
        "media_function": page["decision"]["media"]["function"],
        "pages": [{"role": p["decision"]["role"], "composition": p["decision"]["composition"],
                   "claim_source": p["claim"]["source"]} for p in pages],
    }


def _naive(pages: list) -> dict:
    """无判断基线：把所有页当成「一段文字 + 一个标题」的排版任务，不做任何证据判断。

    这就是「有架构、没有设计判断」的系统会给出的结果，用作对照。
    """
    return {
        "claim_source": "author", "claim_text": pages[0]["title"],
        "role": "explain", "evidence": "proposition", "composition": "statement",
        "delta": False, "denominator_proven": False, "questions": [],
        "media_decision": "none", "media_source": "rule:no-media", "media_function": "",
        "pages": [{"role": "explain", "composition": "statement", "claim_source": "author"}
                  for _ in pages],
    }


def _to_yaml(value) -> str:
    import yaml
    return yaml.safe_dump(value, allow_unicode=True, sort_keys=False)


def check(case: dict, actual: dict) -> list[str]:
    """返回失败原因列表（空 = 通过）。expect / expect_index0 / expect_index1 均可。"""
    failures = []
    for key, expected in (case.get("expect") or {}).items():
        failures += _assert(key, expected, actual, prefix="")
    if "expect_index0" in case:
        for key, expected in case["expect_index0"].items():
            failures += _assert(key, expected, actual, prefix="pages[0].")
    if "expect_index1" in case:
        for key, expected in case["expect_index1"].items():
            page = (actual.get("pages") or [{}, {}])[1]
            value = _lookup(page, key)
            if value != expected:
                failures.append(f"pages[1].{key}: 期望 {expected!r}，实际 {value!r}")
    return failures


def _lookup(actual: dict, key: str):
    if key.startswith("question"):
        return actual.get("questions")
    return actual.get(key)


def _assert(key: str, expected, actual: dict, prefix: str) -> list[str]:
    page = actual.get("pages") or [{}]
    if prefix == "pages[0].":
        source = page[0]
    else:
        source = actual
    if key.endswith("_not"):
        field = key[:-4]
        value = source.get(field)
        return [] if value != expected else [f"{prefix}{field}: 不应为 {expected!r}"]
    if key == "media_decision":
        return ([] if source.get("media_decision") == expected
                else [f"{prefix}media_decision: 期望 {expected!r}，实际 {source.get('media_decision')!r}"])
    if key == "media_source_prefix":
        value = str(source.get("media_source") or "")
        return ([] if value.startswith(expected)
                else [f"{prefix}media_source: 期望以 {expected!r} 开头，实际 {value!r}"])
    if key == "media_function":
        return ([] if source.get("media_function") == expected
                else [f"{prefix}media_function: 期望 {expected!r}，实际 {source.get('media_function')!r}"])
    if key == "question":
        codes = source.get("questions") or []
        return [] if expected in codes else [f"{prefix}question: 缺 {expected}，实际 {codes}"]
    if key == "claim_contains":
        text = source.get("claim_text") or ""
        return [] if expected in text else [f"{prefix}claim: {text!r} 不含 {expected!r}"]
    value = source.get(key)
    return [] if value == expected else [f"{prefix}{key}: 期望 {expected!r}，实际 {value!r}"]


def evaluate(policy: Policy, *, baseline: bool = False) -> dict:
    suite = load_cases()
    results = []
    for case in suite["cases"]:
        actual = run_case(case, policy, baseline=baseline)
        failures = check(case, actual)
        results.append({"id": case["id"], "category": case["category"], "statement": case["statement"],
                        "protocol": case["protocol"], "passed": not failures,
                        "failures": failures, "actual": actual})
    total = len(results)
    passed = sum(1 for r in results if r["passed"])
    by_category: dict[str, dict] = {}
    for r in results:
        bucket = by_category.setdefault(r["category"], {"total": 0, "passed": 0})
        bucket["total"] += 1
        bucket["passed"] += 1 if r["passed"] else 0
    return {"suite": str(CASES.name), "baseline": baseline, "total": total, "passed": passed,
            "accuracy": round(passed / total, 4) if total else 0.0,
            "by_category": by_category,
            "failures": [{"id": r["id"], "failures": r["failures"], "protocol": r["protocol"]}
                         for r in results if not r["passed"]],
            "results": results}


def ablate(policy: Policy) -> dict:
    """逐条词表消融：去掉一条词表后，哪些用例会翻转。"""
    base = evaluate(policy)
    base_fail = {r["id"] for r in base["results"] if not r["passed"]}
    out = []
    for name in ABLATABLE:
        mutated = _policy_without(policy, name)
        result = evaluate(mutated)
        flipped = sorted({r["id"] for r in result["results"] if not r["passed"]} - base_fail)
        out.append({"lexicon": name, "flipped_cases": flipped,
                    "covered": bool(flipped),
                    "note": ("该词表被当前用例集覆盖：删掉会改变判断" if flipped
                             else "当前用例集没有覆盖到这条规则的独立作用：不证明它无用，只说明没测到")})
    uncovered = [row["lexicon"] for row in out if not row["covered"]]
    return {"baseline_accuracy": base["accuracy"], "rows": out,
            "uncovered_lexicons": uncovered,
            "reading": ("消融只回答「用例集测到没测到」；未覆盖不等于可删。"
                        "任何一条被删前必须先补一条会失败的用例。")}


def _policy_without(policy: Policy, lexicon: str) -> Policy:
    strategy = copy.deepcopy(policy.strategy)
    strategy["judgment"]["lexicons"][lexicon] = []
    return Policy(strategy=strategy, composition=policy.composition, sha256=policy.sha256,
                  paths=policy.paths)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="判断基准")
    parser.add_argument("--baseline", action="store_true", help="与无判断基线对照")
    parser.add_argument("--ablate", action="store_true", help="逐条词表消融")
    parser.add_argument("--report", help="把报告写成 JSON")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    policy = load_policy()
    candidate = evaluate(policy)
    report = {"candidate": candidate, "policy_sha256": policy.sha256}

    if args.baseline:
        report["baseline"] = evaluate(policy, baseline=True)
    if args.ablate:
        report["ablation"] = ablate(policy)

    if not args.quiet:
        print(f"判断基准：{candidate['passed']}/{candidate['total']} 通过 "
              f"({candidate['accuracy']:.1%})  ·  用例集 {candidate['suite']}")
        for category, bucket in sorted(candidate["by_category"].items()):
            print(f"  {category:12s} {bucket['passed']}/{bucket['total']}")
        if candidate["failures"]:
            print("失败用例：")
            for item in candidate["failures"]:
                print(f"  {item['id']}: {'; '.join(item['failures'])}  ← {item['protocol']}")
        if args.baseline:
            base = report["baseline"]
            print(f"无判断基线：{base['passed']}/{base['total']} ({base['accuracy']:.1%})"
                  f"  →  判断协议带来的差距 {candidate['accuracy'] - base['accuracy']:+.1%}")
        if args.ablate:
            print("词表消融（翻转 = 用例集测到了该规则的独立作用）：")
            for row in report["ablation"]["rows"]:
                mark = "覆盖" if row["covered"] else "未覆盖"
                print(f"  {row['lexicon']:18s} {mark:6s} {','.join(row['flipped_cases']) or '—'}")
        print("注意：本基准只测离散判断，不测构图、断行或美感；视觉质量需人工盲评（SKILL §6）。")

    if args.report:
        Path(args.report).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if candidate["passed"] == candidate["total"] else 1


if __name__ == "__main__":
    sys.exit(main())
