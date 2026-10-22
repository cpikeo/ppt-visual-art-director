#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pvad.py · 唯一入口：plan → build → check（一趟确定性批处理）。

  plan    brief → plan.json（判断 + 依据 + 未决问题）
  build   plan + brief → deck.json（几何 + 排版 + 数据；等价于一次完整草稿，作者可改）
  check   deck.json（或 build.py） → 硬门 + PPTX + 回读对账 + 渲染对照 → PASS/BLOCK

纪律：
  * 失败只有一种处理：读 fix_groups，按根因组一次改完，再跑同一条命令。
  * release 必须过来源硬门与回读对账；任一缺失即 BLOCK，不拿「结构预览正常」当验收。
  * 退出码：0 通过 / 1 拦截 / 2 用法或输入错误（策略、brief、路径）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
SCHEMA_DIR = ROOT / "schema"
for path in (str(SCRIPTS),):
    if path not in sys.path:
        sys.path.insert(0, path)

from policy import PolicyError, load as load_policy, to_json as policy_json  # noqa: E402


# ──────────────────────────────────────────────────────────────────────────
# 工具
# ──────────────────────────────────────────────────────────────────────────
def read_json(path) -> dict:
    source = Path(path).expanduser()
    if not source.is_file():
        raise InputError(f"找不到文件：{source}")
    try:
        value = json.loads(source.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise InputError(f"JSON 解析失败：{source}:{exc.lineno}:{exc.colno} — {exc.msg}") from None
    if not isinstance(value, dict):
        raise InputError(f"顶层必须是对象：{source}")
    return value


def write_json(path, value) -> str:
    target = Path(path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True)
    _atomic_write(target, (text + "\n").encode("utf-8"))
    return str(target)


def _atomic_write(target: Path, payload: bytes) -> None:
    tmp = target.with_name(target.name + ".tmp")
    tmp.write_bytes(payload)
    tmp.replace(target)


def sha256_file(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def digest(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()


class InputError(RuntimeError):
    """输入/用法错误：退出码 2。"""


def load_deck(path) -> tuple[dict, str]:
    source = Path(path).expanduser()
    if not source.is_file():
        raise InputError(f"找不到 deck：{source}")
    if source.suffix.lower() == ".py":
        return _load_python_deck(source), str(source)
    deck = read_json(source)
    return deck, str(source)


def _load_python_deck(source: Path) -> dict:
    """作者写的 build.py：显式可信代码（不是沙箱），必须定义 build_deck(plan, brief)。"""
    text = source.read_text(encoding="utf-8")
    try:
        code = compile(text, str(source), "exec")
    except SyntaxError as exc:
        raise InputError(f"build.py 语法错误：{source}:{exc.lineno} — {exc.msg}") from None
    namespace = {"__name__": "pvad_author_build", "__file__": str(source)}
    exec(code, namespace)  # noqa: S102 - 作者代码，文档已声明信任边界
    if "build_deck" not in namespace:
        raise InputError("build.py 必须定义 build_deck(plan, brief) 或改用 deck.json")
    return namespace["build_deck"](namespace.get("PLAN"), namespace.get("BRIEF"))


def validate_deck(deck: dict) -> list[str]:
    return validate_against(deck, "deck.schema.json")


def validate_against(value: dict, schema_file: str) -> list[str]:
    import jsonschema
    schema = json.loads((SCHEMA_DIR / schema_file).read_text(encoding="utf-8"))
    validator = jsonschema.Draft202012Validator(schema)
    return [f"{'/'.join(str(p) for p in e.absolute_path) or '<root>'}: {e.message}"
            for e in sorted(validator.iter_errors(value), key=lambda e: list(e.absolute_path))]


# ──────────────────────────────────────────────────────────────────────────
# plan
# ──────────────────────────────────────────────────────────────────────────
def run_plan(args, policy) -> int:
    from brief import BriefError, load as load_brief
    from judgment import judge_deck
    from typography import FontBook

    brief_path = Path(args.brief).expanduser()
    try:
        brief, brief_sha = load_brief(brief_path, policy, SCHEMA_DIR)
    except BriefError as exc:
        raise InputError(str(exc)) from None
    base_dir = brief_path.expanduser().resolve().parent
    plan = judge_deck(brief, policy, base_dir=base_dir)
    plan_errors = validate_against(plan, "plan.schema.json")
    if plan_errors:
        raise InputError("plan 结构与 schema 不符（这是引擎缺陷，不是稿子问题）：\n"
                         + "\n".join(f"  - {e}" for e in plan_errors[:8]))
    plan["brief_path"] = str(brief_path.resolve())
    plan["policy_files"] = list(policy.paths)
    plan["measurement"] = {"degraded": FontBook(policy).degraded,
                           "font_files": FontBook(policy).resolved}
    plan_path = write_json(args.out, plan)

    summary = {
        "plan_path": plan_path,
        "pages": len(plan["pages"]),
        "policy_sha256": plan["policy_sha256"],
        "brief_sha256": brief_sha,
        "claims_absent": [p["id"] for p in plan["pages"] if p["claim"]["source"] == "absent"],
        "open_questions": sum(len(p["open_questions"]) for p in plan["pages"]),
        "warnings": [w["code"] for w in plan["deck"]["warnings"]],
        "plan_digest": digest(plan),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if args.strict and summary["claims_absent"]:
        print("BLOCK — 有页面主张未决，先改内容再落版：", ", ".join(summary["claims_absent"]))
        return policy.req("verification.exit_codes.blocked")
    return policy.req("verification.exit_codes.ok")


# ──────────────────────────────────────────────────────────────────────────
# build
# ──────────────────────────────────────────────────────────────────────────
def run_build(args, policy) -> int:
    from brief import BriefError, load as load_brief
    from compose import compose_deck
    from typography import FontBook

    plan = read_json(args.plan)
    if plan.get("schema") != "pvad-plan-v1":
        raise InputError(f"plan 版本不认识：{plan.get('schema')!r}")
    plan_errors = validate_against(plan, "plan.schema.json")
    if plan_errors:
        raise InputError("plan 结构与 schema 不符：\n" + "\n".join(f"  - {e}" for e in plan_errors[:10]))
    if plan.get("policy_sha256") != policy.sha256:
        raise InputError("plan 与当前策略指纹不一致：策略改过就要重跑 plan，不要用旧判断配新几何")
    brief_path = args.brief or plan.get("brief_path")
    if not brief_path:
        raise InputError("缺少 brief：plan 里没有 brief_path，请显式给 --brief")
    try:
        brief, _ = load_brief(brief_path, policy, SCHEMA_DIR)
    except BriefError as exc:
        raise InputError(str(exc)) from None
    if brief.get("_brief_sha256") != plan.get("brief_sha256"):
        raise InputError("brief 在 plan 之后被改过：重跑 plan，避免判断与原文错位")

    blocking_codes = set(policy.req("verification.build_blocking_open_questions"))
    unresolved = [
        {"slide": page.get("id"), "code": question.get("code"), "question": question.get("question")}
        for page in plan.get("pages") or []
        for question in page.get("open_questions") or []
        if question.get("code") in blocking_codes
    ]
    if unresolved:
        print(json.dumps({"status": "BLOCK — BUILD", "unresolved": unresolved},
                         ensure_ascii=False, indent=2))
        return policy.req("verification.exit_codes.blocked")

    deck = compose_deck(brief, plan, policy, fontbook=FontBook(policy),
                        base_dir=Path(brief_path).expanduser().resolve().parent)
    errors = validate_deck(deck)
    if errors:
        raise InputError("deck 结构与 schema 不符：\n" + "\n".join(f"  - {e}" for e in errors[:10]))
    path = write_json(args.out, deck)
    print(json.dumps({"deck_path": path, "slides": len(deck["slides"]),
                      "elements": sum(len(s["elements"]) for s in deck["slides"]),
                      "deck_digest": digest(deck)}, ensure_ascii=False, indent=2))
    return policy.req("verification.exit_codes.ok")


# ──────────────────────────────────────────────────────────────────────────
# check
# ──────────────────────────────────────────────────────────────────────────
def run_media(args, policy) -> int:
    """把「还需要生产的图」连同生产说明导出，交给摄影师或外部生成器。

    本包不调用图像生成：只出说明（SKILL「素材艺术指导与合规」），出图后的质量由 QC 硬门与人判分别负责。
    """
    from prompt import manifest

    deck_path = Path(args.deck).expanduser()
    try:
        deck = json.loads(deck_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise InputError(f"读不了 deck：{deck_path} — {exc}") from None
    errors = validate_deck(deck)
    if errors:
        raise InputError("deck 结构与 schema 不符：\n" + "\n".join(f"  - {e}" for e in errors[:8]))
    brief = manifest(deck, policy)
    if args.out:
        write_json(args.out, brief)
    summary = {
        "images": brief["images"],
        "to_produce": brief["to_produce"],
        "entries": [{k: e[k] for k in ("id", "status", "ratio", "target_px", "fingerprint")}
                    for e in brief["entries"]],
        "out": str(args.out) if args.out else None,
        "note": "提示词只说明拍什么/为什么/给谁让位/交付多大；不保证好看，QC 与盲评另有其门",
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


def run_handoff(args, policy) -> int:
    """从 plan 派生最小 THINK→BUILD 交接；不创建第二份权威真值。"""
    from workflows import handoff_markdown

    plan = read_json(args.plan)
    errors = validate_against(plan, "plan.schema.json")
    if errors:
        raise InputError("plan 结构与 schema 不符：\n" + "\n".join(f"  - {e}" for e in errors[:10]))
    text = handoff_markdown(plan)
    target = Path(args.out).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    _atomic_write(target, text.encode("utf-8"))
    print(json.dumps({"handoff_path": str(target), "pages": len(plan.get("pages") or []),
                      "source": str(Path(args.plan).resolve()), "derived": True},
                     ensure_ascii=False, indent=2))
    return policy.req("verification.exit_codes.ok")


def run_impact(args, policy) -> int:
    """比较改稿前后 deck JSON，给出保守验证范围和变更摘要。"""
    from workflows import impact

    before, _ = load_deck(args.before)
    after, _ = load_deck(args.after)
    for label, deck in (("before", before), ("after", after)):
        errors = validate_deck(deck)
        if errors:
            raise InputError(f"{label} deck 结构与 schema 不符：\n" + "\n".join(f"  - {e}" for e in errors[:10]))
    result = impact(before, after)
    if args.out:
        write_json(args.out, result)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return policy.req("verification.exit_codes.ok")


def run_package(args, policy) -> int:
    """Create Draft/Review/Release package; Release is impossible while delivery is blocked."""
    from release import ReleaseEvidenceError, build_package

    try:
        result = build_package(args.manifest, args.stage, args.out)
    except (OSError, json.JSONDecodeError, ReleaseEvidenceError) as exc:
        raise InputError(f"交付包生成失败：{exc}") from None
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return policy.req("verification.exit_codes.ok")


def run_check(args, policy) -> int:
    from compile import CompileError, compile_deck, readback
    from fidelity import compare
    from guard import check as guard_check
    from render import (contact_sheet, page_model_from_readback, page_model_from_spec,
                        pixel_diff, render_page)
    from typography import FontBook

    mode = args.mode or policy.req("verification.default_mode")
    if mode not in policy.req("verification.modes"):
        raise InputError(f"未知 mode：{mode}")
    deck, deck_path = load_deck(args.deck)
    schema_errors = validate_deck(deck)
    out_dir = Path(args.out_dir or (Path(args.output).parent / "pvad-report")).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = out_dir / "manifest.json"
    manifest = {
        "schema": "pvad-manifest-v1",
        "status": "BLOCKED",
        "mode": mode,
        "deck_path": str(Path(deck_path).resolve()),
        "deck_sha256": sha256_file(deck_path) if Path(deck_path).suffix != ".py" else digest(deck),
        "spec_digest": digest(deck),
        "policy": policy_json(policy),
        "verification": {
            "structural": "pending",
            "pptx_readback": "pending",
            "spec_render": "pending",
            "target_powerpoint": "NOT VERIFIED — 未在目标 PowerPoint 中打开、编辑、另存",
            "target_renderer": "NOT VERIFIED — 未使用 PowerPoint / LibreOffice 渲染真实投影结果",
            "independent_visual_review": "NOT RUN — 需人工盲评（evals/blind_pack.py 生成 pack 后回收 sheet）",
        },
    }
    write_json(manifest_path, manifest)   # 先落未完成状态，中断后不得读到 PASS

    exit_block = policy.req("verification.exit_codes.blocked")
    if schema_errors:
        guard = {"schema": "pvad-guard-v1", "mode": mode, "status": "BLOCK",
                 "errors": [{"code": "SCHEMA_FAIL", "detail": e, "fix": "按 deck.schema.json 改结构"}
                            for e in schema_errors],
                 "warnings": [], "fix_groups": [{"code": "SCHEMA_FAIL", "count": len(schema_errors),
                                                 "hint": "先过 schema，再谈排版", "items": []}],
                 "stats": {}, "degraded": False}
        manifest.update({"status": "BLOCKED", "guard": guard})
        _apply_delivery_gates(manifest, mode)
        _write_manifest_and_summary(manifest_path, manifest)
        print(json.dumps({"status": "BLOCKED", "fix_groups": guard["fix_groups"]},
                         ensure_ascii=False, indent=2))
        return exit_block

    fontbook = FontBook(policy)
    guard = guard_check(deck, policy, mode=mode)
    result: dict = {"status": guard["status"], "mode": mode, "guard": guard}

    pptx_path = Path(args.output).expanduser()
    if mode == "spec":
        # spec 模式：只回答「结构能不能过」，不编译、不出预览（零 python-pptx）
        manifest["guard"] = guard
        manifest["status"] = "PASS" if guard["status"] == "PASS" else "BLOCKED"
        manifest["verification"]["structural"] = "PASS" if guard["status"] == "PASS" else "BLOCK"
        _apply_delivery_gates(manifest, mode)
        _write_manifest_and_summary(manifest_path, manifest)
        print(json.dumps(_printable(result, mode, policy), ensure_ascii=False, indent=2))
        return policy.req("verification.exit_codes.ok") if manifest["status"] == "PASS" else exit_block
    if guard["status"] == "PASS" or args.force_compile:
        try:
            compile_report = compile_deck(deck, pptx_path,
                                          base_dir=(deck.get("meta") or {}).get("base_dir") or None)
        except CompileError as exc:
            raise InputError(f"编译失败：{exc}") from None
        rb = readback(pptx_path)
        fidelity = compare(deck, rb, policy)
        result["compile"] = compile_report
        result["fidelity"] = fidelity
        manifest["verification"]["structural"] = "PASS"
        manifest["verification"]["pptx_readback"] = fidelity["status"]
        if fidelity["status"] != "MATCH":
            guard_errors = [{"code": "LAYOUT_DRIFT", "detail": f"{m['kind']}: {m.get('slide')}/{m.get('element')} "
                                                             f"{m.get('detail','')}", "fix": ""}
                            for m in fidelity["mismatches"][:20]]
            result["status"] = "BLOCK"
            result["fidelity_block"] = guard_errors

        # 渲染对照：spec 与 PPTX 回读各画一遍，像素差只可能来自编译/字段丢失
        pages = _sample_pages(deck, policy, mode, args.pages)
        spec_pngs, rb_pngs, diffs = [], [], []
        for index in pages:
            spec_png = render_page(page_model_from_spec(deck, deck["slides"][index]),
                                   out_dir / "pages" / f"{index + 1:02d}-spec.png", policy, fontbook)
            rb_png = render_page(page_model_from_readback(deck, rb["slides"][index], deck["slides"][index]),
                                 out_dir / "pages" / f"{index + 1:02d}-pptx.png", policy, fontbook)
            diff = pixel_diff(spec_png, rb_png, policy)
            diff["page"] = deck["slides"][index]["id"]
            spec_pngs.append(spec_png)
            rb_pngs.append(rb_png)
            diffs.append(diff)
        limit = float(policy.num("thresholds.fidelity_pixel_mismatch_max"))
        worst = max((d["mismatch_ratio"] for d in diffs), default=0.0)
        result["pixel"] = {"pages": diffs, "worst_ratio": worst, "limit": limit,
                           "status": "MATCH" if worst <= limit else "DIVERGENT"}
        manifest["verification"]["spec_render"] = "PASS" if worst <= limit else "DIVERGENT"
        if worst > limit and mode == "release":
            result["status"] = "BLOCK"
        sheet = contact_sheet(spec_pngs, out_dir / "preview-spec.png", policy,
                              labels=[f"{i + 1:02d} {deck['slides'][i]['id']} · {deck['slides'][i].get('composition','')}"
                                      for i in pages])
        sheet_rb = contact_sheet(rb_pngs, out_dir / "preview-pptx.png", policy,
                                 labels=[f"{i + 1:02d} {deck['slides'][i]['id']} · readback" for i in pages])
        result["previews"] = {"spec_sheet": sheet, "pptx_sheet": sheet_rb,
                              "pages": spec_pngs + rb_pngs}
        manifest["previews"] = result["previews"]
        manifest["output_pptx"] = str(pptx_path.resolve())
        manifest["output_sha256"] = compile_report["output_sha256"]

    manifest["guard"] = guard
    manifest["fidelity"] = result.get("fidelity")
    manifest["pixel"] = result.get("pixel")
    manifest["status"] = "PASS" if result["status"] == "PASS" and _release_ready(result, mode, policy) else "BLOCKED"

    external_gates = None
    if args.evidence:
        from release import ReleaseEvidenceError, evaluate_evidence, load_evidence
        if not manifest.get("output_sha256"):
            raise InputError("没有生成 PPTX，不能接收外部 release evidence")
        try:
            evidence = load_evidence(args.evidence, SCHEMA_DIR)
            external_gates = evaluate_evidence(evidence, manifest["output_sha256"], deck)
        except ReleaseEvidenceError as exc:
            raise InputError(str(exc)) from None
        manifest["release_evidence_path"] = str(Path(args.evidence).expanduser().resolve())
    _apply_delivery_gates(manifest, mode, external_gates)
    _write_manifest_and_summary(manifest_path, manifest)

    printable = _printable(result, mode, policy)
    printable["delivery_status"] = manifest["delivery_status"]
    printable["quality_gates"] = manifest["quality_gates"]
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    if mode == "release" and manifest["delivery_status"] != "READY TO RELEASE":
        return exit_block
    return policy.req("verification.exit_codes.ok") if manifest["status"] == "PASS" else exit_block


def _write_manifest_and_summary(manifest_path: Path, manifest: dict) -> None:
    """Write the auditable manifest and its deterministic one-page delivery view."""
    from release import write_delivery_summary

    summary_path = manifest_path.parent / "delivery-summary.md"
    manifest["delivery_summary"] = str(summary_path.resolve())
    write_json(manifest_path, manifest)
    write_delivery_summary(manifest, summary_path)


def _apply_delivery_gates(manifest: dict, mode: str, external_gates: dict | None = None) -> None:
    """把工程结果映射到稳定质量门；未执行的人工/目标环境检查绝不推定为 PASS。"""
    verification = manifest.get("verification") or {}
    structural = "PASS" if manifest.get("status") == "PASS" and verification.get("structural") == "PASS" else "FAIL"
    readback = verification.get("pptx_readback")
    preview = verification.get("spec_render")
    external_gates = external_gates or {}
    deck_gate = external_gates.get("G-DECK", {"status": "NOT RUN", "evidence": "independent whole-deck visual/narrative review required"})
    target_gate = external_gates.get("G-TGT", {"status": "NOT VERIFIED", "evidence": verification.get("target_powerpoint")})
    semantic_pass = readback == "MATCH" and deck_gate.get("status") == "PASS"
    preview_pass = preview == "PASS" and target_gate.get("status") == "PASS"
    gates = {
        "G-STR": {"status": structural, "evidence": "guard + schema + PPTX compile/readback"},
        "G-SEM": {"status": "PASS" if semantic_pass else ("PARTIALLY VERIFIED" if readback == "MATCH" else "NOT RUN"),
                  "evidence": "PPTX object/payload fidelity + human claim/evidence review" if semantic_pass else "PPTX object/payload fidelity; claim/evidence human review not executed"},
        "G-PRE": {"status": "PASS" if preview_pass else ("PARTIALLY VERIFIED" if preview == "PASS" else "NOT RUN"),
                  "evidence": "internal readback preview + target PowerPoint visual verification" if preview_pass else "actual PPTX readback rendered by internal renderer; not PowerPoint rendering"},
        "G-DECK": deck_gate,
        "G-TGT": target_gate,
        "G-COMPLY": external_gates.get("G-COMPLY", {"status": "NOT RUN", "evidence": "asset authorization/privacy review not automated"}),
    }
    manifest["quality_gates"] = gates
    if mode != "release":
        manifest["delivery_status"] = "NOT ASSESSED — run release checks and external reviews"
        return
    manifest["delivery_status"] = (
        "READY TO RELEASE" if all(g["status"] == "PASS" for g in gates.values())
        else "BLOCK — Delivery"
    )


def _release_ready(result: dict, mode: str, policy) -> bool:
    if result["status"] != "PASS":
        return False
    if mode != "release":
        return True
    fidelity = result.get("fidelity") or {}
    pixel = result.get("pixel") or {}
    return fidelity.get("status") == "MATCH" and pixel.get("status") == "MATCH"


def _sample_pages(deck: dict, policy, mode: str, pages) -> list[int]:
    total = len(deck["slides"])
    if pages:
        chosen = [int(p) - 1 for p in str(pages).split(",") if p.strip().isdigit()]
        return [p for p in chosen if 0 <= p < total]
    if mode != "release":
        limit = int(policy.opt("render.ghost_max_pages", 8))
        step = max(1, total // limit)
        return sorted(set(list(range(0, total, step))[:limit] + [total - 1]))
    return list(range(total))


def _printable(result: dict, mode: str, policy) -> dict:
    guard = result["guard"]
    out = {
        "status": result["status"],
        "mode": mode,
        "errors": [{"code": e["code"], "slide": e.get("slide"), "element": e.get("element"),
                    "detail": e["detail"]} for e in guard["errors"][:20]],
        "warnings": [{"code": w["code"], "slide": w.get("slide"), "detail": w["detail"]}
                     for w in guard["warnings"][:20]],
        "fix_groups": [{"code": g["code"], "count": g["count"], "hint": g["hint"]}
                       for g in guard["fix_groups"]],
        "stats": guard["stats"],
        "degraded_measurement": guard["degraded"],
        "disclaimer": guard["disclaimer"],
    }
    if "fidelity" in result:
        out["fidelity"] = {"status": result["fidelity"]["status"],
                           "summary": result["fidelity"]["summary"]}
    if "pixel" in result:
        out["pixel"] = {"worst_ratio": result["pixel"]["worst_ratio"],
                        "limit": result["pixel"]["limit"], "status": result["pixel"]["status"]}
    if result.get("fidelity_block"):
        out["fidelity_mismatches"] = [e["detail"] for e in result["fidelity_block"][:10]]
    return out


# ──────────────────────────────────────────────────────────────────────────
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pvad.py", description="PPT Visual Art Director · plan/build/check")
    parser.add_argument("--policy-dir", help="覆盖策略目录（用于对照实验）")
    sub = parser.add_subparsers(dest="command", required=True)

    plan = sub.add_parser("plan", help="brief → plan.json（判断）")
    plan.add_argument("brief")
    plan.add_argument("--out", default="plan.json")
    plan.add_argument("--strict", action="store_true", help="有未决主张时返回非零")

    build = sub.add_parser("build", help="plan + brief → deck.json（几何）")
    build.add_argument("plan")
    build.add_argument("--brief", help="覆盖 plan 里记录的 brief 路径")
    build.add_argument("--out", default="deck.json")

    check = sub.add_parser("check", help="deck → 硬门 + PPTX + 回读对账 + 渲染对照")
    check.add_argument("deck")
    check.add_argument("output")
    check.add_argument("--mode", choices=["spec", "draft", "release"], default=None)
    check.add_argument("--out-dir", default=None)
    check.add_argument("--pages", help="只取这些页（1-based，逗号分隔）")
    check.add_argument("--force-compile", action="store_true", help="BLOCK 也编译，便于定位")
    check.add_argument("--evidence", help="绑定本次 PPTX SHA-256 的外部整稿/目标环境/合规证据 YAML/JSON")
    media = sub.add_parser("media", help="deck → 待生产素材的生产说明（提示词 / 比例 / 目标像素）")
    media.add_argument("deck")
    media.add_argument("--out", default=None, help="写入清单文件（默认只打印摘要）")
    handoff = sub.add_parser("handoff", help="plan → 人类可读的最小 THINK→BUILD 交接摘要（派生视图）")
    handoff.add_argument("plan")
    handoff.add_argument("--out", default="handoff.md")
    impact = sub.add_parser("impact", help="改稿前后 deck JSON → 变更摘要与保守验证范围")
    impact.add_argument("before")
    impact.add_argument("after")
    impact.add_argument("--out", default=None)
    package = sub.add_parser("package", help="manifest → Draft / Review / Release 交付包")
    package.add_argument("manifest")
    package.add_argument("--stage", choices=["draft", "review", "release"], required=True)
    package.add_argument("--out", required=True)
    return parser


def policy_usage_code() -> int:
    """策略都没加载成功时也要给出同一个 usage 码：退回规范默认值并说明。"""
    try:
        return int(load_policy().req("verification.exit_codes.usage"))
    except Exception:                     # noqa: BLE001 - 策略坏了才走到这里
        return 2


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        policy = load_policy(Path(args.policy_dir) if args.policy_dir else None)
    except PolicyError as exc:
        print(f"策略错误：{exc}", file=sys.stderr)
        return int(policy_usage_code())
    try:
        if args.command == "plan":
            return run_plan(args, policy)
        if args.command == "build":
            return run_build(args, policy)
        if args.command == "media":
            return run_media(args, policy)
        if args.command == "handoff":
            return run_handoff(args, policy)
        if args.command == "impact":
            return run_impact(args, policy)
        if args.command == "package":
            return run_package(args, policy)
        return run_check(args, policy)
    except InputError as exc:
        print(f"输入错误：{exc}", file=sys.stderr)
        return int(policy.req("verification.exit_codes.usage"))
    except KeyboardInterrupt:
        print("中断：不宣称任何结果", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
