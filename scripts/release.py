#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""External release evidence, deterministic delivery summary and staged packages.

No gate becomes PASS from a declaration alone: evidence is schema-validated, bound to the
exact PPTX SHA-256, and checked against the actual deck where relevant.
"""
from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path

import yaml


class ReleaseEvidenceError(ValueError):
    pass


def load_evidence(path, schema_dir: Path) -> dict:
    source = Path(path).expanduser()
    if not source.is_file():
        raise ReleaseEvidenceError(f"找不到 release evidence：{source}")
    try:
        value = yaml.safe_load(source.read_text(encoding="utf-8"))
    except (OSError, yaml.YAMLError) as exc:
        raise ReleaseEvidenceError(f"release evidence 解析失败：{source} — {exc}") from None
    if not isinstance(value, dict):
        raise ReleaseEvidenceError("release evidence 顶层必须是对象")
    import jsonschema
    schema = json.loads((schema_dir / "release-evidence.schema.json").read_text(encoding="utf-8"))
    errors = sorted(jsonschema.Draft202012Validator(schema).iter_errors(value),
                    key=lambda e: list(e.absolute_path))
    if errors:
        details = [f"{'/'.join(map(str, e.absolute_path)) or '<root>'}: {e.message}" for e in errors[:12]]
        raise ReleaseEvidenceError("release evidence 不符合 schema：\n  - " + "\n  - ".join(details))
    return value


def evaluate_evidence(evidence: dict, artifact_sha256: str, deck: dict) -> dict:
    """Return independent gate results; absent records stay NOT RUN/NOT VERIFIED."""
    if evidence.get("artifact_sha256") != artifact_sha256:
        raise ReleaseEvidenceError("release evidence 的 artifact_sha256 与本次 PPTX 不一致；不得跨版本复用验收")
    return {
        "G-DECK": _evaluate_deck_review(evidence.get("deck_review")),
        "G-TGT": _evaluate_target(evidence.get("target_environment")),
        "G-COMPLY": _evaluate_compliance(evidence.get("compliance"), deck),
    }


def _evaluate_deck_review(record) -> dict:
    if not record:
        return {"status": "NOT RUN", "evidence": "whole-deck review record absent"}
    all_checked = all(record["checklist"].values())
    passable = record["status"] == "PASS" and all_checked and not record["must_fix"]
    return {
        "status": "PASS" if passable else "FAIL",
        "evidence": f"reviewer={record['reviewer']}; reviewed_at={record['reviewed_at']}; "
                    f"checklist={sum(record['checklist'].values())}/7; must_fix={len(record['must_fix'])}",
    }


def _evaluate_target(record) -> dict:
    if not record:
        return {"status": "NOT VERIFIED", "evidence": "target PowerPoint record absent"}
    all_actions = all(record["actions"].values())
    passable = record["status"] == "PASS" and all_actions and not record["issues"]
    return {
        "status": "PASS" if passable else "FAIL",
        "evidence": f"{record['application']} {record['application_version']} on {record['operating_system']}; "
                    f"reviewer={record['reviewer']}; actions={sum(record['actions'].values())}/8; "
                    f"issues={len(record['issues'])}",
    }


def _deck_assets(deck: dict) -> list[tuple[str, str]]:
    return [(slide.get("id", ""), str(element.get("source", "")))
            for slide in deck.get("slides", [])
            for element in slide.get("elements", [])
            if element.get("type") == "image"]


def _evaluate_compliance(record, deck: dict) -> dict:
    if not record:
        return {"status": "NOT RUN", "evidence": "asset/privacy compliance record absent"}
    actual = set(_deck_assets(deck))
    declared = {(item["slide_id"], item["source"]) for item in record["assets"]}
    no_assets_claim_valid = record["no_external_assets"] and not actual and not declared
    complete_assets = (not record["no_external_assets"] and actual == declared and bool(actual))
    rights_ok = all(item["intended_use_cleared"] and item["privacy_cleared"]
                    and item["professional_review"] != "fail"
                    and (not item["generated"] or item["professional_review"] == "pass")
                    for item in record["assets"])
    passable = (record["status"] == "PASS" and not record["issues"]
                and record["sensitive_content_reviewed"]
                and (no_assets_claim_valid or (complete_assets and rights_ok)))
    missing = sorted(actual - declared)
    extra = sorted(declared - actual)
    return {
        "status": "PASS" if passable else "FAIL",
        "evidence": f"reviewer={record['reviewer']}; actual_assets={len(actual)}; declared={len(declared)}; "
                    f"missing={len(missing)}; extra={len(extra)}; issues={len(record['issues'])}",
    }


def delivery_summary(manifest: dict) -> str:
    gates = manifest.get("quality_gates") or {}
    passed = [key for key, value in gates.items() if value.get("status") == "PASS"]
    partial = [key for key, value in gates.items() if value.get("status") in {"PARTIALLY VERIFIED", "NOT VERIFIED"}]
    not_run = [key for key, value in gates.items() if value.get("status") == "NOT RUN"]
    failed = [key for key, value in gates.items() if value.get("status") in {"FAIL", "BLOCK"}]
    guard = manifest.get("guard") or {}
    must_fix = [f"{e.get('code')}: {e.get('detail')}" for e in guard.get("errors", [])]
    if failed or must_fix:
        next_step = "修复失败门和 MUST FIX，并对同类页面扩检后重新生成证据。"
    elif manifest.get("delivery_status") != "READY TO RELEASE":
        next_step = "完成未通过或未执行的强制质量门，并提交绑定本 PPTX 散列的证据。"
    else:
        next_step = "按 Release 包交付；任何内容修改后重新运行全部强制门。"
    def items(values):
        return "、".join(values) if values else "无"
    lines = [
        "# 交付摘要",
        "",
        f"- 工程状态：**{manifest.get('status', 'UNKNOWN')}**",
        f"- 交付状态：**{manifest.get('delivery_status', 'NOT ASSESSED')}**",
        f"- 制品 SHA-256：`{manifest.get('output_sha256', '未生成')}`",
        f"- 已通过质量门：{items(passed)}",
        f"- 部分验证 / 未验证：{items(partial)}",
        f"- 未执行质量门：{items(not_run)}",
        f"- 失败质量门：{items(failed)}",
        f"- MUST FIX：{items(must_fix)}",
        f"- 目标环境：{(gates.get('G-TGT') or {}).get('evidence', '未提供')}",
        f"- 建议下一步：{next_step}",
        "",
        "> 本摘要由 manifest 确定性生成；不得手工覆盖门状态。",
    ]
    return "\n".join(lines) + "\n"


def write_delivery_summary(manifest: dict, path) -> str:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(delivery_summary(manifest), encoding="utf-8")
    return str(target)


def build_package(manifest_path, stage: str, out_dir) -> dict:
    source = Path(manifest_path).expanduser().resolve()
    manifest = json.loads(source.read_text(encoding="utf-8"))
    if stage not in {"draft", "review", "release"}:
        raise ReleaseEvidenceError(f"未知交付阶段：{stage}")
    if stage == "release" and manifest.get("delivery_status") != "READY TO RELEASE":
        raise ReleaseEvidenceError("Release 包要求 delivery_status=READY TO RELEASE；当前制品仍被阻断")
    target = Path(out_dir).expanduser()
    target.mkdir(parents=True, exist_ok=True)
    copied = []

    def copy_if(path, name=None):
        if not path:
            return
        item = Path(path)
        if item.is_file():
            destination = target / (name or item.name)
            shutil.copy2(item, destination)
            copied.append(destination)

    copy_if(source, "manifest.json")
    copy_if(manifest.get("output_pptx"), "deck.pptx")
    previews = manifest.get("previews") or {}
    if stage in {"review", "release"}:
        copy_if(previews.get("spec_sheet"), "preview-spec.png")
        copy_if(previews.get("pptx_sheet"), "preview-pptx.png")
    summary = target / "delivery-summary.md"
    summary.write_text(delivery_summary(manifest), encoding="utf-8")
    copied.append(summary)
    if stage == "review":
        checklist = target / "REVIEW_REQUIRED.md"
        checklist.write_text("# Review 包\n\n请完成整稿、目标环境及合规复核；不得把本包当作正式 Release。\n", encoding="utf-8")
        copied.append(checklist)
    checksums = target / "SHA256SUMS.txt"
    lines = [f"{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.name}" for p in sorted(copied)]
    checksums.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"stage": stage, "out_dir": str(target), "files": [p.name for p in copied] + [checksums.name]}
