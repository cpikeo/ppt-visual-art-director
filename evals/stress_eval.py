#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Measured 20/50-slide engineering stress evaluation; not a visual-quality score."""
from __future__ import annotations

import argparse
import json
import resource
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import yaml
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
PVAD = ROOT / "scripts" / "pvad.py"


def _slide(index: int) -> dict:
    slide = {
        "id": f"s{index:02d}",
        "title": f"第 {index} 页的运营效率必须继续提升",
        "claim": f"第 {index} 页的运营效率必须继续提升",
        "content": f"本页记录第 {index} 个经确认的运营判断，当前指标为 {100 + index} 万元。",
        "source": "压力测试合成数据（非业务事实）",
        "unit": "万元",
        "period": "2026 Q1",
        "basis": "合成口径",
        "role": "explain",
    }
    if index % 10 == 0:
        slide["chart"] = {
            "kind": "column", "categories": ["Q1", "Q2", "Q3"],
            "series": [{"name": "合成指标（万元）", "values": [index, index + 2, index + 4]}],
            "takeaway": "仅用于压力测试",
        }
    if index in {5, 15, 25, 35, 45}:
        slide["asset"] = {
            "path": "assets/large-reused.jpg", "function": "context",
            "reason": "验证大尺寸素材解码与同职能复用，不承载事实证据",
            "subject": "合成的中性几何纹理", "fit": "cover",
        }
    return slide


def _run(command: list[str], *, expected: set[int]) -> dict:
    started = time.perf_counter()
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, timeout=900)
    elapsed = round(time.perf_counter() - started, 4)
    if result.returncode not in expected:
        raise RuntimeError(f"command failed {result.returncode}: {' '.join(command)}\n{result.stderr}\n{result.stdout}")
    return {"seconds": elapsed, "exit": result.returncode,
            "stdout_tail": result.stdout[-500:], "stderr_tail": result.stderr[-500:]}


def evaluate(slides: int) -> dict:
    work = Path(tempfile.mkdtemp(prefix=f"pvad-stress-{slides}-"))
    try:
        assets = work / "assets"
        assets.mkdir()
        # 6 MP source exercises decode/crop without breaching the policy's hard decode budget.
        Image.new("RGB", (3000, 2000), (184, 191, 196)).save(assets / "large-reused.jpg", quality=90)
        brief = {
            "meta": {"title": f"{slides}页压力测试", "lang": "zh"},
            "audience": "工程回归",
            "decision": "确认长稿工程链稳定",
            "tension": "长稿、图表与大素材可能增加耗时和内存",
            "source": "压力测试合成数据（非业务事实）",
            "unit": "万元", "period": "2026 Q1", "basis": "合成口径",
            "slides": [_slide(i) for i in range(1, slides + 1)],
        }
        brief_path = work / "brief.yaml"
        brief_path.write_text(yaml.safe_dump(brief, allow_unicode=True, sort_keys=False), encoding="utf-8")
        plan, deck, pptx, report = work / "plan.json", work / "deck.json", work / "output.pptx", work / "report"
        commands = {}
        commands["plan"] = _run([sys.executable, str(PVAD), "plan", str(brief_path), "--out", str(plan)], expected={0})
        commands["build"] = _run([sys.executable, str(PVAD), "build", str(plan), "--brief", str(brief_path),
                                  "--out", str(deck)], expected={0})
        commands["release_check"] = _run([sys.executable, str(PVAD), "check", str(deck), str(pptx),
                                          "--mode", "release", "--out-dir", str(report)], expected={1})
        package = work / "review-package"
        commands["review_package"] = _run([sys.executable, str(PVAD), "package", str(report / "manifest.json"),
                                           "--stage", "review", "--out", str(package)], expected={0})
        manifest = json.loads((report / "manifest.json").read_text(encoding="utf-8"))
        deck_value = json.loads(deck.read_text(encoding="utf-8"))
        return {
            "slides": slides,
            "commands": commands,
            "max_child_rss_kb": resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss,
            "pptx_bytes": pptx.stat().st_size,
            "deck_json_bytes": deck.stat().st_size,
            "elements": sum(len(slide["elements"]) for slide in deck_value["slides"]),
            "images": sum(1 for slide in deck_value["slides"] for element in slide["elements"] if element["type"] == "image"),
            "charts": sum(1 for slide in deck_value["slides"] for element in slide["elements"] if element["type"] == "chart"),
            "engineering_status": manifest["status"],
            "delivery_status": manifest["delivery_status"],
            "pixel_status": (manifest.get("pixel") or {}).get("status"),
            "fidelity_status": (manifest.get("fidelity") or {}).get("status"),
            "target_powerpoint_seconds": "未测量",
            "external_blind_review": "未测量",
        }
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=str(ROOT / "evals" / "reports" / "stress-report.json"))
    parser.add_argument("--sizes", default="20,50")
    args = parser.parse_args(argv)
    sizes = [int(item) for item in args.sizes.split(",") if item.strip()]
    report = {
        "schema": "pvad-stress-report-v1",
        "note": "合成工程压力测试；不代表视觉质量、真实业务内容或目标 PowerPoint 性能。",
        "environment": {"python": sys.version.split()[0], "platform": sys.platform},
        "runs": [evaluate(size) for size in sizes],
    }
    target = Path(args.out)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
