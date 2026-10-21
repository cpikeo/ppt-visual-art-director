#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试夹具：路径、临时策略副本、最小 brief、跑 CLI 的小工具。"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
for path in (str(SCRIPTS),):
    if path not in sys.path:
        sys.path.insert(0, path)

EXAMPLES = ROOT / "examples"
FIXTURES = ROOT / "tests" / "fixtures"

MINI_BRIEF = """\
audience: 内部评审会
decision: 批准下一步预算
tension: 团队对渠道是否继续投入没有共识
source: 内部经营看板
unit: 万元
period: 2026 Q1–Q3
basis: 已开票收入
slides:
  - id: s01
    title: 加店不再带来增长
    content: |-
      门店数量增长 18%，单店月均产出下降 21%。
    source: 内部经营看板
  - id: s02
    title: 经销渠道的单店产出高出直营门店 2.4 倍
    content: |-
      4.6
      经销渠道单店月均产出
      万元
      1.9
      直营门店单店月均产出
      万元
    source: 渠道月度结算表
    unit: 万元
  - id: s03
    title: 经销渠道收入占比逐季抬升
    content: |-
      2026 Q1 到 Q3，经销渠道收入占比从 34% 升到 51%。
    source: 渠道月度结算表
    chart:
      kind: column
      categories: ["2026 Q1", "2026 Q2", "2026 Q3"]
      series:
        - name: 经销渠道收入占比
          values: [34, 43, 51]
  - id: s04
    title: 请批准 Q1 预算从直营门店转投经销渠道
    content: |-
      直营门店的边际产出已经低于经销渠道。
    source: 财务部
"""


def tmpdir() -> Path:
    return Path(tempfile.mkdtemp(prefix="pvad-test-"))


def copy_policy(tmp: Path) -> Path:
    target = tmp / "policy"
    shutil.copytree(ROOT / "policy", target)
    return target


def write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def write_brief(tmp: Path, name: str = "brief.yaml", text: str | None = None) -> Path:
    return write(tmp / name, text if text is not None else MINI_BRIEF)


def run_cli(*args, env=None) -> subprocess.CompletedProcess:
    command = [sys.executable, str(SCRIPTS / "pvad.py"), *[str(a) for a in args]]
    return subprocess.run(command, capture_output=True, text=True, env=env,
                          cwd=str(ROOT), timeout=300)


def load_json(path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def pipeline(tmp: Path, brief_text: str | None = None, mode: str = "draft"):
    """plan → build → check 一趟；返回 (plan, deck, report, paths)。"""
    brief = write_brief(tmp, text=brief_text)
    plan_path, deck_path = tmp / "plan.json", tmp / "deck.json"
    pptx = tmp / "out.pptx"
    r1 = run_cli("plan", brief, "--out", plan_path)
    assert r1.returncode == 0, r1.stderr
    r2 = run_cli("build", plan_path, "--brief", brief, "--out", deck_path)
    assert r2.returncode == 0, r2.stderr
    r3 = run_cli("check", deck_path, pptx, "--mode", mode, "--out-dir", tmp / "report")
    report = load_json(tmp / "report" / "manifest.json")
    return load_json(plan_path), load_json(deck_path), report, {"pptx": pptx, "report": tmp / "report",
                                                              "exit": r3.returncode, "stderr": r3.stderr}
