#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""One-command internal preflight. External target/reviewer gates remain external.

重复执行契约：同一个 `--out` 可以反复跑。每次开始前只清理本工具自己产出的东西
（package/、pvad-skill.zip、manifest.json、各阶段日志），不碰目录里的其他文件；
这样既不会因为残留目录而误判 FAIL，也不会有上一轮的日志被当成这一轮的证据。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 本工具在 --out 目录里拥有的产物；重跑前清掉，其余文件保留
OWNED_ENTRIES = ("package", "pvad-skill.zip", "manifest.json")
OWNED_SUFFIXES = (".stdout.txt", ".stderr.txt")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def reset_outputs(out: Path) -> list[str]:
    """清掉上一轮预检自己的产物，保证可重复执行且证据不串轮。"""
    removed: list[str] = []
    for name in OWNED_ENTRIES:
        path = out / name
        if path.is_dir() and not path.is_symlink():
            shutil.rmtree(path)
            removed.append(name + "/")
        elif path.exists() or path.is_symlink():
            path.unlink()
            removed.append(name)
    for path in sorted(out.glob("*")):
        if path.is_file() and path.name.endswith(OWNED_SUFFIXES):
            path.unlink()
            removed.append(path.name)
    return removed


def git_state() -> dict:
    """预检绑定的代码版本。在创建输出目录之前调用，避免把本次产物算成未提交改动。"""
    def _git(*args: str) -> str:
        return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True).stdout.strip()

    try:
        commit = _git("rev-parse", "HEAD")
        if len(commit) != 40 or any(ch not in "0123456789abcdef" for ch in commit):
            raise ValueError("no git HEAD")
        return {"commit": commit, "dirty": bool(_git("status", "--porcelain"))}
    except Exception:  # 未安装 git / 非仓库：如实标注，不伪造
        return {"commit": "UNKNOWN", "dirty": None}


def run(name: str, command: list[str], out: Path) -> dict:
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
    (out / f"{name}.stdout.txt").write_text(result.stdout, encoding="utf-8")
    (out / f"{name}.stderr.txt").write_text(result.stderr, encoding="utf-8")
    return {"name": name, "command": command, "exit_code": result.returncode,
            "status": "PASS" if result.returncode == 0 else "FAIL"}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run unified PVAD internal release preflight")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    out = Path(args.out).expanduser()
    git = git_state()
    out.mkdir(parents=True, exist_ok=True)
    reset = reset_outputs(out)
    py = sys.executable
    stages = [
        ("compile", [py, "-m", "compileall", "-q", "scripts", "evals", "tests"]),
        ("schemas", [py, "-c", "import json,pathlib; from jsonschema import Draft202012Validator; [Draft202012Validator.check_schema(json.loads(p.read_text())) for p in pathlib.Path('schema').glob('*.json')]; print('PASS')"]),
        ("tests", [py, "-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py"]),
        ("judgment", [py, "evals/judgment_eval.py", "--baseline", "--quiet", "--report", str(out / "judgment.json")]),
        ("prompt", [py, "evals/prompt_eval.py", "--quiet", "--report", str(out / "prompt.json")]),
        ("typography", [py, "evals/typography_eval.py"]),
        ("package", [py, "scripts/package_skill.py", "--out", str(out / "package"), "--zip", str(out / "pvad-skill.zip")]),
    ]
    results = []
    for name, command in stages:
        result = run(name, command, out)
        results.append(result)
        if result["status"] == "FAIL":
            break
    manifest = {
        "schema": "pvad-release-preflight-v1",
        "created_at": datetime.now(timezone.utc).isoformat(),
        "python": sys.version,
        "git": git,
        "reset": reset,
        "stages": results,
        "status": "PASS" if len(results) == len(stages) and all(r["status"] == "PASS" for r in results) else "FAIL",
        "external_gates": {"target_powerpoint": "NOT VERIFIED", "independent_review": "NOT RUN", "identity_authentication": "NOT VERIFIED"},
    }
    zip_path = out / "pvad-skill.zip"
    if zip_path.is_file():
        manifest["package_zip_sha256"] = sha256(zip_path)
        manifest["package_zip_bytes"] = zip_path.stat().st_size
    else:
        # 打包阶段退出码为 0 却没有产物，同样算失败：不能只看退出码
        manifest["status"] = "FAIL"
        manifest["status_note"] = "打包阶段未产出 pvad-skill.zip"
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0 if manifest["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
