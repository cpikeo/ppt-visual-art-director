#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""One-command internal preflight. External target/reviewer gates remain external."""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


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
    out.mkdir(parents=True, exist_ok=True)
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
        "stages": results,
        "status": "PASS" if len(results) == len(stages) and all(r["status"] == "PASS" for r in results) else "FAIL",
        "external_gates": {"target_powerpoint": "NOT VERIFIED", "independent_review": "NOT RUN", "identity_authentication": "NOT VERIFIED"},
    }
    zip_path = out / "pvad-skill.zip"
    if zip_path.is_file():
        manifest["package_zip_sha256"] = sha256(zip_path)
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0 if manifest["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
