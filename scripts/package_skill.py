#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a clean, allowlisted skill distribution and run smoke checks inside it."""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOP_FILES = ("SKILL.md", "VERSION", "CHANGELOG.md", "requirements.txt")
RUNTIME_DIRS = ("policy", "references", "schema", "scripts", "governance")
OPTIONAL_DIRS = ("tests", "evals")
EXCLUDED_PARTS = {"__pycache__", ".pytest_cache", "reports"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


class PackageSkillError(ValueError):
    pass


def _files(include_tests: bool) -> list[Path]:
    paths = []
    for name in TOP_FILES:
        path = ROOT / name
        if not path.is_file():
            raise PackageSkillError(f"缺少发布必需文件：{name}")
        paths.append(path)
    directories = RUNTIME_DIRS + (OPTIONAL_DIRS if include_tests else ())
    for directory in directories:
        base = ROOT / directory
        if not base.is_dir():
            raise PackageSkillError(f"缺少发布必需目录：{directory}")
        for path in base.rglob("*"):
            if (path.is_file() and not path.is_symlink()
                    and not (set(path.relative_to(ROOT).parts) & EXCLUDED_PARTS)
                    and path.suffix not in EXCLUDED_SUFFIXES):
                paths.append(path)
    return sorted(set(paths), key=lambda path: str(path.relative_to(ROOT)))


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_package(out_dir, *, include_tests: bool = False) -> dict:
    target = Path(out_dir).expanduser()
    if target.exists() and (target.is_symlink() or any(target.iterdir())):
        raise PackageSkillError(f"输出目录必须不存在或为空：{target}")
    target.mkdir(parents=True, exist_ok=True)
    source_files = _files(include_tests)
    entries = []
    for source in source_files:
        relative = source.relative_to(ROOT)
        if relative.parts[0] == "audit":
            raise PackageSkillError("audit 不得进入技能运行包")
        destination = target / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        if _sha(source) != _sha(destination):
            raise PackageSkillError(f"复制后散列不一致：{relative}")
        entries.append({"path": relative.as_posix(), "sha256": _sha(destination),
                        "bytes": destination.stat().st_size})
    manifest = {
        "schema": "pvad-skill-package-v1",
        "version": (target / "VERSION").read_text(encoding="utf-8").strip(),
        "include_tests_and_evals": include_tests,
        "files": entries,
    }
    manifest_path = target / "PACKAGE-MANIFEST.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _smoke(target, include_tests=include_tests)
    return {"out_dir": str(target.resolve()), "version": manifest["version"],
            "files": len(entries) + 1, "bytes": sum(item["bytes"] for item in entries),
            "include_tests_and_evals": include_tests, "smoke": "PASS"}


def build_deterministic_zip(source_dir, zip_path) -> dict:
    source = Path(source_dir).resolve()
    target = Path(zip_path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted((p for p in source.rglob("*") if p.is_file()), key=lambda p: p.relative_to(source).as_posix()):
            relative = path.relative_to(source).as_posix()
            info = zipfile.ZipInfo(relative, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o100644 & 0xFFFF) << 16
            info.create_system = 3
            archive.writestr(info, path.read_bytes(), compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    return {"zip": str(target.resolve()), "zip_sha256": _sha(target), "zip_bytes": target.stat().st_size}


def _smoke(target: Path, *, include_tests: bool) -> None:
    commands = [
        [sys.executable, "scripts/pvad.py", "--help"],
        [sys.executable, "-c",
         "import json,pathlib; from jsonschema import Draft202012Validator; "
         "[Draft202012Validator.check_schema(json.loads(p.read_text())) for p in pathlib.Path('schema').glob('*.json')]; "
         "print('schemas PASS')"],
    ]
    if include_tests:
        commands.append([sys.executable, "-m", "unittest", "discover", "-s", "tests", "-p", "test_documentation.py"])
    for command in commands:
        result = subprocess.run(command, cwd=target, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise PackageSkillError(f"干净发布包 smoke 失败：{' '.join(command)}\n{result.stderr}\n{result.stdout}")
    if (target / "audit").exists():
        raise PackageSkillError("干净发布包中发现 audit/")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Build clean PVAD skill package")
    parser.add_argument("--out", required=True)
    parser.add_argument("--include-tests", action="store_true", help="附带 tests/ 与 evals/（仍排除生成 reports）")
    parser.add_argument("--zip", dest="zip_path", help="可选：同时生成固定时间戳、固定权限、排序稳定的 ZIP")
    args = parser.parse_args(argv)
    try:
        result = build_package(args.out, include_tests=args.include_tests)
        if args.zip_path:
            result.update(build_deterministic_zip(args.out, args.zip_path))
    except PackageSkillError as exc:
        print(f"技能包生成失败：{exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
