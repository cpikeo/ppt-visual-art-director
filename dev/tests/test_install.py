#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 安装与打包回归测试

只验证"别人照文档装完能不能真的用起来"这一客观事实：
- `pip install .` 能构建出 wheel（原仓库因 flat-layout 多顶层包直接失败）；
- wheel 内的模块可 import，`ppt-vad` 控制台入口可执行并返回正确退出码；
- `pyproject.toml` 声明的 Python 版本、依赖与 entry point 与代码实际入口一致。

不在联网受限或缺少 pip 的环境硬撑：直接 skip，并说明"未验证"而不是报"通过"。
"""

from __future__ import annotations

import os
import subprocess
import sys
import sysconfig
from pathlib import Path

import pytest

try:  # Python 3.11+ 标准库；3.10 需 tomli（CI 中按需安装）
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    try:
        import tomli as tomllib  # type: ignore[no-redef]
    except ModuleNotFoundError:
        pytest.skip("缺少 TOML 解析器（tomllib/tomli），打包契约验证未执行（UNVERIFIED）",
                    allow_module_level=True)

REPO_ROOT = Path(__file__).resolve().parents[2]

def _pip_available() -> bool:
    return subprocess.run([sys.executable, "-m", "pip", "--version"],
                          capture_output=True).returncode == 0


@pytest.fixture(scope="module")
def pyproject() -> dict:
    with open(REPO_ROOT / "pyproject.toml", "rb") as f:
        return tomllib.load(f)


def test_pyproject_declares_build_backend_and_console_entry(pyproject):
    """缺少 build-system 时 pip 会退回 legacy 构建，扁平布局必然失败。"""
    assert pyproject["build-system"]["build-backend"] == "setuptools.build_meta"
    assert pyproject["project"]["scripts"]["ppt-vad"] == "scripts.compile_pptx:main"
    assert pyproject["project"]["requires-python"].startswith(">=3.")
    assert any(dep.startswith("python-pptx") for dep in pyproject["project"]["dependencies"])


def test_console_entry_target_is_importable_and_callable():
    import importlib

    module_path, func = "scripts.compile_pptx", "main"
    module = importlib.import_module(module_path)
    assert callable(getattr(module, func)), "pyproject 声明的入口必须真实存在且可调用"


@pytest.mark.skipif(not _pip_available(), reason="环境无 pip：安装验证未执行（UNVERIFIED）")
def test_pip_install_builds_wheel_and_cli_runs(tmp_path):
    """端到端：构建 + 装到隔离路径 + 跑 CLI 编译一份最小 plan。"""
    target = tmp_path / "site"
    build = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--no-deps", "--no-build-isolation",
         "--target", str(target), str(REPO_ROOT)],
        capture_output=True, text=True, timeout=600,
    )
    assert build.returncode == 0, f"pip install . 失败：\n{build.stdout[-3000:]}\n{build.stderr[-3000:]}"
    module_file = target / "scripts" / "compile_pptx.py"
    assert module_file.is_file(), "wheel 内必须包含 scripts/compile_pptx.py"

    plan = {
        "deck_id": "install-check", "title": "安装验证", "audience": "开发",
        "goal": "确认入口可用", "core_claim": "装完就能编译",
        "pages": [{"id": "p1", "role": "open", "claim": "装完就能编译", "elements": [
            {"id": "t", "type": "title", "content": "打包入口可用",
             "x": 0.6, "y": 0.6, "w": 8, "h": 1.2, "font_size": 34, "color": "#111827"}]}],
    }
    plan_path = tmp_path / "plan.json"
    plan_path.write_text(__import__("json").dumps(plan, ensure_ascii=False), encoding="utf-8")
    out_path = tmp_path / "installed.pptx"

    scripts_dir = Path(sysconfig.get_path("scripts"))
    exe = "ppt-vad.exe" if os.name == "nt" else "ppt-vad"
    candidates = [target / "bin" / exe, target / exe, scripts_dir / exe]
    entry = next((c for c in candidates if c.is_file()), None)
    assert entry is not None, f"未找到控制台入口，检查过：{candidates}"

    env = dict(os.environ, PYTHONPATH=str(target))
    run = subprocess.run([str(entry), str(plan_path), str(out_path)],
                         capture_output=True, text=True, timeout=180, env=env)
    assert run.returncode == 0, f"CLI 执行失败：{run.stdout}\n{run.stderr}"
    assert out_path.is_file() and out_path.stat().st_size > 0
