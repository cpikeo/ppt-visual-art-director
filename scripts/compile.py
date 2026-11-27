#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plan.json → 原生可编辑 PPTX（V22：PptxGenJS 主引擎）。

用法（命令接口不变）：
    python scripts/compile.py plan.json out.pptx [--assets DIR] [--report report.json] [--force] [--production]

流水线：Plan 校验与设计信号（checks.py）→ spec.py 解析为 render spec（唯一的设计计算处）
        → render.js（PptxGenJS 生成原生对象）→ post.py（OOXML 补齐 + 图表 XML 兼容性自检 + 可复现打包）
        → 包检查 → 原子落盘。

只做确定性执行：校验契约、按 plan 建对象、写出文件。不自动缩字、换色、补数据或联网取图。
硬错误（E-）在写文件前阻断；输出与报告都不得指向 plan 自己。目标文件已存在时默认不动，确认要替换才加 --force。
先写临时文件再原子替换，失败不留半成品；运行期异常另写 <名称>.error.pptx 供排查，该文件不是交付物。
`--engine legacy`：迁移期回滚通道（python-pptx 旧引擎，仅在仓库完整检出时存在，不随正式包分发）。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from checks import design_signals, validate
from model import (CANVAS_EMU, Report, canvas, default_asset_root, guard_paths, load_plan, package_problems, page_views,
                   place_atomic, show, temp_beside, utf8_stdio, write_atomic)
import post
import spec as specmod

HERE = Path(__file__).resolve().parent
NODE_TIMEOUT_S = 180


class EngineBlock(Exception):
    """PptxGenJS 引擎明确不能可靠生成的 Plan 能力：阻断，不静默降级。"""


def render_pptx(plan: Dict[str, Any], views: List[Dict[str, Any]], asset_root: Path, out: Path, force: bool) -> Dict[str, Any]:
    """spec → PptxGenJS → 后处理 → 包检查 → 原子落盘。"""
    spec = specmod.build_spec(plan, views, asset_root)
    if not (HERE.parent / "node_modules" / "pptxgenjs").is_dir():
        raise RuntimeError("缺少 Node 依赖：在 Skill 根目录执行 `npm ci`（需要 Node ≥ 18；版本已锁定在 package-lock.json）")
    with tempfile.TemporaryDirectory(prefix="vad-") as td:
        spec_path, raw = Path(td) / "spec.json", Path(td) / "raw.pptx"
        spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        proc = subprocess.run(["node", str(HERE / "render.js"), str(spec_path), str(raw)], capture_output=True, text=True,
                              timeout=NODE_TIMEOUT_S, cwd=str(HERE.parent))
        if proc.returncode != 0:
            msg = (proc.stderr or proc.stdout).strip()
            if "BLOCK " in msg:
                raise EngineBlock(msg.split("BLOCK ", 1)[1].splitlines()[0])
            raise RuntimeError("render.js 失败：" + msg[-600:])
        tmp = temp_beside(out)
        try:
            try:
                stats = post.postprocess(raw, spec, tmp)
            except post.PostError as exc:
                raise PackageError(str(exc))
            problems = package_problems(tmp, check_suffix=False)   # 先验临时件再落盘：不合格的包不得以交付名出现
            if problems:
                raise PackageError("；".join(problems))
            place_atomic(tmp, out, force)
        finally:
            if tmp.exists():
                tmp.unlink()
    return stats


class PackageError(Exception):
    """成品包不合格（内容类型、宏等）：交付路径未写。"""


def error_deck_path(out: Path) -> Path:
    return out.with_name(out.stem + ".error.pptx")


def write_error_deck(out: Path, message: str, force: bool = False) -> None:
    """诊断稿：同样先写临时文件再原子替换，写到一半失败不留半成品。"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = CANVAS_EMU
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(5.9))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = "编译失败，此文件不是交付物。\n\n" + message[:1500]
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(14 if run.text.startswith("编译失败") else 11)
    tmp = temp_beside(out)
    try:
        prs.save(str(tmp))
        place_atomic(tmp, out, force)
    finally:
        if tmp.exists():
            tmp.unlink()


def diagnostic_target(plan_path: Path, out_path: Path, report_path: Optional[Path], force: bool) -> Tuple[Optional[Path], str]:
    """诊断稿能不能写、写到哪：与交付物、plan、报告走同一套 guard_paths 判定。

    返回 (路径或 None, 说明)。不得覆盖 plan / 交付物 / 报告，不得经符号链接或写进目录；
    已存在且没有 --force 时不覆盖。
    """
    target = error_deck_path(out_path)
    probe = Report()
    inputs: Dict[str, Optional[Path]] = {"plan": plan_path, "输出文件": out_path}
    if report_path is not None:
        inputs["--report"] = report_path
    guard_paths(probe, inputs, [("诊断文件", target, ".pptx")], force=True)
    if probe.blocked:
        return None, "诊断文件未写：" + "；".join(i.message for i in probe.items if i.level == "HARD")
    if target.exists() and not force:
        return None, f"诊断文件 {target} 已存在，未覆盖（加 --force 才写）"
    return target, f"诊断文件：{target}（不是交付物）"


def main(argv: Optional[List[str]] = None) -> int:
    utf8_stdio()
    parser = argparse.ArgumentParser(description="plan.json → 原生可编辑 PPTX")
    parser.add_argument("plan")
    parser.add_argument("out")
    parser.add_argument("--assets", default=None, help="素材目录；缺省为 plan 同级 assets/ 或 plan 所在目录")
    parser.add_argument("--report", default=None, help="把诊断写成 JSON")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    parser.add_argument("--engine", choices=["pptxgenjs", "legacy"], default="pptxgenjs", help="生成引擎；legacy 仅为迁移期回滚通道")
    parser.add_argument("--production", action="store_true", help="生产模式：设计判断仍是待判断（draft / 占位符 / 缺构图理由）即阻断；默认草稿模式只提示")
    args = parser.parse_args(argv)
    if args.engine == "legacy":
        legacy = HERE / "legacy" / "compile_legacy.py"
        if not legacy.exists():
            print("[HARD] E-ENGINE legacy 引擎不在正式包内；请使用完整仓库检出")
            return 2
        sys.path[:0] = [str(HERE), str(HERE / "legacy")]
        import compile_legacy
        return compile_legacy.main([a for a in (argv if argv is not None else sys.argv[1:]) if a not in ("--engine", "legacy")])

    plan_path = Path(args.plan).resolve()
    out_path = Path(args.out).resolve()
    asset_root = default_asset_root(plan_path, args.assets)

    rep = Report()
    report_path = Path(args.report).resolve() if args.report else None
    # 源文件与交付物分家：路径判定统一走 model.guard_paths（输出/报告不得覆盖 plan、素材，报告只能是 .json）。
    guard_paths(rep, {"plan": plan_path}, [("输出文件", Path(args.out), ".pptx"), ("--report", Path(args.report) if args.report else None, ".json")], args.force)
    plan = load_plan(plan_path, rep)
    if plan is None or rep.blocked:
        show(rep.items)
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        return 2

    try:
        views = page_views(plan)   # 遍历与几何解析一次，validate / design_signals / build 共用（plan 在此之后不再被修改）
        rep.extend(validate(plan, asset_root, args.production, views))
        rep.extend(design_signals(plan, views))
    except Exception as exc:   # 校验层遇到未预料的输入：按文档化退出码 3 报告，不抛 Traceback、不写文件
        print(f"[HARD] E-RUNTIME 校验期失败：{type(exc).__name__}: {str(exc)[:200]}（plan 里可能有非法数值或结构；未写出文件）")
        return 3
    show(rep.items)
    if rep.blocked:
        print(f"\n阻断：{rep.count('HARD')} 个硬错误，未写出文件。")
        if args.report:
            write_atomic(Path(args.report), json.dumps({"blocked": True, "issues": rep.to_dict()},
                                                       ensure_ascii=False, indent=2).encode("utf-8"), args.force)
        return 2

    try:
        render_pptx(plan, views, asset_root, out_path, args.force)
    except EngineBlock as exc:
        print(f"[HARD] E-ENGINE 引擎阻断（不静默降级）：{exc}")
        return 2
    except FileExistsError:
        print(f"[HARD] E-EXIST {out_path} 在编译期间已被创建；未覆盖（确认要替换再加 --force）")
        return 2
    except PackageError as exc:
        print(f"[HARD] E-FORMAT 成品包不合格，交付路径未写：{exc}")
        return 3
    except subprocess.TimeoutExpired:
        print(f"[HARD] E-RUNTIME render.js 超过 {NODE_TIMEOUT_S}s 未返回，已终止；未写出文件")
        return 3
    except Exception as exc:  # 运行期异常：不覆盖正式输出
        error_path, note = diagnostic_target(plan_path, out_path, report_path, args.force)
        if error_path is not None:
            try:
                write_error_deck(error_path, f"{type(exc).__name__}: {exc}", args.force)
            except Exception as diag_exc:  # 诊断稿本身写失败不能掩盖原始错误
                note = f"诊断文件写入失败：{type(diag_exc).__name__}: {diag_exc}"
        print(f"[HARD] E-RUNTIME 运行期失败：{type(exc).__name__}: {str(exc)[:200]}\n{note}")
        return 3

    print(f"\n已写出 {out_path}（{len(plan['pages'])} 页，{canvas()[0]:.3f}×{canvas()[1]:.3f} in = 1280×720 px）"
          f"  警告 {rep.count('SOFT')}，提示 {rep.count('INFO')}")
    if args.report:
        write_atomic(Path(args.report), json.dumps({
            "blocked": False,
            "output": str(out_path),
            "pages": len(plan["pages"]),
            "sha256": hashlib.sha256(out_path.read_bytes()).hexdigest(),
            "issues": rep.to_dict(),
        }, ensure_ascii=False, indent=2).encode("utf-8"), args.force)
    return 0

if __name__ == "__main__":
    sys.exit(main())
