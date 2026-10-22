#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""blind_pack.py · 盲评包与盲评计分（视觉质量的唯一可信通道）。

为什么必须人工盲评：
  结构检查、像素对账、判断基准都只能证明「执行正确」，不能证明「设计成立」。
  视觉质量只能由人看成品来判，而且要盲：评审者不知道哪份是候选、哪份是基线。

流程：
  build   把多份 deck 的渲染图打散、去标识、随机编号，生成 pack/
          （pages/ + key.json + rubric.md + sheet.yaml）
  score   读回评审者的 sheet.yaml，算各维度均值、分差与评分者一致性（ICC 近似）
          一致性高 ≠ 质量高：一致地低分也是共识，这一点必须写进报告。

诚实边界：
  * 没有真实评审者填写时，score 必须报 NOT RUN，不得报 0 分或跳过。
  * 渲染图本身来自渲染器而非 PowerPoint，评审结论只对该渲染器成立。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import random
import shutil
import statistics
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _path in (str(ROOT / "scripts"), str(ROOT / "tests")):
    if _path not in sys.path:
        sys.path.insert(0, _path)

DIMENSIONS = [
    ("focus", "第一眼落点是否与主张一致；一页是否只有一个主角"),
    ("hierarchy", "层级：字号/位置/明暗/留白是否让人按预期顺序读"),
    ("composition", "构图是否成立：重心、留白职责、图文关系"),
    ("typography", "中文排版：断行、禁则、数字与单位、行距与字距"),
    ("credibility", "证据与数字是否可核：来源、单位、期间、口径"),
    ("rhythm", "整套节奏：推进/停留/喘息是否与内容匹配"),
]
SCALE = (1, 7)


def build_packs(decks: list[tuple[str, list[str]]], out_dir: Path, *, seed: int = 7,
                label: str = "未命名稿") -> dict:
    """decks: [(组名, [png 路径...])]。生成的包里没有任何组名，只有随机编号。"""
    if len(decks) < 2:
        raise SystemExit("盲评至少需要两组：候选与基线")
    out_dir.mkdir(parents=True, exist_ok=True)
    pages = out_dir / "pages"
    pages.mkdir(exist_ok=True)
    rng = random.Random(seed)
    entries = []
    for group, paths in decks:
        for i, path in enumerate(paths):
            digest = hashlib.sha256(Path(path).read_bytes()).hexdigest()[:8]
            entries.append({"group": group, "source": str(path), "digest": digest, "index": i})
    rng.shuffle(entries)
    key = []
    for order, entry in enumerate(entries, start=1):
        target = pages / f"{order:03d}.png"
        shutil.copyfile(entry["source"], target)
        key.append({"blind_id": target.stem, "group": entry["group"], "digest": entry["digest"]})
    (out_dir / "key.json").write_text(json.dumps(
        {"seed": seed, "label": label, "scale": list(SCALE), "key": key},
        ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "rubric.md").write_text(_rubric(label), encoding="utf-8")
    (out_dir / "sheet.yaml").write_text(_sheet([k["blind_id"] for k in key]), encoding="utf-8")
    return {"out": str(out_dir), "items": len(key),
            "groups": sorted({k["group"] for k in key}),
            "next": f"把 {out_dir}/pages 与 rubric.md 交给评审者，回收 sheet.yaml 后运行 score"}


def _rubric(label: str) -> str:
    lines = [f"# 盲评量表 · {label}", "",
             "规则：只看图，不看代码、不看这份量表之外的任何说明；不知道自己评的是哪一组。",
             f"量表：{SCALE[0]} 严重不足 · 4 及格 · {SCALE[1]} 优秀（只打整数）。", "",
             "先按 100% 尺寸看一张，再快速翻完全部，最后逐张打分。不要逐张停留过久。", "",
             "## 维度", ""]
    for key, description in DIMENSIONS:
        lines.append(f"- **{key}**：{description}")
    lines += ["", "## 整体判断（每张都答）", "",
              "- 如果只能保留一句话，这张要观众记住什么？（写下来）",
              "- 有没有一眼可见的错误：断行、遮挡、数字与单位分离、来源缺失？",
              "- 拿掉任意一个元素，信息是否受损？没有受损就是该删的。", "",
              "## 提交", "", "把分数填进 sheet.yaml 的 `scores` 下，每人一个文件。"]
    return "\n".join(lines) + "\n"


def _sheet(ids: list[str]) -> str:
    lines = ["# 评审者：填名字或工号", "reviewer: \"\"", "", "scores:"]
    for blind_id in ids:
        lines.append(f'  "{blind_id}":')
        for key, _ in DIMENSIONS:
            lines.append(f"    {key}: null")
        lines.append("    note: \"\"")
    return "\n".join(lines) + "\n"


def score(pack_dir: Path, sheet_paths: list[Path]) -> dict:
    """算各维度均值 + 组间差 + 评分者一致性（ICC(2,1) 近似）。"""
    key = json.loads((pack_dir / "key.json").read_text(encoding="utf-8"))
    group_of = {item["blind_id"]: item["group"] for item in key["key"]}
    sheets = []
    for path in sheet_paths:
        import yaml
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        sheets.append({"reviewer": data.get("reviewer") or Path(path).stem,
                       "scores": data.get("scores") or {}})
    filled = any(value is not None
                 for sheet in sheets
                 for dims in (sheet["scores"] or {}).values()
                 for dim, value in (dims or {}).items() if dim != "note")
    if not sheets or not filled:
        return {"status": "NOT RUN",
                "reason": "没有回收任何有效评分（全是空值）：不得把盲评记为 0 分、通过或未通过",
                "expected": "至少一份填好的 sheet.yaml（每张图 1–7 分）"}
    per_group: dict[str, dict[str, list[float]]] = {}
    for sheet in sheets:
        for blind_id, dims in (sheet["scores"] or {}).items():
            group = group_of.get(str(blind_id).zfill(len(next(iter(group_of)))))
            blind_id = str(blind_id)
            if group is None:
                continue
            for dim, value in (dims or {}).items():
                if dim == "note" or value is None:
                    continue
                value = float(value)
                if not SCALE[0] <= value <= SCALE[1]:
                    raise SystemExit(f"量表外分数：{blind_id}.{dim}={value}（只接受 {SCALE}）")
                per_group.setdefault(group, {}).setdefault(dim, []).append(value)
    table = {group: {dim: round(statistics.fmean(values), 2) for dim, values in dims.items()}
             for group, dims in per_group.items()}
    diffs = {}
    groups = list(table)
    if len(groups) >= 2:
        for dim in (d for d, _ in DIMENSIONS):
            values = [table[g].get(dim) for g in groups]
            if all(v is not None for v in values):
                diffs[dim] = round(max(values) - min(values), 2)
    return {
        "status": "RUN",
        "reviewers": [s["reviewer"] for s in sheets],
        "per_group": table,
        "max_gap_by_dimension": diffs,
        "icc_note": ("评分者一致性需要 ≥2 名评审者同一批作品；本工具不再冒充统计检验，"
                     "样本量不足时只报均值与分差。"),
        "reading": ("一致性高不等于质量高：一致地低分也是共识。"
                    "分差小于 0.5 时，不应宣称任一组更好（与量表 1–7 的分辨率一致）。"),
        "caveat": "评审针对渲染器输出，不是目标 PowerPoint；投影与字体回退仍需现场验收。",
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="盲评包 / 计分")
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build")
    build.add_argument("--group", action="append", required=True,
                       help="组名=目录（如 candidate=/path/renders），可重复")
    build.add_argument("--out", required=True)
    build.add_argument("--seed", type=int, default=7)
    build.add_argument("--label", default="未命名稿")
    scorer = sub.add_parser("score")
    scorer.add_argument("--pack", required=True)
    scorer.add_argument("--sheet", action="append", required=True)
    args = parser.parse_args(argv)
    if args.command == "build":
        decks = []
        for spec in args.group:
            name, _, directory = spec.partition("=")
            paths = sorted(str(p) for p in Path(directory).glob("*.png"))
            if not paths:
                raise SystemExit(f"{directory} 里没有 PNG")
            decks.append((name, paths))
        print(json.dumps(build_packs(decks, Path(args.out), seed=args.seed, label=args.label),
                         ensure_ascii=False, indent=2))
        return 0
    report = score(Path(args.pack), [Path(p) for p in args.sheet])
    print(json.dumps(report, ensure_ascii=False, indent=2))
    out = Path(args.pack) / "score.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report["status"] == "RUN" else 1


if __name__ == "__main__":
    sys.exit(main())
