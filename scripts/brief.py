#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""brief.py · 输入契约的读取、校验、归一化。

只做确定性工作：解析 → Schema 校验 → 补显式默认 → 计算指纹。
不做设计判断（判断在 judgment.py），不猜内容（缺什么记什么）。
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

BRIEF_SCHEMA = "brief.schema.json"
_BRIEF_KEYS = {"meta", "audience", "decision", "tension", "visual_world", "source",
               "unit", "period", "basis", "brand", "constraints", "slides"}


class BriefError(ValueError):
    pass


def read_brief(path: str | Path) -> tuple[dict, str]:
    """读 YAML/JSON brief。返回 (raw, sha256)。不可信输入只走数据格式，不执行代码。"""
    source = Path(path).expanduser()
    if not source.is_file():
        raise BriefError(f"找不到 brief：{source}")
    raw_bytes = source.read_bytes()
    if len(raw_bytes) > 2_000_000:
        raise BriefError(f"brief 过大（{len(raw_bytes)} 字节）：上限 2MB")
    text = raw_bytes.decode("utf-8")
    suffix = source.suffix.lower()
    try:
        if suffix in (".json",):
            value = json.loads(text)
        elif suffix in (".yml", ".yaml"):
            import yaml
            value = yaml.safe_load(text)
        else:
            raise BriefError(f"brief 只接受 .yaml/.yml/.json：{source}")
    except BriefError:
        raise
    except Exception as exc:
        raise BriefError(f"brief 解析失败：{source} — {exc}") from None
    if not isinstance(value, dict):
        raise BriefError(f"brief 顶层必须是对象：{source}")
    return value, hashlib.sha256(raw_bytes).hexdigest()


def validate(brief: dict, schema_dir: Path) -> list[str]:
    """Schema 校验；返回错误清单（空 = 通过）。"""
    import jsonschema
    schema = json.loads((Path(schema_dir) / BRIEF_SCHEMA).read_text(encoding="utf-8"))
    validator = jsonschema.Draft202012Validator(schema)
    errors = []
    for e in sorted(validator.iter_errors(brief), key=lambda e: list(e.absolute_path)):
        where = "/".join(str(p) for p in e.absolute_path) or "<root>"
        errors.append(f"{where}: {e.message}")
    return errors


def _clean(text) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def normalize(brief: dict) -> tuple[dict, list[dict]]:
    """归一化，并显式记录每一处「补的默认值」（不静默补）。"""
    notes: list[dict] = []
    out: dict = {}
    out["meta"] = {k: _clean(v) for k, v in (brief.get("meta") or {}).items() if v}
    for key in ("audience", "decision", "tension", "source", "unit", "period", "basis"):
        out[key] = _clean(brief.get(key))
        if not out[key]:
            notes.append({"field": key, "kind": "empty", "detail": f"{key} 未声明：判断会把它记为未断言/未决，不补编"})
    brand = dict(brief.get("brand") or {})
    out["brand"] = {"colors": dict(brand.get("colors") or {}), "fonts": dict(brand.get("fonts") or {})}
    constraints = dict(brief.get("constraints") or {})
    out["constraints"] = {
        "avoid": [str(x) for x in (constraints.get("avoid") or []) if str(x).strip()],
        "density": constraints.get("density") or "",
        "max_slides": constraints.get("max_slides"),
    }

    slides: list[dict] = []
    seen: set[str] = set()
    for index, raw in enumerate(brief.get("slides") or []):
        sid = _clean(raw.get("id")) or f"s{index + 1:02d}"
        if sid in seen:
            raise BriefError(f"页 id 重复：{sid}（id 必须全稿唯一，取证顺序依赖它）")
        seen.add(sid)
        slide = {"id": sid, "index": index}
        slide["title"] = _clean(raw.get("title"))
        slide["content"] = _clean(raw.get("content"))
        for key in ("claim", "insight", "source", "unit", "period", "basis", "notes"):
            slide[key] = _clean(raw.get(key))
        for key in ("role", "evidence", "composition", "density", "focus"):
            slide[key] = (raw.get(key) or "").strip()
        slide["metrics"] = [_clean_metric(m) for m in (raw.get("metrics") or [])]
        chart = raw.get("chart")
        slide["chart"] = _clean_chart(chart) if chart else None
        slide["asset"] = _clean_asset(raw.get("asset")) if raw.get("asset") else None
        if not slide["content"] and not slide["metrics"] and not slide["chart"]:
            notes.append({"field": f"slides[{index}].content", "kind": "absent",
                          "detail": f"{sid} 只有标题、没有证据：判断卡记 absent，不虚构正文"})
        slides.append(slide)
    out["visual_world"] = _clean(brief.get("visual_world"))
    out["slides"] = slides

    unknown = set(brief) - _BRIEF_KEYS
    if unknown:
        notes.append({"field": ",".join(sorted(unknown)), "kind": "unknown",
                      "detail": "Schema 已拒绝未知字段时不会走到这里；此提示仅供直接调用 normalize 的路径"})
    return out, notes


def _clean_asset(asset: dict) -> dict:
    """素材声明归一：safe_area 统一成 {x,y,width,height}（比例坐标），路径不在这里解析。"""
    out = dict(asset)
    area = asset.get("safe_area")
    if isinstance(area, (list, tuple)) and len(area) == 4:
        out["safe_area"] = {"x": float(area[0]), "y": float(area[1]),
                            "width": float(area[2]), "height": float(area[3])}
    elif isinstance(area, dict):
        out["safe_area"] = {k: float(area[k]) for k in ("x", "y", "width", "height")}
    elif area is None:
        out["safe_area"] = None
    out["path"] = _clean(asset.get("path"))
    for key in ("function", "fit", "text_color", "reason", "subject"):
        if key in out:
            out[key] = _clean(out[key])
    avoid = asset.get("avoid")
    out["avoid"] = [_clean(x) for x in avoid] if isinstance(avoid, (list, tuple)) else []
    return out


def _clean_metric(metric: dict) -> dict:
    value = metric.get("value")
    if isinstance(value, str):
        value = value.strip()
    return {"value": value, "unit": _clean(metric.get("unit")), "label": _clean(metric.get("label")),
            "desc": _clean(metric.get("desc")), "period": _clean(metric.get("period"))}


def _clean_chart(chart: dict) -> dict:
    return {
        "kind": chart.get("kind"),
        "categories": [_clean(c) for c in (chart.get("categories") or [])],
        "series": [{"name": _clean(s.get("name")), "values": [float(v) for v in (s.get("values") or [])]}
                   for s in (chart.get("series") or [])],
        "takeaway": _clean(chart.get("takeaway")),
        "unit": _clean(chart.get("unit")),
    }


def load(path: str | Path, policy, schema_dir: Path) -> tuple[dict, str]:
    """读 + 校验 + 归一化。校验失败直接抛（不降级运行）。"""
    raw, digest = read_brief(path)
    errors = validate(raw, schema_dir)
    if errors:
        raise BriefError("brief 校验失败：\n" + "\n".join(f"  - {e}" for e in errors[:12]))
    normalized, notes = normalize(raw)
    normalized["_brief_sha256"] = digest
    normalized["_brief_notes"] = notes
    return normalized, digest
