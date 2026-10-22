#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""policy.py · 结构化策略的唯一入口（读 + 校验 + 访问）。

设计约束（对应 SKILL「五层架构与权威源」）：
  * 本模块不含业务阈值。所有数值来自 policy/strategy.yaml 与 policy/composition.yaml；
    访问不存在的键必须显式给默认值，否则抛 PolicyError —— 不允许隐式默认值藏在代码里。
  * 加载即校验：JSON Schema 管形状，validate_semantics 管引用完整性。
  * 策略文件带 SHA-256：判断与验收报告记录该指纹，回放时能确认「同一套策略」。
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path

POLICY_DIR = Path(__file__).resolve().parent.parent / "policy"
SCHEMA_DIR = Path(__file__).resolve().parent.parent / "schema"

_MISSING = object()


class PolicyError(ValueError):
    """策略文件缺失、结构非法或引用断裂。"""


def _read_yaml(path: Path) -> dict:
    import yaml
    try:
        value = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise PolicyError(f"策略文件 YAML 解析失败：{path} — {exc}") from None
    if not isinstance(value, dict):
        raise PolicyError(f"策略文件顶层必须是对象：{path}")
    return value


def _schema(name: str) -> dict:
    path = SCHEMA_DIR / name
    if not path.is_file():
        raise PolicyError(f"缺少 Schema：{path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _validate(instance: dict, schema_name: str, label: str) -> None:
    try:
        import jsonschema
    except ImportError:  # pragma: no cover - 环境缺失时明确降级，不静默跳过
        raise PolicyError("校验策略需要 jsonschema：pip install jsonschema") from None
    validator = jsonschema.Draft202012Validator(_schema(schema_name))
    errors = sorted(validator.iter_errors(instance), key=lambda e: list(e.absolute_path))
    if errors:
        lines = [f"  - {label}: {'/'.join(str(p) for p in e.absolute_path) or '<root>'}: {e.message}"
                 for e in errors[:12]]
        raise PolicyError("策略结构非法：\n" + "\n".join(lines))


def _digest(*paths: Path) -> str:
    h = hashlib.sha256()
    for path in sorted(paths):
        h.update(path.name.encode("utf-8"))
        h.update(path.read_bytes())
    return h.hexdigest()


@dataclass(frozen=True)
class Policy:
    strategy: dict
    composition: dict
    sha256: str
    paths: tuple
    prompts: dict | None = None

    # ── 访问 ────────────────────────────────────────────────────────────
    def req(self, dotted: str):
        """必填键；缺失即错误（防止隐式默认值）。"""
        value = self._walk(self.strategy, dotted)
        if value is _MISSING:
            raise PolicyError(f"策略缺键：{dotted}（请在 policy/strategy.yaml 显式给出）")
        return value

    def opt(self, dotted: str, default):
        """可选键，但调用点必须显式声明默认值。"""
        value = self._walk(self.strategy, dotted)
        return default if value is _MISSING else value

    def num(self, dotted: str) -> float:
        value = self.req(dotted)
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            raise PolicyError(f"策略键 {dotted} 应为数值，实际 {type(value).__name__}")
        return value

    def enum(self, name: str) -> tuple:
        values = (self.strategy.get("enums") or {}).get(name)
        if not values:
            raise PolicyError(f"enums 缺枚举：{name}")
        return tuple(values)

    def is_enum(self, name: str, value) -> bool:
        return value in self.enum(name)

    # 排版语义阶梯
    def size(self, key: str) -> float:
        return self._ladder("scale", key)

    def floor(self, key: str) -> float:
        return self._ladder("floors", key)

    def line_height(self, key: str) -> float:
        return self._ladder("line_height", key)

    def tracking(self, key: str) -> float:
        return float(self.req(f"typography.tracking_pt.{key}"))

    def max_lines(self, key: str) -> int:
        value = self.req(f"typography.max_lines.{key}")
        return int(value)

    def recipe(self, name: str) -> dict:
        recipes = self.composition.get("recipes") or {}
        if name not in recipes:
            raise PolicyError(f"缺少构图配方：{name}（enums.composition 与 composition.yaml 必须对齐）")
        return recipes[name]

    def colors(self) -> dict:
        return dict(self.req("colors.defaults"))

    # ── 内部 ────────────────────────────────────────────────────────────
    def _ladder(self, section: str, key: str) -> float:
        value = self.req(f"typography.{section}.{key}")
        return float(value)

    @staticmethod
    def _walk(root, dotted: str):
        node = root
        for part in dotted.split("."):
            if not isinstance(node, dict) or part not in node:
                return _MISSING
            node = node[part]
        return node


def load(policy_dir: Path | None = None, *, validate: bool = True) -> Policy:
    """加载策略。默认路径为技能包内 policy/。"""
    base = Path(policy_dir) if policy_dir else POLICY_DIR
    strategy_path = base / "strategy.yaml"
    composition_path = base / "composition.yaml"
    prompts_path = base / "prompts.yaml"
    for path in (strategy_path, composition_path, prompts_path):
        if not path.is_file():
            raise PolicyError(f"缺少策略文件：{path}")
    strategy = _read_yaml(strategy_path)
    composition = _read_yaml(composition_path)
    prompts = _read_yaml(prompts_path)
    policy = Policy(strategy=strategy, composition=composition,
                    sha256=_digest(strategy_path, composition_path, prompts_path),
                    paths=(str(strategy_path), str(composition_path), str(prompts_path)),
                    prompts=prompts)
    if validate:
        validate_semantics(policy)
    return policy


def _walk_phrases(node, prefix=""):
    """遍历短语表，产出 (路径, 文本)。只用于自查风格词。"""
    if isinstance(node, str):
        yield prefix, node
    elif isinstance(node, list):
        for i, item in enumerate(node):
            yield from _walk_phrases(item, f"{prefix}[{i}]")
    elif isinstance(node, dict):
        for key, value in node.items():
            yield from _walk_phrases(value, f"{prefix}.{key}" if prefix else key)


def validate_semantics(policy: Policy) -> None:
    """引用完整性：Schema 管不到、但运行时会崩或静默失真的跨字段关系。"""
    _validate(policy.strategy, "policy.schema.json", "strategy.yaml")
    _validate(policy.composition, "composition.schema.json", "composition.yaml")
    _validate(policy.prompts or {}, "prompts.schema.json", "prompts.yaml")

    problems: list[str] = []

    # 0. 提示词层：枚举与短语表双向对齐；顺序覆盖；策略自身不许注入风格词
    prompts = policy.prompts or {}
    phrases = prompts.get("phrases") or {}
    for enum_name, slot in (("media_function", "function"), ("media_role", "role")):
        values = set(policy.enum(enum_name))
        keys = set(phrases.get(slot) or {})
        if values != keys:
            problems.append(f"prompts.phrases.{slot} 与 enums.{enum_name} 不一致："
                            f"缺 {sorted(values - keys)}，多 {sorted(keys - values)}")
    missing_slots = {"role", "subject", "function", "framing", "ratio"} - set(prompts.get("clause_order") or [])
    if missing_slots:
        problems.append(f"prompts.clause_order 缺少必需段：{sorted(missing_slots)}")
    banned = [w.lower() for w in (prompts.get("style_words_banned") or [])]
    for slot, value in _walk_phrases(phrases):
        hit = [w for w in banned if w in value.lower()]
        if hit:
            problems.append(f"prompts.phrases.{slot} 含风格词 {hit}：风格由作者写在 asset.subject，"
                            "策略层注入风格会让整批稿子趋同")

    # 1. 字号阶梯 / 下限 / 行高 / 字距 / 行数必须同键，避免「有字号没下限」
    scale = policy.req("typography.scale")
    for section in ("floors", "line_height", "tracking_pt", "max_lines"):
        keys = set(policy.req(f"typography.{section}"))
        missing = sorted(set(scale) - keys)
        extra = sorted(keys - set(scale))
        if missing or extra:
            problems.append(f"typography.{section} 与 scale 键不一致：缺 {missing or '无'}，多 {extra or '无'}")
    for key, floor in (policy.req("typography.floors") or {}).items():
        if key in scale and floor > scale[key]:
            problems.append(f"typography.floors.{key}={floor} 大于 scale.{key}={scale[key]}")

    # 2. 颜色 token 与 defaults 必须一一对应（多写 token 不写值 = 渲染时静默回落）
    tokens = set(policy.req("colors.tokens"))
    defaults = set(policy.colors())
    if tokens != defaults:
        problems.append(f"colors.tokens 与 colors.defaults 不一致：仅 token {sorted(tokens - defaults)}，"
                        f"仅默认值 {sorted(defaults - tokens)}")
    for token, value in policy.colors().items():
        if not isinstance(value, str) or not (len(value) == 7 and value.startswith("#")):
            problems.append(f"colors.defaults.{token} 必须是 #RRGGBB，实际 {value!r}")

    # 3. 枚举与配方双向对齐
    compositions = set(policy.enum("composition"))
    recipes = set(policy.composition.get("recipes") or {})
    if compositions != recipes:
        problems.append(f"enums.composition 与 composition.yaml 配方不一致：缺 {sorted(compositions - recipes)}，"
                        f"多 {sorted(recipes - compositions)}")

    # 3.5 配方自带设计属性：字号档、第一落点、禁令清单（这三样过去写在 judgment.py 里）
    scale_keys = set(policy.req("typography.scale"))
    focus_values = set(policy.enum("focus"))
    for name, recipe in (policy.composition.get("recipes") or {}).items():
        if str(recipe.get("ladder")) not in scale_keys:
            problems.append(f"composition.{name}.ladder={recipe.get('ladder')!r} 不在 typography.scale 里")
        if str(recipe.get("focus")) not in focus_values:
            problems.append(f"composition.{name}.focus={recipe.get('focus')!r} 不在 enums.focus 里")
        if not (recipe.get("must_not") or []):
            problems.append(f"composition.{name}.must_not 为空：配方必须写明「这个形状下不该做什么」")

    # 3.6 槽位高度必须装得下它自己声明的行数（几何与排版规则的一致性）
    from typography import ink_height
    ch = float(policy.req("canvas.height")) - float(policy.req("canvas.margin.top")) \
        - float(policy.req("canvas.margin.bottom"))
    for name, recipe in (policy.composition.get("recipes") or {}).items():
        for i, slot in enumerate(recipe.get("slots") or []):
            lines = slot.get("max_lines")
            key = slot.get("size") or recipe.get("ladder")
            if not lines or slot.get("type") != "text" or key not in scale:
                continue
            needed = ink_height(int(lines), float(scale[key]), float(policy.line_height(key)), policy)
            have = float(slot["zone"][3]) * ch
            if have + 1.0 < needed:
                problems.append(
                    f"composition.{name}.slots[{i}:{slot.get('role')}] 高度 {have:.0f}px 装不下 "
                    f"{lines} 行 {key}（需要 {needed:.0f}px）：要么加高槽位，要么减 max_lines")

    # 4. 槽位引用：字号档 / 颜色 token / 元素类型 必须存在
    for name, recipe in (policy.composition.get("recipes") or {}).items():
        for i, slot in enumerate(recipe.get("slots") or []):
            where = f"composition.{name}.slots[{i}:{slot.get('role')}]"
            if slot["type"] not in set(policy.enum("slot_type")):
                problems.append(f"{where} 的 type={slot['type']!r} 不在 enums.slot_type")
            if slot.get("size") and slot["size"] not in scale:
                problems.append(f"{where} 的 size={slot['size']!r} 不在 typography.scale")
            if slot.get("color") and slot["color"] not in tokens:
                problems.append(f"{where} 的 color={slot['color']!r} 不在 colors.tokens")
            zone = slot["zone"]
            if zone[0] + zone[2] > 1.0001 or zone[1] + zone[3] > 1.0001:
                problems.append(f"{where} 的 zone 越出内容区：{zone}")

    # 5. 词表与探针：有词表就必须给探针，否则命中后无处可问（只提示不追问 = 噪声）
    lexicons = set(policy.req("judgment.lexicons"))
    probes = set(policy.req("judgment.probes"))
    for required in ("comparison_end", "part_whole", "sequence", "structure",
                     "claim_cue", "assertion_cue", "change_cue", "imperative_cue"):
        if required not in lexicons:
            problems.append(f"judgment.lexicons 缺 {required}：证据形状判断会退化为纯默认值")
    unknown = sorted(probes - lexicons)
    if unknown:
        problems.append(f"judgment.probes 指向不存在的词表：{unknown}")

    # 6. 素材策略：枚举与策略必须自洽，否则运行期会静默退回默认行为
    assets = policy.req("assets")
    fits = {"cover", "contain"}
    bad = sorted(set(assets["crop_allow"]) - fits)
    if bad:
        problems.append(f"assets.crop_allow 含未知 fit：{bad}")
    if assets["oversample"] <= 0:
        problems.append("assets.oversample 必须 > 0")
    thr = policy.req("colors.contrast")
    if assets["text_over_image"]["min_contrast_ratio"] < thr["large_min"]:
        problems.append("assets.text_over_image.min_contrast_ratio 不得低于 colors.contrast.large_min："
                        "压图文字的对比要求不能比正文更松")
    if assets["qc"]["flat_std"] >= assets["qc"]["busy_std"]:
        problems.append("assets.qc.flat_std 必须小于 busy_std")
    for required in ("visual_subject",):
        if required not in lexicons:
            problems.append(f"judgment.lexicons 缺 {required}：素材判断会失去唯一的提示通道")

    # 7. 验收模式与出口码
    modes = set(policy.req("verification.modes"))
    if policy.req("verification.default_mode") not in modes:
        problems.append("verification.default_mode 不在 verification.modes 里")
    if not set(policy.req("verification.release_requires")) <= {"source", "unit", "period", "basis"}:
        problems.append("verification.release_requires 只接受 source/unit/period/basis")

    if problems:
        raise PolicyError("策略语义校验失败：\n" + "\n".join(f"  - {p}" for p in problems))


def to_json(policy: Policy) -> dict:
    """报告里审计用的策略摘要（不含正文，只含指纹与关键阈值）。"""
    return {
        "policy_sha256": policy.sha256,
        "policy_id": policy.strategy.get("id"),
        "composition_id": policy.composition.get("id"),
        "files": list(policy.paths),
        "canvas": policy.req("canvas"),
        "key_thresholds": {
            k: policy.strategy["thresholds"][k]
            for k in ("ink_overlap_px", "position_drift_px", "fidelity_pixel_mismatch_max",
                      "numeric_density_high", "min_gap_px", "chart_scale_skew_ratio")
        },
    }
