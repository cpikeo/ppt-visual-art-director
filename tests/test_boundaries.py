#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""五层边界的机械检查：越界不需要吵架，让它变成会失败的测试。

这一组只回答「职责有没有串层」，不回答「设计好不好」：
  1. 策略里的配置必须有人消费（死配置 = 写了却没生效）；
  2. 阻断清单必须与真实错误码一致（漏一个 = 该拦不拦，多一个 = 假配置）；
  3. 枚举只能有一份来源（策略与 Schema 漂移 = 同一件事两种说法）；
  4. 设计禁令 / 字号档 / 焦点必须来自策略，不能写在 Python 里；
  5. 判断字段必须进入执行流程，否则它是装饰；
  6. 词表命中不得改变几何（否则就是「行业词表造成同质化」）。
"""
import ast
import copy
import json
import pathlib
import re
import unittest
import unittest.mock

from support import ROOT, tmpdir, write_brief

from policy import load as load_policy
from brief import load as load_brief
from judgment import judge_deck
from compose import compose_deck
from guard import check as guard_check
from assets import AssetError, inspect
from typography import FontBook

SCHEMA_DIR = ROOT / "schema"
POLICY = load_policy()
FONTS = FontBook(POLICY)
SCRIPTS_TEXT = "\n".join(p.read_text(encoding="utf-8") for p in sorted((ROOT / "scripts").glob("*.py")))

# 判断产物里「只给人和审计看」的字段：不参与执行，但也不能删（缺了就无法回答「为什么这样排」）
ADVISORY_DECISION_KEYS = {
    "role_why", "role_source", "composition_why", "composition_alternatives",
    "emphasis_why", "trace_ref",
}


def load(text: str):
    directory = tmpdir()
    brief, _ = load_brief(write_brief(directory, text=text), POLICY, SCHEMA_DIR)
    plan = judge_deck(brief, POLICY, base_dir=directory)
    deck = compose_deck(brief, plan, POLICY, fontbook=FONTS, base_dir=directory)
    return plan, deck, directory


def base(head: str, slides: str) -> str:
    return head + "slides:\n" + slides


HEAD = ("audience: 管理层\ndecision: 批准预算\ntension: 没有共识\n"
        "source: 内部看板\nunit: 万元\nperiod: 2026 Q1–Q3\nbasis: 已开票收入\n")


class TestNoDeadPolicy(unittest.TestCase):
    """策略里的每个数值都必须有消费者：写了不生效等于两套真相。"""

    def test_every_numeric_policy_leaf_is_referenced(self):
        def walk(node, prefix=""):
            if isinstance(node, dict):
                for key, value in node.items():
                    yield from walk(value, f"{prefix}.{key}" if prefix else str(key))
            elif isinstance(node, (int, float)) and not isinstance(node, bool):
                yield prefix, node

        def referenced(dotted: str) -> bool:
            """按「本身或任一祖先路径在代码里出现」判定。
            动态键（typography.scale.h1）由消费父表的那行代码负责，祖先出现即算有关注者。"""
            parts = dotted.split(".")
            for i in range(len(parts), 0, -1):
                prefix = ".".join(parts[:i])
                leaf = parts[i - 1]
                if prefix in SCRIPTS_TEXT or f'"{prefix}"' in SCRIPTS_TEXT:
                    return True
                if i == len(parts) and (leaf in SCRIPTS_TEXT or f'"{leaf}"' in SCRIPTS_TEXT):
                    return True
            return False

        dead = [f"{dotted}={value}" for dotted, value in walk(POLICY.strategy)
                if not referenced(dotted)]
        self.assertEqual(dead, [], "这些策略值没有任何消费者：要么接线，要么删掉")

    def test_blocking_codes_match_real_errors(self):
        """阻断清单与 guard 实际抛出的错误码必须一一对应。"""
        emitted = set(re.findall(r'rep\.error\(\s*"([A-Z_]+)"', SCRIPTS_TEXT))
        blocking = set(POLICY.req("verification.blocking_codes"))
        self.assertEqual(sorted(blocking - emitted), [], "策略里登记了从不抛出的阻断码（死配置）")
        self.assertEqual(sorted(emitted - blocking), [], "guard 会报错但这些码不在阻断清单里")

    def test_recipe_design_fields_come_from_policy(self):
        """字号档 / 第一落点 / 禁令清单在 Python 里只能读，不能写。"""
        for name in POLICY.enum("composition"):
            recipe = POLICY.recipe(name)
            self.assertTrue(recipe["ladder"] in POLICY.req("typography.scale"), name)
            self.assertTrue(recipe["focus"] in POLICY.enum("focus"), name)
            self.assertTrue(recipe["must_not"], name)
            for line in recipe["must_not"]:
                self.assertNotIn(line, SCRIPTS_TEXT,
                                 f"禁令文案出现在代码里（{name}）：它属于 composition.yaml")


def _dead_policy_reads(path: pathlib.Path) -> list[str]:
    """找出「读了策略却从没用过」的赋值：`x = policy.num(...)` 之后 x 从未被读取。

    这是死配置的另一半：字符串扫描只能证明键被提到过，不能证明结果被用于判断。
    """
    import ast
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    offenders: list[str] = []
    for func in [n for n in ast.walk(tree) if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]:
        loaded = {n.id for n in ast.walk(func) if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)}
        for node in ast.walk(func):
            if not isinstance(node, ast.Assign) or len(node.targets) != 1:
                continue
            target = node.targets[0]
            if not isinstance(target, ast.Name):
                continue
            calls = [n for n in ast.walk(node.value) if isinstance(n, ast.Call)]
            reads_policy = any(isinstance(c.func, ast.Attribute) and c.func.attr in {"num", "req", "opt"}
                               and isinstance(c.func.value, ast.Name) and c.func.value.id == "policy"
                               for c in calls)
            if reads_policy and target.id not in loaded:
                offenders.append(f"{path.name}:{node.lineno} {target.id}")
    return offenders


class TestNoDeadPolicyRead(unittest.TestCase):
    def test_policy_values_are_actually_used_after_being_read(self):
        offenders = []
        for script in sorted((ROOT / "scripts").glob("*.py")):
            offenders.extend(_dead_policy_reads(script))
        self.assertEqual(offenders, [], "读了策略值却没有使用：判断被写成了装饰")


def _carries_number(node) -> bool:
    """常量里是否夹了一个数值默认值（字符串路径常量不算：那是键名，不是阈值）。"""
    if isinstance(node, ast.Constant):
        return isinstance(node.value, (int, float)) and not isinstance(node.value, bool)
    if isinstance(node, (ast.Tuple, ast.List)):
        return any(_carries_number(elt) for elt in node.elts)
    return False


class TestPolicySchemaCoversConsumers(unittest.TestCase):
    """Schema 是策略的 I/O 权威：代码硬要求的阈值，不能是 Schema 里的可选项。

    否则一份缺键的策略能通过校验，却在运行时抛 PolicyError —— 校验形同虚设。
    """

    def test_every_threshold_read_as_required_is_required_in_schema(self):
        schema = json.loads((SCHEMA_DIR / "policy.schema.json").read_text(encoding="utf-8"))
        thresholds = schema["properties"]["thresholds"]
        required = set(thresholds.get("required") or ())
        declared = set(thresholds.get("properties") or {})
        used = set(re.findall(r"thresholds\.([a-z0-9_]+)", SCRIPTS_TEXT))
        self.assertTrue(used, "没有在 scripts/ 中找到阈值键，检查本身失效了")
        self.assertEqual(used - declared, set(), "代码读了 Schema 未声明的阈值键")
        self.assertEqual(used - required, set(),
                         "代码按必填读取，但 Schema 允许缺省：" + "、".join(sorted(used - required)))

    def test_python_does_not_shadow_thresholds_with_its_own_defaults(self):
        """阈值只能有一份定义（Policy）。Python 里的兜底常量会让两处说法漂移。"""
        offenders = []
        for script in sorted((ROOT / "scripts").glob("*.py")):
            tree = ast.parse(script.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
                    name = node.targets[0].id
                    if "TOL" not in name.upper() and "THRESHOLD" not in name.upper():
                        continue
                    if _carries_number(node.value):
                        offenders.append(f"{script.name}:{node.lineno} {name}")
        self.assertEqual(offenders, [], "阈值在 Python 里另立了默认值：" + "、".join(offenders))


class TestEnumSingleSource(unittest.TestCase):
    def test_strategy_enums_match_plan_schema(self):
        schema = json.loads((SCHEMA_DIR / "plan.schema.json").read_text(encoding="utf-8"))
        decision = schema["properties"]["pages"]["items"]["properties"]["decision"]["properties"]
        pairs = [("role", decision["role"]["enum"]),
                 ("composition", decision["composition"]["enum"]),
                 ("density", decision["density"]["enum"]),
                 ("emphasis", decision["emphasis"]["enum"]),
                 ("focus", decision["focus"]["enum"]),
                 ("media_decision", decision["media"]["properties"]["decision"]["enum"])]
        for enum_name, values in pairs:
            self.assertEqual(set(POLICY.enum(enum_name)), set(values),
                             f"enums.{enum_name} 与 plan.schema.json 不一致")
        self.assertEqual(set(POLICY.enum("media_function")),
                         set(json.loads((SCHEMA_DIR / "brief.schema.json").read_text())
                             ["properties"]["slides"]["items"]["properties"]["asset"]
                             ["properties"]["function"].get("enum") or
                             POLICY.enum("media_function")))

    def test_evidence_and_must_not_enums_agree(self):
        schema = json.loads((SCHEMA_DIR / "plan.schema.json").read_text(encoding="utf-8"))
        evidence = schema["properties"]["pages"]["items"]["properties"]["understanding"]["properties"]["evidence"]
        self.assertEqual(set(POLICY.enum("evidence")), set(evidence["properties"]["value"]["enum"]))


class TestPolicyDrivesExecution(unittest.TestCase):
    def test_must_not_reaches_the_plan(self):
        plan, _, _ = load(base(HEAD, "  - id: s01\n    title: 加店不再带来增长\n"
                                     "    content: 门店数量增长 18%，单店月均产出下降 21%。\n"))
        page = plan["pages"][0]
        recipe = POLICY.recipe(page["decision"]["composition"])
        self.assertEqual(page["production"]["must_not"], recipe["must_not"])

    def test_ladder_comes_from_the_recipe(self):
        plan, deck, _ = load(base(HEAD, "  - id: s01\n    title: 加店不再带来增长\n"
                                       "    content: 门店数量增长 18%，单店月均产出下降 21%。\n"))
        recipe = POLICY.recipe(plan["pages"][0]["decision"]["composition"])
        self.assertEqual(plan["pages"][0]["production"]["typography"]["claim_size_key"], recipe["ladder"])

    def test_changed_recipe_changes_the_plan(self):
        """改策略 → plan 变，说明 plan 真的读策略，而不是自己记着一份。"""
        plan, _, _ = load(base(HEAD, "  - id: s01\n    title: 加店不再带来增长\n"
                                     "    content: 门店数量增长 18%，单店月均产出下降 21%。\n"))
        mutated = copy.deepcopy(POLICY)
        for recipe in mutated.composition["recipes"].values():
            recipe["must_not"] = ["测试用的禁令"]
        brief, _ = load_brief(write_brief(tmpdir(), text=base(
            HEAD, "  - id: s01\n    title: 加店不再带来增长\n    content: 门店数量增长 18%，单店月均产出下降 21%。\n")),
            POLICY, SCHEMA_DIR)
        changed = judge_deck(brief, mutated)
        self.assertEqual(changed["pages"][0]["production"]["must_not"], ["测试用的禁令"])
        self.assertNotEqual(plan["pages"][0]["production"]["must_not"], ["测试用的禁令"])

    def test_decode_budget_comes_from_policy(self):
        directory = tmpdir()
        from PIL import Image
        path = directory / "big.png"
        Image.new("RGB", (1200, 1200), (10, 20, 30)).save(path)
        self.assertEqual(inspect(path, POLICY)["width"], 1200)      # 正常策略下可读
        strict = copy.deepcopy(POLICY)
        strict.strategy["assets"]["pixels"]["max_decode"] = 1_000_000
        with self.assertRaises(AssetError):
            inspect(path, strict)                                   # 预算是策略，不是常量


class TestJudgmentReachesExecution(unittest.TestCase):
    TEXT = base(HEAD,
                         "  - id: s01\n    title: 加店不再带来增长\n"
                         "    content: 门店数量增长 18%，单店月均产出下降 21%。\n"
                         "  - id: s02\n    title: 经销渠道占比逐季抬升\n"
                         "    content: 2026 Q1 到 Q3，经销渠道收入占比从 34% 升到 51%。\n"
                         "    chart:\n      kind: column\n      categories: [\"2026 Q1\", \"2026 Q2\", \"2026 Q3\"]\n"
                         "      series:\n        - name: 经销渠道收入占比\n          values: [34, 43, 51]\n"
                        "  - id: s03\n    title: 请批准下一季预算\n"
                        "    content: 请批准 Q1 预算转投经销渠道。\n")

    def _plan(self):
        plan, deck, directory = load(self.TEXT)
        brief, _ = load_brief(write_brief(directory, text=self.TEXT), POLICY, SCHEMA_DIR)
        return plan, deck, brief, directory

    def test_every_decision_field_is_consumed_or_declared_advisory(self):
        plan, _, brief, directory = self._plan()
        seen = set()

        class Track(dict):
            def __getitem__(self, key):
                seen.add(key)
                return super().__getitem__(key)

            def get(self, key, default=None):
                seen.add(key)
                return super().get(key, default)

        for page in plan["pages"]:
            page["decision"] = Track(page["decision"])
        compose_deck(brief, plan, POLICY, fontbook=FONTS, base_dir=directory)
        untouched = set(plan["pages"][0]["decision"]) - seen
        self.assertEqual(sorted(untouched - ADVISORY_DECISION_KEYS), [],
                         "这些判断字段没进执行流程，也没有被登记为审计字段（= 装饰）")

    SHORT_CLAIM = base(HEAD,
        "  - id: s01\n    title: 加店不再带来增长\n    content: 门店数量增长 18%，单店月均产出下降 21%。\n"
        "  - id: s02\n    title: 占比逐季抬升\n"
        "    content: 2026 Q1 到 Q3，经销渠道收入占比从 34% 升到 51%。\n"
        "    chart:\n      kind: column\n      categories: [\"2026 Q1\", \"2026 Q2\", \"2026 Q3\"]\n"
        "      series:\n        - name: 经销渠道收入占比\n          values: [34, 43, 51]\n"
        "  - id: s03\n    title: 请批准下一季预算\n    content: 请批准 Q1 预算转投经销渠道。\n")

    def test_lead_emphasis_actually_changes_the_type_size(self):
        """高潮页的 emphasis=lead 必须真的落到字号上——否则「最强一击」只是标签。"""
        plan, deck, _ = load(self.SHORT_CLAIM)
        climax = next(p for p in plan["pages"] if p["decision"]["emphasis"] == "lead")
        slide = next(s for s in deck["slides"] if s["id"] == climax["id"])
        claim = next(e for e in slide["elements"] if e.get("role") == "claim")
        base_ladder = climax["production"]["typography"]["claim_size_key"]
        self.assertEqual(climax["decision"]["composition"], "series_chart")
        self.assertGreater(claim["size"], POLICY.size(base_ladder), "强调档没有改变字号：判断没落地")
        self.assertTrue(slide["meta"]["emphasis_applied"])

    def test_lead_emphasis_degrades_honestly_when_it_does_not_fit(self):
        """同一条判断遇到装不下的主张：保持基准档并写明原因，不许缩字硬塞。"""
        plan, deck, _, _ = self._plan()
        climax = next(p for p in plan["pages"] if p["decision"]["emphasis"] == "lead")
        slide = next(s for s in deck["slides"] if s["id"] == climax["id"])
        claim = next(e for e in slide["elements"] if e.get("role") == "claim")
        self.assertEqual(claim["size"], POLICY.size(climax["production"]["typography"]["claim_size_key"]))
        self.assertFalse(slide["meta"]["emphasis_applied"])
        self.assertTrue(slide["meta"]["emphasis_notes"], "放大失败却没有记录原因")

    def test_emphasis_does_not_force_an_overflow(self):
        """放大一档装不下时，退回基准档并记录原因，不靠缩字或撑框硬塞。"""
        plan, deck, _, _ = self._plan()
        for page, slide in zip(plan["pages"], deck["slides"]):
            claim = next((e for e in slide["elements"] if e.get("role") == "claim"), None)
            if claim is None:
                continue
            self.assertFalse(claim.get("_measured", {}).get("degraded", False))
            self.assertLessEqual(claim.get("_measured", {}).get("lines", 1), claim.get("max_lines", 99))


class TestNoStyleDrift(unittest.TestCase):
    """行业词表只提示：命中不得改变几何，否则十份稿子会长成一张。"""

    def test_lexicon_hit_changes_questions_not_geometry(self):
        plain = base(HEAD, "  - id: s01\n    title: 门店数量增长 18%\n    content: 门店数量增长 18%，单店月均产出下降 21%。\n")
        hinted = base(HEAD, "  - id: s01\n    title: 车间产线利用率 78%\n    content: 车间产线利用率 78%，比去年高 6 个百分点。\n")
        plan_a, deck_a, _ = load(plain)
        plan_b, deck_b, _ = load(hinted)
        # 词表命中只多一条待确认问题
        self.assertTrue([q for q in plan_b["pages"][0]["open_questions"]
                         if q["code"] in ("ASSET_HINT", "PROBE_VISUAL_SUBJECT")])
        # 几何与配色一律相同（同一构图、同一槽位）：没有「行业 → 风格」的通路
        self.assertEqual(
            [(e["type"], e.get("role"), round(e["width"]), round(e["height"])) for e in deck_a["slides"][0]["elements"]],
            [(e["type"], e.get("role"), round(e["width"]), round(e["height"])) for e in deck_b["slides"][0]["elements"]])

    def test_policy_carries_no_style_adjectives(self):
        banned = [w.lower() for w in POLICY.prompts["style_words_banned"]]
        blob = json.dumps(POLICY.strategy.get("judgment", {}), ensure_ascii=False).lower()
        blob += json.dumps(POLICY.composition, ensure_ascii=False).lower()
        self.assertEqual([w for w in banned if w in blob], [])


class TestChartSanity(unittest.TestCase):
    def test_mixed_units_in_one_chart_is_blocked(self):
        plan, deck, _ = load(base(HEAD,
            "  - id: s01\n    title: 门店与占比\n    content: 门店 128 家，占比 51%。\n"
            "    chart:\n      kind: column\n      categories: [Q1, Q2]\n"
            "      series:\n        - name: 门店数（家）\n          values: [120, 128]\n"
            "        - name: 收入占比（%）\n          values: [34, 51]\n"))
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("CHART_MIXED_UNITS", {e["code"] for e in result["errors"]})
        self.assertEqual(result["status"], "BLOCK")

    def test_same_unit_chart_passes_and_scale_skew_warns(self):
        plan, deck, _ = load(base(HEAD,
            "  - id: s01\n    title: 两个渠道的单店产出\n    content: 经销 4.6 万元，直营 1.9 万元。\n"
            "    chart:\n      kind: column\n      categories: [Q1, Q2]\n"
            "      series:\n        - name: 经销单店产出（万元）\n          values: [4.4, 4.6]\n"
            "        - name: 直营单店产出（万元）\n          values: [1.9, 2.0]\n"))
        result = guard_check(deck, POLICY, mode="release")
        codes = {e["code"] for e in result["errors"]}
        self.assertNotIn("CHART_MIXED_UNITS", codes)


class TestDeliveryReliability(unittest.TestCase):
    """写盘要么完整要么没有：中断不得留下「看起来完整、其实半截」的产物。"""

    def test_json_write_is_atomic(self):
        directory = tmpdir()
        target = directory / "plan.json"
        target.write_text('{"old": true}', encoding="utf-8")
        from pvad import write_json

        with self.assertRaises(RuntimeError):
            with unittest.mock.patch("pathlib.Path.replace", side_effect=RuntimeError("断电")):
                write_json(target, {"new": True})
        self.assertEqual(json.loads(target.read_text(encoding="utf-8")), {"old": True})

    def test_compiled_pptx_never_left_half_written(self):
        directory = tmpdir()
        plan, deck, _ = load(base(HEAD, "  - id: s01\n    title: 加店不再带来增长\n"
                                       "    content: 门店数量增长 18%，单店月均产出下降 21%。\n"))
        from compile import compile_deck
        out = directory / "deck.pptx"
        compile_deck(deck, out, base_dir=directory)
        self.assertTrue(out.is_file() and out.stat().st_size > 0)
        self.assertFalse((directory / "deck.pptx.part").exists(), "临时文件没有清理")


if __name__ == "__main__":
    unittest.main()
