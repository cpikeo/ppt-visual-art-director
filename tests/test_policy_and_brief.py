#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""策略与输入契约：形状、引用完整性、拒绝而非静默降级。"""
import copy
import unittest

import support
from support import ROOT, tmpdir, write, write_brief

from policy import PolicyError, load as load_policy
from brief import BriefError, load as load_brief

SCHEMA_DIR = ROOT / "schema"


class TestStrategyPolicy(unittest.TestCase):
    def setUp(self):
        self.policy = load_policy()

    def test_schema_and_semantics_pass(self):
        self.assertTrue(self.policy.sha256)
        self.assertEqual(len(self.policy.sha256), 64)
        self.assertIn("metric_row", self.policy.enum("composition"))

    def test_missing_key_has_no_hidden_default(self):
        with self.assertRaises(PolicyError):
            self.policy.req("typography.scale.nonexistent")

    def test_enum_rejects_unknown_value(self):
        self.assertFalse(self.policy.is_enum("evidence", "vibes"))
        self.assertTrue(self.policy.is_enum("evidence", "series"))

    def test_scale_ladder_is_monotone(self):
        scale = self.policy.req("typography.scale")
        order = ["display", "h1", "h2", "lede", "body", "label", "micro", "source"]
        values = [scale[k] for k in order]
        self.assertEqual(values, sorted(values, reverse=True),
                         "字号阶梯必须单调递减，否则层级无法用语义档表达")

    def test_floor_never_above_scale(self):
        for key, floor in self.policy.req("typography.floors").items():
            self.assertLessEqual(floor, self.policy.req(f"typography.scale.{key}"))

    def test_lexicon_without_probe_is_rejected(self):
        tmp = tmpdir()
        policy_dir = support.copy_policy(tmp)
        text = (policy_dir / "strategy.yaml").read_text(encoding="utf-8")
        text = text.replace("  probes:\n", "  probes_removed:\n")  # 破坏引用
        (policy_dir / "strategy.yaml").write_text(text, encoding="utf-8")
        with self.assertRaises(PolicyError):
            load_policy(policy_dir)

    def test_composition_recipe_must_exist_for_every_enum(self):
        tmp = tmpdir()
        policy_dir = support.copy_policy(tmp)
        path = policy_dir / "composition.yaml"
        text = path.read_text(encoding="utf-8").replace("  statement:\n", "  statement_typo:\n", 1)
        path.write_text(text, encoding="utf-8")
        with self.assertRaises(PolicyError) as ctx:
            load_policy(policy_dir)
        self.assertIn("enums.composition", str(ctx.exception))

    def test_slot_size_must_come_from_scale(self):
        tmp = tmpdir()
        policy_dir = support.copy_policy(tmp)
        path = policy_dir / "composition.yaml"
        text = path.read_text(encoding="utf-8").replace("size: h1", "size: huge", 1)
        path.write_text(text, encoding="utf-8")
        with self.assertRaises(PolicyError):
            load_policy(policy_dir)

    def test_color_tokens_must_match_defaults(self):
        tmp = tmpdir()
        policy_dir = support.copy_policy(tmp)
        path = policy_dir / "strategy.yaml"
        text = path.read_text(encoding="utf-8").replace('    hairline: "#D8D3CB"\n', "")
        path.write_text(text, encoding="utf-8")
        with self.assertRaises(PolicyError):
            load_policy(policy_dir)


class TestBriefContract(unittest.TestCase):
    def setUp(self):
        self.policy = load_policy()

    def test_minimal_brief_loads_with_explicit_notes(self):
        brief, digest = load_brief(write_brief(tmpdir()), self.policy, SCHEMA_DIR)
        self.assertEqual(len(brief["slides"]), 4)
        self.assertEqual(len(digest), 64)
        # 每个补默认的地方都要留下记录（这里靠 brief 自己写全，所以只允许 unknown 类提示）
        self.assertFalse([n for n in brief["_brief_notes"] if n["kind"] in ("empty", "absent")])
        self.assertTrue(brief["slides"][0]["id"].startswith("s"))

    def test_missing_audience_is_noted_not_invented(self):
        tmp = tmpdir()
        text = support.MINI_BRIEF.replace("audience: 内部评审会\n", "")
        brief, _ = load_brief(write_brief(tmp, text=text), self.policy, SCHEMA_DIR)
        self.assertEqual(brief["audience"], "")
        self.assertTrue(any(n["field"] == "audience" and n["kind"] == "empty" for n in brief["_brief_notes"]))

    def test_unknown_field_is_rejected(self):
        tmp = tmpdir()
        text = support.MINI_BRIEF.replace("tension:", "tone_of_voice:")
        with self.assertRaises(BriefError) as ctx:
            load_brief(write_brief(tmp, text=text), self.policy, SCHEMA_DIR)
        self.assertIn("Additional properties", str(ctx.exception))

    def test_illegal_enum_is_rejected(self):
        tmp = tmpdir()
        text = support.MINI_BRIEF.replace("    title: 加店不再带来增长",
                                          "    title: 加店不再带来增长\n    composition: fancy")
        with self.assertRaises(BriefError):
            load_brief(write_brief(tmp, text=text), self.policy, SCHEMA_DIR)

    def test_duplicate_slide_id_is_rejected(self):
        tmp = tmpdir()
        text = support.MINI_BRIEF.replace("  - id: s02", "  - id: s01")
        with self.assertRaises(BriefError) as ctx:
            load_brief(write_brief(tmp, text=text), self.policy, SCHEMA_DIR)
        self.assertIn("id 重复", str(ctx.exception))

    def test_yaml_bomb_size_guard(self):
        tmp = tmpdir()
        big = tmp / "big.yaml"
        big.write_text("slides:\n  - title: " + "x" * 2_100_000 + "\n", encoding="utf-8")
        with self.assertRaises(BriefError) as ctx:
            load_brief(big, self.policy, SCHEMA_DIR)
        self.assertIn("过大", str(ctx.exception))

    def test_evidence_page_without_content_is_marked_absent(self):
        tmp = tmpdir()
        text = ("audience: 评委会\ndecision: 批准\nslides:\n"
                "  - id: s01\n    title: 只有标题没有证据\n")
        brief, _ = load_brief(write_brief(tmp, text=text), self.policy, SCHEMA_DIR)
        self.assertTrue(any(n["kind"] == "absent" for n in brief["_brief_notes"]))
        self.assertEqual(brief["slides"][0]["content"], "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
