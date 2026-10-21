#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""硬门回归：每一种 BLOCK 都要能被一个最小改动触发，且修复提示可执行。

这些用例保证的是「错误不会被静默放过」，不是「稿子好看」。
"""
import copy
import unittest

import support
from support import ROOT, tmpdir

from policy import load as load_policy
from brief import load as load_brief
from judgment import judge_deck
from compose import compose_deck
from guard import check as guard_check

SCHEMA_DIR = ROOT / "schema"
POLICY = load_policy()


def base_deck():
    tmp = tmpdir()
    brief, _ = load_brief(support.write_brief(tmp), POLICY, SCHEMA_DIR)
    return compose_deck(brief, judge_deck(brief, POLICY), POLICY)


def codes(result):
    return {e["code"] for e in result["errors"]}


def warnings(result):
    return {w["code"] for w in result["warnings"]}


class TestBaseline(unittest.TestCase):
    def test_baseline_passes_in_draft(self):
        result = guard_check(base_deck(), POLICY, mode="draft")
        self.assertEqual(result["status"], "PASS", result["errors"])
        self.assertTrue(result["stats"]["measured"])

    def test_guard_is_deterministic(self):
        deck = base_deck()
        first = guard_check(deck, POLICY, mode="draft")
        second = guard_check(copy.deepcopy(deck), POLICY, mode="draft")
        self.assertEqual(first, second)


class TestGeometryGuards(unittest.TestCase):
    def test_text_overflow_is_blocked(self):
        deck = base_deck()
        element = deck["slides"][0]["elements"][1]
        element["height"] = 20
        element["max_lines"] = 1
        element["text"] = "这是一段明显装不进一个小框的长文本，用来验证溢出会被拦住。"
        result = guard_check(deck, POLICY, mode="draft")
        self.assertEqual(result["status"], "BLOCK")
        self.assertTrue({"TEXT_OVERFLOW", "TEXT_TOO_WIDE"} & codes(result))

    def test_out_of_canvas_is_blocked(self):
        deck = base_deck()
        deck["slides"][0]["elements"][1]["x"] = 1200
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("OUT_OF_CANVAS", codes(result))

    def test_degenerate_geometry_is_blocked(self):
        deck = base_deck()
        deck["slides"][0]["elements"][1]["width"] = 0
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("GEOMETRY_INVALID", codes(result))

    def test_overlap_is_blocked(self):
        deck = base_deck()
        claim = deck["slides"][0]["elements"][1]
        victim = deck["slides"][0]["elements"][2]
        victim["x"], victim["y"] = claim["x"] + 10, claim["y"] + 10
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("ELEMENT_OVERLAP", codes(result))

    def test_size_below_floor_is_blocked(self):
        deck = base_deck()
        element = deck["slides"][0]["elements"][2]
        element["size"] = 8
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("TEXT_SIZE_FLOOR", codes(result))

    def test_duplicate_element_id_is_blocked(self):
        deck = base_deck()
        deck["slides"][0]["elements"][2]["id"] = deck["slides"][0]["elements"][1]["id"]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("ELEMENT_ID_DUPLICATE", codes(result))

    def test_empty_slide_is_blocked(self):
        deck = base_deck()
        deck["slides"][0]["elements"] = []
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("EMPTY_SLIDE", codes(result))


class TestContractGuards(unittest.TestCase):
    def test_misspelled_field_is_blocked(self):
        deck = base_deck()
        element = deck["slides"][0]["elements"][2]
        element["font_size"] = element.pop("size")
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("FIELD_UNKNOWN", codes(result))
        self.assertIn("size", " ".join(e["detail"] for e in result["errors"]))

    def test_unknown_color_token_is_blocked(self):
        deck = base_deck()
        deck["slides"][0]["elements"][2]["color"] = "brand_blue"
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("TOKEN_UNKNOWN", codes(result))

    def test_missing_theme_token_is_blocked(self):
        deck = base_deck()
        deck["theme"]["colors"].pop("accent")
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("TOKEN_UNKNOWN", codes(result))

    def test_three_font_families_are_blocked(self):
        deck = base_deck()
        deck["theme"]["fonts"]["display"] = "Helvetica"
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("FONT_FAMILIES", codes(result))

    def test_accent_budget_is_enforced(self):
        deck = base_deck()
        for element in deck["slides"][0]["elements"]:
            if element["type"] == "text":
                element["color"] = "accent"
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("ACCENT_OVER_BUDGET", codes(result))


class TestChartGuards(unittest.TestCase):
    def _chart_deck(self):
        deck = base_deck()
        deck["slides"][2]["elements"] = [e for e in deck["slides"][2]["elements"] if e["type"] == "chart"]
        chart = deck["slides"][2]["elements"][0]
        return deck, chart

    def test_series_length_mismatch_is_blocked(self):
        deck, chart = self._chart_deck()
        chart["series"][0]["values"] = chart["series"][0]["values"][:-1]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("CHART_INVALID", codes(result))

    def test_nan_value_is_blocked(self):
        deck, chart = self._chart_deck()
        chart["series"][0]["values"][0] = float("nan")
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("CHART_INVALID", codes(result))

    def test_unknown_chart_kind_is_blocked(self):
        deck, chart = self._chart_deck()
        chart["chart_kind"] = "sunburst"
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("CHART_INVALID", codes(result))

    def test_donut_with_two_series_is_blocked(self):
        deck, chart = self._chart_deck()
        chart["chart_kind"] = "donut"
        chart["series"].append({"name": "第二序列", "values": [1, 2, 3]})
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("CHART_INVALID", codes(result))

    def test_donut_not_summing_to_100_warns(self):
        deck, chart = self._chart_deck()
        chart["chart_kind"] = "donut"
        chart["series"] = [{"name": "量", "values": [10, 20, 30]}]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("DONUT_NOT_PERCENT", warnings(result))


class TestProvenanceGuards(unittest.TestCase):
    def test_release_without_period_is_blocked(self):
        deck = base_deck()
        for slide in deck["slides"]:
            slide["meta"]["period"] = ""
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("PROVENANCE_MISSING", codes(result))

    def test_draft_tolerates_missing_period_but_release_does_not(self):
        deck = base_deck()
        for slide in deck["slides"]:
            slide["meta"]["period"] = ""
        self.assertEqual(guard_check(copy.deepcopy(deck), POLICY, mode="draft")["status"], "PASS")
        self.assertEqual(guard_check(deck, POLICY, mode="release")["status"], "BLOCK")

    def test_metric_without_unit_is_blocked_in_release(self):
        deck = base_deck()
        for slide in deck["slides"]:
            for metric in slide["meta"].get("metrics") or []:
                metric["unit"] = ""
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("PROVENANCE_MISSING", codes(result))

    def test_multi_unit_page_satisfies_unit_by_per_metric_units(self):
        """并排指标页本来就没有单一页级单位：每个指标各自带单位即视为齐备。"""
        deck = base_deck()
        slide = deck["slides"][1]
        slide["meta"]["unit"] = ""
        slide["meta"]["metrics"] = [{"label": "门店数量", "value": "128", "unit": "家"},
                                    {"label": "单店产出降幅", "value": "21", "unit": "%"}]
        result = guard_check(deck, POLICY, mode="release")
        self.assertNotIn("'unit'", " ".join(e["detail"] for e in result["errors"]))

    def test_metric_units_across_pages_must_match(self):
        deck = base_deck()
        deck["meta"]["metric_units"] = {"经销渠道单店月均产出": ["万元", "元"]}
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("UNIT_INCONSISTENT", codes(result))

    def test_declared_source_without_element_is_blocked(self):
        deck = base_deck()
        slide = deck["slides"][0]
        slide["elements"] = [e for e in slide["elements"] if e.get("role") != "source"]
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("PROVENANCE_MISSING", codes(result))


class TestObservations(unittest.TestCase):
    def test_focus_competition_is_only_a_warning(self):
        deck = base_deck()
        claim = deck["slides"][0]["elements"][1]
        lede = deck["slides"][0]["elements"][2]
        lede["size"] = claim["size"]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("FOCUS_COMPETITION", warnings(result))
        self.assertNotIn("FOCUS_COMPETITION", codes(result))

    def test_duplicate_text_is_flagged(self):
        deck = base_deck()
        elements = deck["slides"][0]["elements"]
        elements[2]["text"] = elements[1]["text"]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("DUPLICATE_TEXT", warnings(result))

    def test_fix_groups_are_grouped_by_root_cause(self):
        deck = base_deck()
        for element in deck["slides"][0]["elements"]:
            element["y"] = 700
        result = guard_check(deck, POLICY, mode="draft")
        groups = {g["code"] for g in result["fix_groups"]}
        self.assertIn("OUT_OF_CANVAS", groups)
        group = [g for g in result["fix_groups"] if g["code"] == "OUT_OF_CANVAS"][0]
        self.assertGreater(group["count"], 1)
        self.assertTrue(group["hint"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
