#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""判断协议回归：每条用例都对应 SKILL.md 里的一条判断纪律。

这些用例不证明构图或美感，只证明「判断走的是内容，不是表面词汇」，
以及「判不出来时如实记 absent，而不是编一个结论」。
"""
import unittest

import support
from support import ROOT, tmpdir, write_brief

from policy import load as load_policy
from brief import load as load_brief
from judgment import judge_deck

SCHEMA_DIR = ROOT / "schema"
BASE = ("audience: 管理层\ndecision: 批准预算\ntension: 没有共识\n"
        "source: 内部看板\nunit: 万元\nperiod: 2026 Q1\nbasis: 已开票收入\n")


def plan_for(slides_yaml: str):
    policy = load_policy()
    brief, _ = load_brief(write_brief(tmpdir(), text=BASE + slides_yaml), policy, SCHEMA_DIR)
    return judge_deck(brief, policy), policy


class TestClaim(unittest.TestCase):
    def test_author_claim_wins_over_body(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 交付周期
    claim: 交付周期必须减半
    content: 交付平均 42 天。
""")
        page = plan["pages"][0]
        self.assertEqual(page["claim"]["source"], "author")
        self.assertEqual(page["claim"]["text"], "交付周期必须减半")

    def test_topic_title_is_not_promoted_to_claim(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 渠道结构
    content: |-
      经销 51%
      直营 34%
      电商 15%
""")
        self.assertEqual(plan["pages"][0]["claim"]["source"], "absent")
        codes = [q["code"] for q in plan["pages"][0]["open_questions"]]
        self.assertIn("CLAIM_ABSENT", codes)

    def test_body_conclusion_sentence_can_become_claim(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 渠道复盘
    content: |-
      上半年投入增加 12%，收入只增加 3%。
      因此继续加店的边际收益已经不成立。
""")
        page = plan["pages"][0]
        self.assertEqual(page["claim"]["source"], "derived")
        self.assertIn("边际收益", page["claim"]["text"])

    def test_cover_meta_line_does_not_change_role(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 发布周期可以再压缩一半
    content: 2026-10-21 内部评审 版本 v2
  - id: s02
    title: 请批准下一步
    content: 决定：批准。
""")
        self.assertEqual(plan["pages"][0]["decision"]["role"], "establish")


class TestEvidenceShape(unittest.TestCase):
    def test_percentages_summing_to_100_are_not_part_whole(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 用药结果
    content: |-
      缓解率 60%
      不良反应率 40%
      两项相加为 100%，但同一患者可同时满足。
""")
        page = plan["pages"][0]
        self.assertEqual(page["understanding"]["evidence"]["value"], "comparison")
        self.assertFalse(page["understanding"]["denominator"]["proven"])
        self.assertEqual(page["decision"]["composition"], "metric_row")
        self.assertTrue(any(a["rejected_because"] for a in page["decision"]["composition_alternatives"]))

    def test_real_part_whole_with_shared_denominator(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 收入构成
    content: |-
      经销 51%
      直营 34%
      电商 15%
""")
        page = plan["pages"][0]
        self.assertEqual(page["understanding"]["evidence"]["value"], "part_whole")
        self.assertEqual(page["decision"]["composition"], "part_whole")

    def test_yoy_single_metric_is_not_a_series(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 单店产出同比下滑 21%
    content: 单店月均产出同比下降 21%。
""")
        page = plan["pages"][0]
        self.assertNotEqual(page["understanding"]["evidence"]["value"], "series")
        self.assertIn(page["understanding"]["evidence"]["value"], ("comparison", "number", "proposition"))

    def test_parallel_duties_are_not_a_sequence(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 两个组的职责
    content: |-
      青禾组：负责质量测试
      澄海组：负责项目交付
""")
        page = plan["pages"][0]
        self.assertNotEqual(page["understanding"]["evidence"]["value"], "sequence")
        self.assertNotEqual(page["decision"]["composition"], "sequence_steps")

    def test_real_sequence_with_order_words(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 上线分三步
    content: |-
      第一步：冻结需求
      第二步：小流量灰度
      第三步：全量发布
""")
        self.assertEqual(plan["pages"][0]["understanding"]["evidence"]["value"], "sequence")
        self.assertEqual(plan["pages"][0]["decision"]["composition"], "sequence_steps")

    def test_two_ends_with_delta_only_when_same_unit(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 经销渠道单店产出高出直营 2.4 倍
    content: |-
      4.6
      经销渠道单店月均产出
      万元
      1.9
      直营门店单店月均产出
      万元
""")
        page = plan["pages"][0]
        self.assertTrue(page["understanding"]["comparison"]["delta"])
        self.assertEqual(page["decision"]["composition"], "two_end")

    def test_mixed_units_do_not_produce_delta(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 门店与产出
    content: |-
      128
      门店数量
      家
      4.6
      单店月产出
      万元
""")
        page = plan["pages"][0]
        self.assertFalse(page["understanding"]["comparison"]["delta"])
        self.assertEqual(page["decision"]["composition"], "metric_row")

    def test_author_chart_sets_shape(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 占比逐季抬升
    content: Q1 到 Q3 从 34% 升到 51%。
    chart:
      kind: column
      categories: ["2026 Q1", "2026 Q2", "2026 Q3"]
      series:
        - name: 占比
          values: [34, 43, 51]
""")
        self.assertEqual(plan["pages"][0]["understanding"]["evidence"]["value"], "series")
        self.assertEqual(plan["pages"][0]["decision"]["composition"], "series_chart")


class TestDeckLevel(unittest.TestCase):
    def test_missing_deck_fields_are_warned(self):
        policy = load_policy()
        brief, _ = load_brief(write_brief(tmpdir(), text="slides:\n  - id: s01\n    title: 一句话主张\n"),
                              policy, SCHEMA_DIR)
        plan = judge_deck(brief, policy)
        codes = {w["code"] for w in plan["deck"]["warnings"]}
        self.assertIn("DECK_FIELD_EMPTY", codes)

    def test_climax_and_rhythm_are_derived(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 加店不再带来增长
    content: 门店增长 18%，单店产出下降 21%。
  - id: s02
    title: 经销渠道单店产出高出直营 2.4 倍
    content: |-
      4.6
      经销渠道单店月均产出
      万元
      1.9
      直营门店单店月均产出
      万元
  - id: s03
    title: 请批准 Q1 预算调整
    content: 决定：批准。
""")
        moves = [p["decision"]["rhythm"]["move"] for p in plan["pages"]]
        self.assertEqual(moves[0], "open")
        self.assertEqual(moves[-1], "close")
        self.assertEqual(plan["deck"]["coherence"]["climax"], "s02")
        self.assertEqual(plan["pages"][1]["decision"]["emphasis"], "lead")

    def test_plan_is_deterministic(self):
        text = BASE + """slides:
  - id: s01
    title: 门店在增，产出在跌
    content: |-
      128
      门店数量
      家
      21%
      单店产出降幅
"""
        first, _ = plan_for(text.replace(BASE, ""))
        second, _ = plan_for(text.replace(BASE, ""))
        import json
        self.assertEqual(json.dumps(first, ensure_ascii=False, sort_keys=True),
                         json.dumps(second, ensure_ascii=False, sort_keys=True))


class TestProbes(unittest.TestCase):
    def test_causal_language_produces_a_question_not_a_change(self):
        plan, _ = plan_for("""slides:
  - id: s01
    title: 因此我们应当转向
    content: 门店上升因此产出下降。
""")
        page = plan["pages"][0]
        codes = [q["code"] for q in page["open_questions"]]
        self.assertIn("PROBE_CAUSAL_RISK", codes)
        self.assertNotIn("sequence", page["understanding"]["evidence"]["value"])

    def test_missing_provenance_is_asked(self):
        policy = load_policy()
        brief, _ = load_brief(write_brief(tmpdir(), text="slides:\n  - id: s01\n    title: 收入 4.6 万元\n"
                                                          "    content: |-\n      4.6\n      收入\n"),
                              policy, SCHEMA_DIR)
        plan = judge_deck(brief, policy)
        codes = [q["code"] for q in plan["pages"][0]["open_questions"]]
        self.assertIn("PROVENANCE_MISSING", codes)


if __name__ == "__main__":
    unittest.main(verbosity=2)
