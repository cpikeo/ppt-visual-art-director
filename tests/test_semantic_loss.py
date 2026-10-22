#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""语义不丢失的契约：作者写了什么，产物里就必须有什么。

这一组用例来自第三轮真实案例对照里发现的四处「静默丢失」：
  * 案例 B s04/s10：作者声明了 metrics，构图为 image_narrative，数字在成片里全没了；
  * 案例 B s05：六道工序被 items[:5] 截断，第六道从产物里消失，guard 仍 PASS；
  * 案例 B s05/s07：环图分片、图表配色不跟主题（render 读策略默认调色板）；
  * 案例 B s05：页码「05」与工序序号「05」被判成 DUPLICATE_TEXT（定位符当成内容）。

它们共同的问题不是「不好看」，而是**产物与作者意图不一致且无人报告**。
测试通过≠设计优秀；但「内容悄悄少一条」必须在测试里被固定住。
"""
import shutil
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


def deck_for(text: str, *, with_asset: bool = False):
    tmp = tmpdir()
    if with_asset:
        target = tmp / "assets"
        target.mkdir(parents=True, exist_ok=True)
        shutil.copy(ROOT / "examples" / "assets" / "warehouse.jpg", target / "warehouse.jpg")
    brief, _ = load_brief(support.write_brief(tmp, text=text), POLICY, SCHEMA_DIR)
    return compose_deck(brief, judge_deck(brief, POLICY), POLICY)


def roles(slide: dict) -> list:
    return [str(e.get("role") or "") for e in slide.get("elements") or []]


IMAGE_WITH_METRICS = """\
audience: 消费品牌基金
decision: 是否参与 A 轮
tension: 市场把精品巧克力当甜食
source: 内部台账
unit: 万元
period: "2026"
basis: 已结算
slides:
  - id: s01
    title: 三个庄园，三种发酵曲线
    content: |-
      三家自有庄园分布在北纬 18 度带。
    source: 庄园记录
    metrics:
      - value: 3
        unit: 个
        label: 自有庄园
      - value: 640
        unit: 米
        label: 平均海拔
      - value: 6
        unit: 天
        label: 木箱发酵
    asset:
      path: assets/warehouse.jpg
      function: proof
      fit: cover
      subject: 仓库里的托盘与货架，自然光，无人
      safe_area: [0.5, 0.1, 0.44, 0.74]
      text_color: dark
"""

SEVEN_STEPS = """\
audience: 内部评审会
decision: 批准下一步预算
tension: 工序是否可以删减
source: 工艺手册
unit: 万元
period: "2026"
basis: 标准批次
slides:
  - id: s01
    title: 七道工序，一道都不能少
    content: |-
      1) 采收：只取成熟果
      2) 发酵：木箱 6 天
      3) 干燥：竹席 9 天
      4) 烘焙：单独曲线
      5) 精磨：石磨 48 小时
      6) 调温：31℃ 恒温
      7) 入模：静置 12 小时
    source: 工艺手册
"""


class TestDeclaredEvidenceIsLaidOut(unittest.TestCase):
    def test_image_page_still_carries_the_declared_metrics(self):
        """影像页不能吞掉作者声明的数字：图负责看见，数字负责可核验。"""
        deck = deck_for(IMAGE_WITH_METRICS, with_asset=True)
        slide = deck["slides"][0]
        self.assertEqual(slide["composition"], "image_narrative")
        values = [r for r in roles(slide) if r.startswith("metric_value")]
        self.assertEqual(len(values), 3, roles(slide))
        self.assertIn("media", roles(slide))

    def test_guard_blocks_when_declared_metrics_are_missing(self):
        deck = deck_for(IMAGE_WITH_METRICS, with_asset=True)
        slide = deck["slides"][0]
        slide["elements"] = [e for e in slide["elements"]
                             if not str(e.get("role") or "").startswith("metric_value")]
        result = guard_check(deck, POLICY, mode="draft")
        self.assertEqual(result["status"], "BLOCK")
        self.assertIn("EVIDENCE_NOT_LAID_OUT", {e["code"] for e in result["errors"]})

    def test_two_end_values_count_as_laid_out(self):
        deck = deck_for("""\
audience: 内部评审会
decision: 批准换线预算
tension: 产能是不是瓶颈
source: 生产台账
unit: 吨
period: "2026"
basis: 成品产出
slides:
  - id: s01
    title: 产能不是野心，是交付节奏
    content: |-
      现有工坊年产 180 吨，恒温工坊投产后年产 320 吨。
    source: 生产台账
    composition: two_end
    metrics:
      - value: 180
        unit: 吨
        label: 现有年产
      - value: 320
        unit: 吨
        label: 投产后年产
""")
        slide = deck["slides"][0]
        self.assertIn("left_value", roles(slide))
        result = guard_check(deck, POLICY, mode="draft")
        self.assertNotIn("EVIDENCE_NOT_LAID_OUT", {e["code"] for e in result["errors"]})


class TestStepOverflowIsReported(unittest.TestCase):
    def test_overflow_is_capped_by_policy_and_reported(self):
        """列数上限来自策略；超出的工序必须被记账，不能消失。"""
        deck = deck_for(SEVEN_STEPS)
        slide = deck["slides"][0]
        limit = int(POLICY.num("thresholds.max_steps"))
        laid_out = [r for r in roles(slide) if r.startswith("step_title")]
        self.assertEqual(len(laid_out), limit)
        self.assertEqual(slide["meta"]["steps_dropped"], 7 - limit)
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("CONTENT_TRUNCATED", {w["code"] for w in result["warnings"]})

    def test_within_limit_is_not_reported(self):
        deck = deck_for(SEVEN_STEPS.replace("      7) 入模：静置 12 小时\n", ""))
        slide = deck["slides"][0]
        self.assertEqual(slide["meta"]["steps_dropped"], 0)
        result = guard_check(deck, POLICY, mode="draft")
        self.assertNotIn("CONTENT_TRUNCATED", {w["code"] for w in result["warnings"]})


class TestLocatorTextIsNotContent(unittest.TestCase):
    STATEMENT = """\
audience: 内部评审会
decision: 批准换线
tension: 批次一致性
source: 生产台账
unit: 万元
period: "2026"
basis: 标准批次
slides:
  - id: s01
    title: 编号 05 的批次是最好的证据
    claim: 一致性来自记录
    role: explain
    composition: statement
    content: |-
      编号 05 的批次记录完整，可追溯到每一箱。
    source: 生产台账
"""

    def test_page_number_matching_a_body_number_is_not_a_duplicate(self):
        """页码是定位符，不是内容：它与页面里的同名数字不构成重复信息。

        案例 B s05 的回归：页码「05」与工序序号「05」被判 DUPLICATE_TEXT。
        定位符不参与内容重复判定，也不计入「数字证据」密度。
        """
        deck = deck_for(self.STATEMENT)
        slide = deck["slides"][0]
        eyebrow = next(e for e in slide["elements"] if e.get("role") == "eyebrow")
        eyebrow["text"] = "05"
        body = next(e for e in slide["elements"] if e.get("role") == "lede")
        self.assertIn("05", body["text"])
        result = guard_check(deck, POLICY, mode="draft")
        duplicates = [w for w in result["warnings"] if w["code"] == "DUPLICATE_TEXT"]
        self.assertEqual(duplicates, [], duplicates)

    def test_same_number_in_two_body_elements_is_still_a_warning(self):
        """定位豁免只给定位符：正文里重复同一个数字仍要报。"""
        deck = deck_for(self.STATEMENT)
        slide = deck["slides"][0]
        body = next(e for e in slide["elements"] if e.get("role") == "lede")
        twin = dict(body, id="s01.twin", role="note", y=body["y"] + 96)
        slide["elements"].append(twin)
        result = guard_check(deck, POLICY, mode="draft")
        self.assertIn("DUPLICATE_TEXT", {w["code"] for w in result["warnings"]})


if __name__ == "__main__":
    unittest.main(verbosity=2)
