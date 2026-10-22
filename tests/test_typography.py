#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文排版度量：数字与单位绑定、禁则、墨迹几何、断行偏好。

这些是确定性的机器可判事实；真实 PowerPoint 的字体回落与断行不在其中（见 SKILL「中文排版」「证据层级」）。
"""
import unittest

import support
from policy import load as load_policy
from typography import (FontBook, bind_units, ink_height, ink_top, orphan_issues, reading_seconds,
                        tokenize, wrap)

POLICY = load_policy()
FONTS = FontBook(POLICY)


class TestNumberUnitBinding(unittest.TestCase):
    def test_binding_inserts_nbsp(self):
        self.assertIn("\u00a0", bind_units("单店月均产出 186 公里", POLICY))
        self.assertIn("\u00a0", bind_units("占比 34%", POLICY))

    def test_unit_is_never_split_from_number(self):
        text = bind_units("单店月均产出为 186 公里，需要复核", POLICY)
        for width in (60, 90, 120, 180):
            metrics = wrap(text, width, 22, policy=POLICY, fontbook=FONTS)
            for line in metrics.lines:
                self.assertFalse(line.rstrip().endswith("186\u00a0"), f"数字与单位被拆开：{metrics.lines}")
                self.assertFalse(line.lstrip().startswith("公里"), f"单位落在行首：{metrics.lines}")

    def test_multi_char_unit_is_one_token(self):
        tokens = tokenize("成本 4.6\u00a0万元", POLICY)
        joined = [tok for tok in tokens if "万元" in tok]
        self.assertTrue(joined, tokens)          # 「万元」整体出现，而不是「万」+「元」
        self.assertNotIn("万", tokens)
        self.assertNotIn("元", tokens)

    def test_punctuation_is_not_swallowed_by_numbers(self):
        tokens = tokenize("增长 18%，然后下滑", POLICY)
        self.assertIn("18%", tokens)
        self.assertIn("，", tokens)


class TestKinsoku(unittest.TestCase):
    def test_line_never_starts_with_closing_punctuation(self):
        text = "第一，门店在增，产出在跌；第二，渠道结构变化明显。" * 3
        for width in (100, 160, 220, 320):
            metrics = wrap(text, width, 20, policy=POLICY, fontbook=FONTS)
            for i, line in enumerate(metrics.lines):
                if i:
                    self.assertNotIn(line[0], "，。、；：？！）", f"行首禁则失败：{metrics.lines}")

    def test_orphan_line_is_reported(self):
        metrics = wrap("这是一句会被截得很短的话啊", 60, 20, policy=POLICY, fontbook=FONTS)
        kinds = {issue.kind for issue in orphan_issues(metrics, POLICY)}
        self.assertTrue(kinds, f"应至少报出一条孤字/禁则问题：{metrics.lines}")

    def test_prefers_breaking_after_punctuation(self):
        text = "门店数量增长 18%，单店月均产出下降 21%，渠道结构正在变化"
        metrics = wrap(text, 260, 22, policy=POLICY, fontbook=FONTS)
        self.assertGreater(metrics.count, 1)
        self.assertTrue(any(line.rstrip().endswith("，") for line in metrics.lines[:-1]),
                        metrics.lines)


class TestInkGeometry(unittest.TestCase):
    def test_single_line_ink_is_not_zero(self):
        height = ink_height(1, 88, 1.06, POLICY)
        self.assertGreaterEqual(height, 88 * 1.2 - 0.01)

    def test_ink_grows_with_lines(self):
        self.assertLess(ink_height(2, 60, 1.12, POLICY), ink_height(3, 60, 1.12, POLICY))

    def test_bottom_anchor_moves_ink_up(self):
        top = ink_top(400, 200, 2, 60, 1.1, "bottom", POLICY)
        self.assertAlmostEqual(top + ink_height(2, 60, 1.1, POLICY), 600, places=3)

    def test_frame_height_does_not_inflate_ink(self):
        """框大不等于墨迹大：底对齐时墨迹贴框底，向上展开。"""
        top_tall = ink_top(100, 400, 1, 40, 1.4, "bottom", POLICY)
        top_short = ink_top(100, 100, 1, 40, 1.4, "bottom", POLICY)
        self.assertGreater(top_tall, top_short)


class TestSemanticLadder(unittest.TestCase):
    def test_tracking_shrinks_for_display_and_opens_for_small_text(self):
        self.assertLess(POLICY.tracking("display"), POLICY.tracking("body"))

    def test_max_lines_are_declared_per_ladder(self):
        for key in ("display", "h1", "h2", "body"):
            self.assertGreaterEqual(POLICY.max_lines(key), 1)

    def test_reading_seconds_is_deterministic(self):
        self.assertEqual(reading_seconds("门店增长 18%", POLICY), reading_seconds("门店增长 18%", POLICY))


class TestDegradation(unittest.TestCase):
    def test_missing_fonts_are_reported_not_silent(self):
        policy = load_policy()
        policy.strategy["typography"]["font_files"] = {"cn": ["/nope/cn.ttf"], "cn_bold": ["/nope/b.ttf"],
                                                      "latin": ["/nope/l.ttf"]}
        book = FontBook(policy)
        self.assertTrue(book.degraded)
        metrics = wrap("没有字体时也要能给出估算", 100, 20, policy=policy, fontbook=book)
        self.assertTrue(metrics.degraded)


if __name__ == "__main__":
    unittest.main(verbosity=2)
