#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import unittest

from compose import _append_part_whole_chart
from policy import load


class TestDonutLegend(unittest.TestCase):
    def test_manual_legend_uses_same_palette_as_donut_points(self):
        elements = []

        def add_text_box(role, text, box, style, color, **kwargs):
            elements.append({"id": f"s01.{role}", "type": "text", "role": role,
                             "text": text, "x": box[0], "y": box[1],
                             "width": box[2], "height": box[3], "color": color})

        chart = {"kind": "donut", "categories": ["A", "B", "C", "D"],
                 "series": [{"name": "mix", "values": [40, 25, 20, 15]}]}
        _append_part_whole_chart(elements, "s01.chart", (72, 216, 520, 384),
                                 chart, load(), add_text_box, None)
        chart_element = next(e for e in elements if e["type"] == "chart")
        keys = [e for e in elements if str(e.get("role", "")).startswith("legend_key")]
        labels = [e for e in elements if str(e.get("role", "")).startswith("legend_label")]
        self.assertFalse(chart_element["show_legend"])
        self.assertEqual([e["color"] for e in keys], ["primary", "accent", "muted", "ink"])
        self.assertEqual([e["text"] for e in labels], chart["categories"])
        self.assertTrue(all(e["text"] == "■" for e in keys))


if __name__ == "__main__":
    unittest.main(verbosity=2)
