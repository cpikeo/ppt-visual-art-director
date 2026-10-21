#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""真实产物：PPTX 编译、回读对账、渲染对照、字节确定性。

这一层证明的是「判断有没有被完整执行」，不是「设计好不好」。
所有断言都落在真实 .pptx 产物或其回读事实上。
"""
import copy
import hashlib
import unittest
import zipfile

import support
from support import ROOT, tmpdir

from policy import load as load_policy
from brief import load as load_brief
from judgment import judge_deck
from compose import compose_deck
from compile import compile_deck, readback
from fidelity import compare
from typography import FontBook
from render import page_model_from_readback, page_model_from_spec, pixel_diff, render_page

SCHEMA_DIR = ROOT / "schema"
POLICY = load_policy()
FONTS = FontBook(POLICY)


def make(tmp: str = None):
    tmp = tmpdir()
    brief, _ = load_brief(support.write_brief(tmp), POLICY, SCHEMA_DIR)
    deck = compose_deck(brief, judge_deck(brief, POLICY), POLICY, fontbook=FONTS)
    return deck, tmp


class TestCompileArtifact(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.deck, cls.tmp = make()
        cls.path = cls.tmp / "out.pptx"
        cls.report = compile_deck(cls.deck, cls.path)

    def test_pptx_is_a_real_package(self):
        self.assertTrue(self.path.exists())
        with zipfile.ZipFile(self.path) as package:
            names = package.namelist()
        self.assertIn("[Content_Types].xml", names)
        self.assertTrue(any(n.startswith("ppt/slides/slide") for n in names))
        self.assertTrue(any(n.startswith("ppt/charts/chart") for n in names),
                        "图表必须是原生 chart part，不能用图片冒充")

    def test_slide_count_matches_spec(self):
        self.assertEqual(self.report["counts"]["text"] > 0, True)
        self.assertEqual(self.report["counts"]["chart"], 1)
        rb = readback(self.path)
        self.assertEqual(len(rb["slides"]), len(self.deck["slides"]))

    def test_bytes_are_deterministic(self):
        second = self.tmp / "out2.pptx"
        second_report = compile_deck(self.deck, second)
        self.assertEqual(self.report["output_sha256"], second_report["output_sha256"])
        self.assertEqual(self.path.read_bytes(), second.read_bytes())

    def test_nbsp_survives_round_trip(self):
        rb = readback(self.path)
        texts = [s.get("text", "") for slide in rb["slides"] for s in slide["shapes"]]
        bound = [t for t in texts if "\u00a0" in t]
        self.assertTrue(bound, "数字与单位之间的不断行空格必须在 PPTX 里保留")
        self.assertTrue(any("万元" in t for t in bound))

    def test_east_asian_font_is_bound(self):
        """只写 latin 字体时中文会回落成系统默认字体，所以 ea 必须一起写。"""
        with zipfile.ZipFile(self.path) as package:
            xml = package.read("ppt/slides/slide1.xml").decode("utf-8")
        self.assertIn('typeface="Noto Sans CJK SC"', xml)

    def test_chart_data_is_editable(self):
        rb = readback(self.path)
        chart = [s for slide in rb["slides"] for s in slide["shapes"] if s.get("type") == "chart"][0]
        self.assertEqual(chart["kind"], "column")
        self.assertEqual(chart["categories"], ["2026 Q1", "2026 Q2", "2026 Q3"])
        self.assertEqual(chart["series"][0]["values"], [34.0, 43.0, 51.0])

    def test_shapes_are_native(self):
        rb = readback(self.path)
        kinds = {s["type"] for slide in rb["slides"] for s in slide["shapes"]}
        self.assertTrue(kinds <= {"text", "chart", "shape"}, kinds)
        self.assertIn("text", kinds)


class TestFidelity(unittest.TestCase):
    def test_readback_matches_spec(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        result = compare(deck, readback(path), POLICY)
        self.assertEqual(result["status"], "MATCH", result["mismatches"][:5])
        self.assertGreater(result["matched_elements"], 10)

    def test_drift_is_detected(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        rb = readback(path)
        rb["slides"][0]["shapes"][1]["x"] += 40
        result = compare(deck, rb, POLICY)
        self.assertEqual(result["status"], "MISMATCH")
        self.assertIn("drift", result["summary"])

    def test_dropped_element_is_detected(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        rb = readback(path)
        rb["slides"][0]["shapes"].pop()
        result = compare(deck, rb, POLICY)
        self.assertIn("dropped", result["summary"])

    def test_rewritten_text_is_detected(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        rb = readback(path)
        for shape in rb["slides"][0]["shapes"]:
            if shape.get("type") == "text":
                shape["text"] = shape["text"] + "X"
                break
        result = compare(deck, rb, POLICY)
        self.assertIn("payload", result["summary"])

    def test_chart_payload_change_is_detected(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        rb = readback(path)
        chart = [s for slide in rb["slides"] for s in slide["shapes"] if s.get("type") == "chart"][0]
        chart["series"][0]["values"][0] = 99.0
        result = compare(deck, rb, POLICY)
        self.assertIn("payload", result["summary"])


class TestRenderComparison(unittest.TestCase):
    def test_spec_and_readback_render_identically(self):
        deck, tmp = make()
        path = tmp / "deck.pptx"
        compile_deck(deck, path)
        rb = readback(path)
        for index in (0, 1, 2):
            a = render_page(page_model_from_spec(deck, deck["slides"][index]), tmp / f"a{index}.png",
                            POLICY, FONTS)
            b = render_page(page_model_from_readback(deck, rb["slides"][index], deck["slides"][index]),
                            tmp / f"b{index}.png", POLICY, FONTS)
            diff = pixel_diff(a, b, POLICY)
            self.assertEqual(diff["status"], "OK")
            self.assertLessEqual(diff["mismatch_ratio"], POLICY.num("thresholds.fidelity_pixel_mismatch_max"))

    def test_renderer_detects_a_real_difference(self):
        deck, tmp = make()
        model = page_model_from_spec(deck, deck["slides"][0])
        a = render_page(model, tmp / "x.png", POLICY, FONTS)
        model2 = copy.deepcopy(model)
        for item in model2["items"]:
            if item["kind"] == "text":
                item["color"] = "#FF0000"
        b = render_page(model2, tmp / "y.png", POLICY, FONTS)
        diff = pixel_diff(a, b, POLICY)
        self.assertGreater(diff["mismatch_ratio"], 0.0)

    def test_renderer_output_size_matches_canvas(self):
        from PIL import Image
        deck, tmp = make()
        path = render_page(page_model_from_spec(deck, deck["slides"][0]), tmp / "p.png", POLICY, FONTS)
        image = Image.open(path)
        self.assertEqual(image.size, (int(deck["canvas"]["width"]), int(deck["canvas"]["height"])))


if __name__ == "__main__":
    unittest.main(verbosity=2)
