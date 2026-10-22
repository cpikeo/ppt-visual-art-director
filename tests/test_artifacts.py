#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""真实产物：PPTX 编译、回读对账、渲染对照、字节确定性。

这一层证明的是「判断有没有被完整执行」，不是「设计好不好」。
所有断言都落在真实 .pptx 产物或其回读事实上。
"""
import copy
import tempfile
import unittest
import zipfile
from pathlib import Path

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

    def test_charts_use_the_deck_palette_not_the_policy_default(self):
        """图表配色跟稿件主题走。

        回归自案例 A：渲染器直接读 policy.colors()，成片里的柱子与扇区永远是策略默认
        色（#1F4E5A/#B4522A），brief 声明的四个主导色根本没进入图表；而 spec 与
        readback 两条路径都错成一样，像素差 0，于是「预览与交付一致」掩盖了颜色错。
        """
        from PIL import Image
        deck, tmp = make()
        deck["theme"]["colors"].update({"primary": "#5A1F2A", "accent": "#8C6B3F"})
        chart_slide = [s for s in deck["slides"] if any(e["type"] == "chart" for e in s["elements"])][0]
        path = render_page(page_model_from_spec(deck, chart_slide), tmp / "chart.png", POLICY, FONTS)
        image = Image.open(path).convert("RGB")

        def near(target, tolerance=16):
            return any(all(abs(pixel[i] - target[i]) <= tolerance for i in range(3))
                       for _, pixel in image.getcolors(maxcolors=1 << 24))

        self.assertTrue(near((0x5A, 0x1F, 0x2A)), "主题主色没有出现在图表里")
        self.assertFalse(near((0x1F, 0x4E, 0x5A)), "图表仍在用策略默认色 #1F4E5A")

    def test_bar_chart_is_horizontal_and_reads_bottom_up(self):
        """横向条形自下而上画——PowerPoint 的 BAR_CLUSTERED 就是这样排类的。

        这条契约曾经缺失：render 把 bar 当 column 画成竖柱，而编译出的 PPTX 是横向
        条形图，预览与交付根本不是同一张图（像素对照仍是 0，因为两条路径共用同一个错）。
        """
        from PIL import Image
        deck = {
            "canvas": {"width": 1280, "height": 720},
            "theme": {"colors": {"background": "#FFFFFF", "ink": "#000000", "muted": "#666666",
                                 "hairline": "#CCCCCC"}},
            "slides": [{"id": "s01", "background": "background", "elements": [
                {"id": "s01.chart", "type": "chart", "chart_kind": "bar", "x": 80, "y": 120,
                 "width": 720, "height": 420, "categories": ["A", "B", "C"],
                 "series": [{"name": "v", "values": [50, 150, 100]}], "show_values": False}]}],
        }
        tmp = Path(tempfile.mkdtemp())
        path = render_page(page_model_from_spec(deck, deck["slides"][0]), tmp / "bar.png", POLICY, FONTS)
        image = Image.open(path).convert("RGB")
        primary = POLICY.colors()["primary"].lstrip("#")
        colour = tuple(int(primary[i:i + 2], 16) for i in (0, 2, 4))
        width, _ = image.size
        runs = {}
        for y in range(image.size[1]):
            xs = [x for x in range(width) if image.getpixel((x, y)) == colour]
            if xs:
                runs[y] = max(xs)
        bands = []
        for y in sorted(runs):
            if bands and y - bands[-1][-1] <= 2:
                bands[-1].append(y)
            else:
                bands.append([y])
        widths = [max(runs[y] for y in band) for band in bands]
        self.assertEqual(len(widths), 3, "三个类别应各有一条横向条带")
        # 取值 50 / 150 / 100：自下而上，中间那条最长，最下面那条最短
        self.assertLess(widths[0], widths[1])
        self.assertLess(widths[2], widths[1])

    def test_renderer_output_size_matches_canvas(self):
        from PIL import Image
        deck, tmp = make()
        path = render_page(page_model_from_spec(deck, deck["slides"][0]), tmp / "p.png", POLICY, FONTS)
        with Image.open(path) as image:
            self.assertEqual(image.size, (int(deck["canvas"]["width"]), int(deck["canvas"]["height"])))


if __name__ == "__main__":
    unittest.main(verbosity=2)
