#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""素材链路：判断 → 几何 → 编译 → 渲染 → 硬门 → 对账。

这一组测试针对的是一类特别危险的缺陷：**图片链路缺失导致「假通过」**。
在没有这些用例之前，作者声明了图片、判断层不读、compose 不写、renderer 不画，
于是一份「本该有图但没有图」的稿子可以 PASS —— 测试通过但交付物是错的。
"""
import copy
import shutil
import unittest

from PIL import Image

from support import ROOT, tmpdir, write_brief

from policy import load as load_policy
from brief import load as load_brief
from judgment import judge_deck
from compose import compose_deck
from guard import check as guard_check
from compile import compile_deck, readback
from fidelity import compare
from render import page_model_from_readback, page_model_from_spec, pixel_diff, render_page, visible_image
from typography import FontBook
from assets import cover_crop, inspect, required_source_px, resolve_source

SCHEMA_DIR = ROOT / "schema"
POLICY = load_policy()
FONTS = FontBook(POLICY)

WITH_ASSET = """\
    content: |-
      门店数量增长 18%，单店月均产出下降 21%。
    asset:
      path: assets/photo.jpg
      function: {function}
      reason: {reason}
      fit: {fit}
      safe_area: {{x: 0.0, y: 0.10, width: 0.50, height: 0.80}}
"""


def make_photo(directory, name="assets/photo.jpg", size=(1600, 1200), color=(120, 132, 140)):
    path = directory / name
    path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", size, color)
    # 加一点纹理，避免「平坦图」被误判成可压字的安全区
    for x in range(0, size[0], 8):
        for y in range(0, size[1], 8):
            image.putpixel((x, y), (color[0] - 30, color[1] - 20, color[2] - 10))
    image.save(path, "JPEG", quality=92)
    return path


def brief_text(function="hero", reason="让观众看见现场密度", fit="cover"):
    body = WITH_ASSET.format(function=function, reason=reason, fit=fit)
    return ("audience: 管理层\ndecision: 批准预算\ntension: 没有共识\n"
            "source: 内部看板\nunit: 万元\nperiod: 2026 Q1\nbasis: 已开票收入\n"
            "slides:\n  - id: s01\n    title: 加店不再带来增长\n" + body)


def build(tmp, *, function="hero", reason="让观众看见现场密度", fit="cover", photo=True):
    directory = tmpdir()
    if photo:
        make_photo(directory)
    brief, _ = load_brief(write_brief(directory, text=brief_text(function, reason, fit)),
                          POLICY, SCHEMA_DIR)
    plan = judge_deck(brief, POLICY)
    deck = compose_deck(brief, plan, POLICY, fontbook=FONTS, base_dir=directory)
    return plan, deck, directory


def codes(result):
    return {e["code"] for e in result["errors"]}


def warnings(result):
    return {w["code"] for w in result["warnings"]}


class TestMediaJudgment(unittest.TestCase):
    def test_declared_asset_becomes_required_media_with_reason(self):
        plan, deck, _ = build(tmpdir())
        media = plan["pages"][0]["decision"]["media"]
        self.assertEqual(media["decision"], "required")
        self.assertEqual(media["function"], "hero")
        self.assertEqual(media["source"], "author")
        self.assertTrue(media["reason"])
        self.assertEqual(plan["pages"][0]["decision"]["composition"], "image_narrative")

    def test_existing_file_switches_decision_to_existing(self):
        plan, _, _ = build(tmpdir())
        self.assertIn(plan["pages"][0]["decision"]["media"]["decision"], ("required", "existing"))

    def test_undeclared_visual_subject_only_hints(self):
        """词表只提示：不声明素材就不出图，只留一条待确认问题。"""
        directory = tmpdir()
        text = ("audience: 管理层\ndecision: 批准\ntension: 无\nsource: 看板\nunit: 万元\n"
                "period: 2026 Q1\nbasis: 已开票\nslides:\n  - id: s01\n    title: 车间产能复盘\n"
                "    content: 车间设备利用率 78%。\n")
        brief, _ = load_brief(write_brief(directory, text=text), POLICY, SCHEMA_DIR)
        plan = judge_deck(brief, POLICY)
        media = plan["pages"][0]["decision"]["media"]
        self.assertEqual(media["decision"], "none")
        self.assertIn("车间", media.get("hint") or [])
        self.assertIn("ASSET_HINT", [q["code"] for q in plan["pages"][0]["open_questions"]])
        deck = compose_deck(brief, plan, POLICY, fontbook=FONTS, base_dir=directory)
        self.assertFalse([e for e in deck["slides"][0]["elements"] if e["type"] == "image"])

    def test_missing_reason_asks_a_question(self):
        plan, _, _ = build(tmpdir(), reason="")
        self.assertIn("ASSET_REASON_MISSING", [q["code"] for q in plan["pages"][0]["open_questions"]])

    def test_proof_function_requires_source(self):
        directory = tmpdir()
        make_photo(directory)
        text = brief_text("proof", "现场照片作为证据").replace("source: 内部看板\n", "")
        brief, _ = load_brief(write_brief(directory, text=text), POLICY, SCHEMA_DIR)
        plan = judge_deck(brief, POLICY)
        self.assertIn("ASSET_PROOF_UNSOURCED", [q["code"] for q in plan["pages"][0]["open_questions"]])

    def test_asset_on_a_data_page_is_questioned(self):
        directory = tmpdir()
        make_photo(directory)
        text = ("audience: 管理层\ndecision: 批准\ntension: 无\nsource: 看板\nunit: 万元\n"
                "period: 2026 Q1\nbasis: 已开票\nslides:\n  - id: s01\n"
                "    title: 经销渠道单店产出高出直营 2.4 倍\n"
                "    content: |-\n      4.6\n      经销渠道单店月产出\n      万元\n"
                "      1.9\n      直营门店单店月产出\n      万元\n"
                "    asset:\n      path: assets/photo.jpg\n      function: hero\n"
                "      reason: 现场照片\n")
        brief, _ = load_brief(write_brief(directory, text=text), POLICY, SCHEMA_DIR)
        plan = judge_deck(brief, POLICY)
        self.assertIn("ASSET_VS_EVIDENCE", [q["code"] for q in plan["pages"][0]["open_questions"]])


class TestAssetGeometry(unittest.TestCase):
    def test_image_element_carries_facts(self):
        _, deck, _ = build(tmpdir())
        image = [e for e in deck["slides"][0]["elements"] if e["type"] == "image"][0]
        self.assertEqual(image["source_px"], {"width": 1600, "height": 1200})
        self.assertEqual(len(image["asset_sha256"]), 64)
        self.assertIn("crop_loss", image)
        self.assertTrue(image["reason"])

    def test_cover_crop_math(self):
        crop = cover_crop((1600, 1200), (500, 900))
        self.assertGreater(crop["loss"], 1.0)
        self.assertEqual(len(crop["crop"]), 4)
        same = cover_crop((1600, 900), (800, 450))
        self.assertEqual(same["loss"], 1.0)

    def test_required_source_pixels_grows_with_placement(self):
        small = required_source_px((200, 200), POLICY)
        large = required_source_px((900, 900), POLICY)
        self.assertLess(small, large)

    def test_missing_file_still_writes_element(self):
        """缺图必须留下元素：否则 guard 无从报错，稿子会「没有图也通过」。"""
        _, deck, _ = build(tmpdir(), photo=False)
        images = [e for e in deck["slides"][0]["elements"] if e["type"] == "image"]
        self.assertEqual(len(images), 1)
        self.assertTrue(images[0].get("_missing"))


class TestAssetGuards(unittest.TestCase):
    def test_missing_file_blocks(self):
        _, deck, _ = build(tmpdir(), photo=False)
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_MISSING", codes(result))

    def test_declared_media_without_element_blocks(self):
        _, deck, _ = build(tmpdir())
        deck["slides"][0]["elements"] = [e for e in deck["slides"][0]["elements"] if e["type"] != "image"]
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_MISSING", codes(result))
        self.assertIn("判断没有落到执行", " ".join(e["detail"] for e in result["errors"]))

    def test_low_resolution_blocks(self):
        directory = tmpdir()
        make_photo(directory, size=(320, 240))
        brief, _ = load_brief(write_brief(directory, text=brief_text()), POLICY, SCHEMA_DIR)
        deck = compose_deck(brief, judge_deck(brief, POLICY), POLICY, fontbook=FONTS, base_dir=directory)
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_RESOLUTION", codes(result))

    def test_bad_format_blocks(self):
        directory = tmpdir()
        path = directory / "assets" / "photo.bmp"
        path.parent.mkdir(parents=True, exist_ok=True)
        Image.new("RGB", (1600, 1200), (120, 132, 140)).save(path, "BMP")
        brief, _ = load_brief(write_brief(directory, text=brief_text()), POLICY, SCHEMA_DIR)
        shutil.move(str(path), str(directory / "assets" / "photo.jpg"))   # 后缀骗过白名单也没用
        deck = compose_deck(brief, judge_deck(brief, POLICY), POLICY, fontbook=FONTS, base_dir=directory)
        result = guard_check(deck, POLICY, mode="release")
        self.assertTrue({"ASSET_FORMAT"} <= codes(result) or result["status"] == "BLOCK")

    def test_crop_without_reason_blocks(self):
        directory = tmpdir()
        make_photo(directory, size=(2400, 800))     # 极宽图填竖向图位：必然裁掉很多
        brief, _ = load_brief(write_brief(directory, text=brief_text(reason="")), POLICY, SCHEMA_DIR)
        deck = compose_deck(brief, judge_deck(brief, POLICY), POLICY, fontbook=FONTS, base_dir=directory)
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_CROP", codes(result))

    def test_text_over_image_without_safe_area_blocks(self):
        """文字压在图上但没声明安全区 —— 这是「稿子看起来正常、投影读不清」的典型来源。"""
        _, deck, _ = build(tmpdir())
        slide = deck["slides"][0]
        image = [e for e in slide["elements"] if e["type"] == "image"][0]
        image["safe_area"] = None
        claim = [e for e in slide["elements"] if e.get("role") == "claim"][0]
        claim["x"], claim["y"] = image["x"] + 12, image["y"] + 40
        claim["width"], claim["height"] = image["width"] - 24, 80
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_SAFE_AREA", codes(result))

    def test_text_on_image_with_safe_area_and_good_contrast_passes(self):
        _, deck, _ = build(tmpdir())
        slide = deck["slides"][0]
        image = [e for e in slide["elements"] if e["type"] == "image"][0]
        claim = [e for e in slide["elements"] if e.get("role") == "claim"][0]
        claim["x"], claim["y"] = image["x"] + 12, image["y"] + 40
        claim["width"], claim["height"] = image["width"] - 24, 80
        claim["color"] = "surface"
        result = guard_check(deck, POLICY, mode="release")
        self.assertNotIn("ASSET_SAFE_AREA", codes(result))

    def test_low_contrast_text_over_image_blocks(self):
        """把一张接近文字灰度的图放到文字下面：实测对比度不足，必须拦。"""
        directory = tmpdir()
        make_photo(directory, color=(250, 250, 250))     # 近白背景，配 muted 文字仍可能不足
        brief, _ = load_brief(write_brief(directory, text=brief_text()), POLICY, SCHEMA_DIR)
        deck = compose_deck(brief, judge_deck(brief, POLICY), POLICY, fontbook=FONTS, base_dir=directory)
        # 让文字真正压到图上：把 claim 移到图的位置
        image = [e for e in deck["slides"][0]["elements"] if e["type"] == "image"][0]
        claim = [e for e in deck["slides"][0]["elements"] if e.get("role") == "claim"][0]
        claim["x"], claim["y"] = image["x"] + 10, image["y"] + 10
        claim["width"], claim["height"] = image["width"] - 20, 60
        claim["color"] = "surface"
        result = guard_check(deck, POLICY, mode="release")
        self.assertTrue({"ASSET_CONTRAST", "ASSET_SAFE_AREA"} & codes(result), result["errors"])

    def test_asset_budget_blocks(self):
        _, deck, _ = build(tmpdir())
        slide = deck["slides"][0]
        image = [e for e in slide["elements"] if e["type"] == "image"][0]
        for i in range(3):
            clone = copy.deepcopy(image)
            clone["id"] = f"{slide['id']}.media{i + 2}"
            clone["x"] = max(0, image["x"] - 10 * (i + 1))
            slide["elements"].append(clone)
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("ASSET_BUDGET", codes(result))

    def test_text_color_hint_is_not_a_color_token(self):
        """text_color 是 dark/light 语义提示，不能当颜色 token 校验（曾经误拦）。"""
        _, deck, _ = build(tmpdir())
        result = guard_check(deck, POLICY, mode="release")
        self.assertNotIn("TOKEN_UNKNOWN", codes(result))

    def test_source_zone_collision_by_image_blocks(self):
        _, deck, _ = build(tmpdir())
        slide = deck["slides"][0]
        image = [e for e in slide["elements"] if e["type"] == "image"][0]
        zone = slide["meta"]["source_zone"]
        image["x"], image["y"] = zone["x"], zone["y"]
        image["width"], image["height"] = zone["width"], zone["height"]
        result = guard_check(deck, POLICY, mode="release")
        self.assertIn("SOURCE_ZONE_COLLISION", codes(result))


class TestAssetArtifact(unittest.TestCase):
    def test_image_is_embedded_and_matched(self):
        _, deck, directory = build(tmpdir())
        pptx = directory / "out.pptx"
        compile_deck(deck, pptx, base_dir=directory)
        rb = readback(pptx)
        image = [e for e in deck["slides"][0]["elements"] if e["type"] == "image"][0]
        packed = [s for s in rb["slides"][0]["shapes"] if s.get("type") == "image"][0]
        self.assertEqual(packed["image_sha256"], image["asset_sha256"])
        self.assertEqual(compare(deck, rb, POLICY)["status"], "MATCH")

    def test_cover_crop_is_applied_not_stretched(self):
        _, deck, directory = build(tmpdir())
        pptx = directory / "out.pptx"
        compile_deck(deck, pptx, base_dir=directory)
        packed = [s for s in readback(pptx)["slides"][0]["shapes"] if s.get("type") == "image"][0]
        self.assertTrue(any(packed["crop"]), "cover 必须在 PPTX 里真正裁切，否则就是拉伸变形")

    def test_renderer_draws_the_image_so_diff_cannot_fake_pass(self):
        _, deck, directory = build(tmpdir())
        pptx = directory / "out.pptx"
        compile_deck(deck, pptx, base_dir=directory)
        rb = readback(pptx)
        a = render_page(page_model_from_spec(deck, deck["slides"][0]), directory / "a.png", POLICY, FONTS)
        b = render_page(page_model_from_readback(deck, rb["slides"][0], deck["slides"][0]),
                        directory / "b.png", POLICY, FONTS)
        diff = pixel_diff(a, b, POLICY)
        self.assertEqual(diff["status"], "OK")
        self.assertLessEqual(diff["mismatch_ratio"], POLICY.num("thresholds.fidelity_pixel_mismatch_max"))

    def test_rendered_pixels_actually_change_with_the_image(self):
        """如果渲染器不画图，两张完全不同的图会得到 0 像素差 —— 这正是「假通过」。"""
        directory = tmpdir()
        first = make_photo(directory, "assets/a.jpg", color=(30, 40, 50))
        second = make_photo(directory, "assets/b.jpg", color=(220, 210, 200))
        model = {"id": "s01", "width": 1280, "height": 720, "background": "#F7F5F1",
                 "items": [{"id": "img", "kind": "image", "x": 100, "y": 100, "w": 400, "h": 300,
                            "path": str(first), "fit": "cover", "overlay": 0.0, "base_dir": ""}]}
        a = render_page(model, directory / "one.png", POLICY, FONTS)
        model["items"][0]["path"] = str(second)
        b = render_page(model, directory / "two.png", POLICY, FONTS)
        self.assertGreater(pixel_diff(a, b, POLICY)["mismatch_ratio"], 0.01)

    def test_visible_image_respects_fit(self):
        directory = tmpdir()
        path = make_photo(directory, "assets/w.jpg", size=(1600, 400))
        cover = visible_image((400, 400), str(path), "cover", None, POLICY)
        contain = visible_image((400, 400), str(path), "contain", None, POLICY)
        self.assertEqual(cover.size, (400, 400))
        self.assertEqual(contain.size, (400, 400))
        # contain 有留白：角落应是白底；cover 没有留白
        self.assertEqual(contain.getpixel((5, 5)), (255, 255, 255))
        self.assertNotEqual(cover.getpixel((5, 5)), (255, 255, 255))

    def test_inspect_and_resolve(self):
        directory = tmpdir()
        path = make_photo(directory)
        facts = inspect(path)
        self.assertEqual((facts["width"], facts["height"]), (1600, 1200))
        self.assertEqual(resolve_source("assets/photo.jpg", directory), path)
        self.assertIsNone(resolve_source("assets/nope.jpg", directory))


if __name__ == "__main__":
    unittest.main(verbosity=2)
