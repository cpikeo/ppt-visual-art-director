#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提示词层：判断一旦说「这页要图」，就得出得了一份能交给人去拍 / 生成的说明。

这一组测试盯的是另一种「假通过」：QC 硬门齐全，但没人说得清那张图该拍什么——
作者只写了一个路径，评审通过，生产端拿到的是空白。提示词层的存在条件因此有三条：
  1. 说明必须来自判断结论（不是从标题里猜关键词重新拼）；
  2. 说明必须可执行（拍什么 / 为什么 / 给谁让位 / 交付多大），且**不含坐标**；
  3. 说明必须可回归（同输入同输出，指纹跟着内容变）。
"""
import copy
import json
import unittest

from support import ROOT, tmpdir, write_brief

from PIL import Image

from policy import PolicyError, load as load_policy, validate_semantics
from brief import load as load_brief
from judgment import judge_deck
from compose import compose_deck
from prompt import (SCHEMA, build, fingerprint, manifest, quiet_side, ratio_of, target_px)
from typography import FontBook

SCHEMA_DIR = ROOT / "schema"
POLICY = load_policy()
FONTS = FontBook(POLICY)


def make_photo(directory, name="assets/photo.jpg", size=(1600, 1200)):
    path = directory / name
    path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", size, (120, 132, 140))
    for x in range(0, size[0], 8):
        for y in range(0, size[1], 8):
            image.putpixel((x, y), (90, 112, 130))
    image.save(path, "JPEG", quality=92)
    return path


def make_brief(directory, *, subject="华东区前置仓的分拣线，两名拣货员在通道间补货",
               function="hero", text_color="dark", area="[0.0, 0.08, 0.46, 0.84]",
               avoid="['brand logos', 'identifiable faces']", photo=True,
               path="assets/photo.jpg"):
    if photo:
        make_photo(directory)
    text = (
        "audience: 管理层\ndecision: 批准预算\ntension: 没有共识\n"
        "source: 内部看板\nunit: 万元\nperiod: 2026 Q1\nbasis: 已开票收入\n"
        "slides:\n  - id: s01\n    title: 加店不再带来增长\n"
        "    content: |-\n      门店数量增长 18%，单店月均产出下降 21%。\n"
        "    asset:\n"
        f"      path: {path}\n"
        f"      function: {function}\n"
        f"      subject: {subject}\n"
        "      reason: 让观众看见现场密度，而不是用一个百分比代替现场\n"
        "      fit: cover\n"
        f"      text_color: {text_color}\n"
        f"      safe_area: {area}\n"
        f"      avoid: {avoid}\n"
    )
    brief, _ = load_brief(write_brief(directory, text=text), POLICY, SCHEMA_DIR)
    return brief


def build_deck(brief, directory):
    plan = judge_deck(brief, POLICY, base_dir=directory)
    return plan, compose_deck(brief, plan, POLICY, fontbook=FONTS, base_dir=directory)


def media_element(deck):
    return next(e for e in deck["slides"][0]["elements"] if e["type"] == "image")


def simple(**kwargs):
    """直接调用拼装器：不经过 brief/plan，测拼装本身。"""
    params = dict(policy=POLICY, function="hero", role="subject",
                  subject="a warehouse picking line", reason="现场密度",
                  box=(0, 0, 496, 624), safe_area={"x": 0.0, "y": 0.08, "width": 0.46, "height": 0.84},
                  text_color="dark")
    params.update(kwargs)
    return build(**params)


class TestBuild(unittest.TestCase):
    def test_prompt_states_four_executable_facts(self):
        brief = simple()
        # 拍什么 / 为什么拍 / 给谁让位 / 交付多大
        self.assertIn("a warehouse picking line", brief["prompt"])
        self.assertIn("main visual statement", brief["prompt"])
        self.assertIn("left side of the frame stays quiet", brief["prompt"])
        self.assertEqual(brief["ratio"], "4:5")
        self.assertGreaterEqual(min(brief["target_px"].values()), 640)
        self.assertEqual(brief["issues"], [])

    def test_prompt_carries_no_coordinates(self):
        """坐标进提示词会把留白画成硬边面板（上游实测）；矩形只属于 QC。"""
        prompt = simple()["prompt"]
        self.assertNotIn("%", prompt)
        self.assertNotIn("px", prompt)
        self.assertNotIn("0.46", prompt)

    def test_quiet_side_follows_the_declared_safe_area(self):
        for area, side in (
            ({"x": 0.0, "y": 0.1, "width": 0.4, "height": 0.8}, "left"),
            ({"x": 0.6, "y": 0.1, "width": 0.4, "height": 0.8}, "right"),
            ({"x": 0.1, "y": 0.0, "width": 0.8, "height": 0.25}, "top"),
            ({"x": 0.1, "y": 0.7, "width": 0.8, "height": 0.25}, "bottom"),
            ({"x": 0.3, "y": 0.3, "width": 0.4, "height": 0.4}, "center"),
        ):
            self.assertEqual(quiet_side(area), side)
            brief = simple(safe_area=area)
            self.assertEqual(brief["quiet_side"], side)
            # 提示词用的是策略里那一句（top 说 "upper part"，不必含 "top"）
            self.assertIn(POLICY.prompts["phrases"]["quiet_area"][side], brief["prompt"])

    def test_no_safe_area_no_quiet_talk(self):
        """版面不压文字就不发「让位」指令：空话会稀释真正的约束。"""
        brief = simple(safe_area=None)
        self.assertIsNone(brief["quiet_side"])
        self.assertNotIn("stays quiet", brief["prompt"])

    def test_tone_only_when_text_color_is_declared(self):
        self.assertIn("dark type", simple(text_color="dark")["prompt"])
        self.assertIn("light type", simple(text_color="light")["prompt"])
        brief = simple(text_color="")
        self.assertNotIn("type", brief["prompt"])
        self.assertIn("PROMPT_TONE_MISSING", [i["code"] for i in brief["issues"]])

    def test_role_changes_the_instructions(self):
        background = simple(role="background")["prompt"]
        subject = simple(role="subject")["prompt"]
        self.assertNotEqual(background, subject)
        self.assertIn("environment", background)
        self.assertIn("independent visual object", subject)

    def test_function_changes_the_duty_sentence(self):
        proof = simple(function="proof")["prompt"]
        hero = simple(function="hero")["prompt"]
        self.assertIn("readable for verification", proof)
        self.assertNotEqual(proof, hero)

    def test_proof_background_is_a_contradiction(self):
        brief = simple(function="proof", role="background")
        self.assertIn("PROMPT_ROLE_CONFLICT", [i["code"] for i in brief["issues"]])

    def test_missing_subject_is_an_issue_not_an_exception(self):
        brief = simple(subject="   ")
        self.assertIn("PROMPT_SUBJECT_MISSING", [i["code"] for i in brief["issues"]])
        self.assertTrue(brief["prompt"])          # 说明仍然产出，缺口如实登记

    def test_negative_excludes_production_failures(self):
        negative = simple()["negative"]
        for term in ("no AI artifacts", "no unintended visible seams", "no watermark"):
            self.assertIn(term, negative)
        # 反向词只排生产失效，不排风格
        for style in ("luxury", "cinematic", "premium"):
            self.assertNotIn(style, negative.lower())

    def test_author_avoid_is_preserved_and_normalised(self):
        negative = simple(avoid=["brand logos", "no faces"])["negative"]
        self.assertIn("no brand logos", negative)
        self.assertIn("no faces", negative)        # 作者已写 no 的不重复成 no no

    def test_target_px_never_below_policy_floor(self):
        tiny = target_px(40, 40, POLICY)
        floor = POLICY.req("assets.min_long_side_px")
        self.assertGreaterEqual(max(tiny.values()), floor)
        over = float(POLICY.req("assets.oversample"))
        big = target_px(496, 624, POLICY)
        self.assertGreaterEqual(big["width"], int(496 * over))

    def test_ratio_snaps_to_usable_framing(self):
        self.assertEqual(ratio_of(1280, 720, POLICY), "16:9")
        self.assertEqual(ratio_of(496, 624, POLICY), "4:5")     # 31:39 没人会照做
        self.assertEqual(ratio_of(1000, 1000, POLICY), "1:1")

    def test_fingerprint_tracks_content(self):
        a, b = simple(), simple()
        self.assertEqual(a["fingerprint"], b["fingerprint"])           # 确定性
        self.assertNotEqual(a["fingerprint"], simple(subject="another field")["fingerprint"])
        self.assertEqual(a["fingerprint"], fingerprint(a))

    def test_prompt_is_deterministic_across_runs(self):
        self.assertEqual(json.dumps(simple(), sort_keys=True),
                         json.dumps(simple(), sort_keys=True))

    def test_too_long_prompt_is_reported_not_truncated(self):
        brief = simple(subject="长句" * 400)
        self.assertIn("PROMPT_TOO_LONG", [i["code"] for i in brief["issues"]])
        self.assertIn("长句", brief["prompt"])     # 不静默截断

    def test_schema_tag_is_stable(self):
        self.assertEqual(simple()["schema"], SCHEMA)


class TestPolicyLayer(unittest.TestCase):
    """五层边界：语言在策略里，代码里不许再写一份。"""

    def test_prompts_are_loaded_and_fingerprinted(self):
        self.assertTrue(POLICY.prompts)
        self.assertTrue(any(p.endswith("prompts.yaml") for p in POLICY.paths))
        self.assertEqual(len(POLICY.sha256), 64)

    def test_phrase_tables_match_enums(self):
        prompts = POLICY.prompts["phrases"]
        self.assertEqual(set(prompts["function"]), set(POLICY.enum("media_function")))
        self.assertEqual(set(prompts["role"]), set(POLICY.enum("media_role")))

    def test_policy_bans_style_words_in_its_own_phrases(self):
        banned = [w.lower() for w in POLICY.prompts["style_words_banned"]]
        blob = json.dumps(POLICY.prompts["phrases"], ensure_ascii=False).lower()
        self.assertEqual([w for w in banned if w in blob], [])

    def test_removing_a_phrase_breaks_the_load(self):
        """策略与枚举不同步时加载即失败——不允许静默丢一句职责说明。"""
        broken = copy.deepcopy(POLICY)
        del broken.prompts["phrases"]["function"]["proof"]
        with self.assertRaises(PolicyError) as ctx:
            validate_semantics(broken)
        self.assertIn("proof", str(ctx.exception))

    def test_style_word_in_phrases_breaks_the_load(self):
        broken = copy.deepcopy(POLICY)
        broken.prompts["phrases"]["function"]["hero"] = "a cinematic hero image"
        with self.assertRaises(PolicyError) as ctx:
            validate_semantics(broken)
        self.assertIn("风格词", str(ctx.exception))


class TestWiring(unittest.TestCase):
    """说明必须由判断驱动，并且只在该生产的时候出现。"""

    def test_missing_file_produces_a_brief_in_the_deck(self):
        directory = tmpdir()
        brief = make_brief(directory, path="assets/not_yet_shot.jpg")
        plan, deck = build_deck(brief, directory)
        media = plan["pages"][0]["decision"]["media"]
        self.assertEqual(media["decision"], "required")
        self.assertEqual(media["subject"][:2], "华东")
        element = media_element(deck)
        self.assertTrue(element["_missing"])
        self.assertIn("prompt", element["brief"])
        self.assertIn("华东区前置仓", element["brief"]["prompt"])
        self.assertEqual(element["brief"]["fingerprint"][:7], "prompt-")

    def test_existing_file_needs_no_prompt(self):
        directory = tmpdir()
        brief = make_brief(directory)
        plan, deck = build_deck(brief, directory)
        self.assertEqual(plan["pages"][0]["decision"]["media"]["decision"], "existing")
        self.assertNotIn("brief", media_element(deck))

    def test_existing_file_found_relative_to_brief_not_cwd(self):
        """相对路径按 brief 目录解析：用 cwd 判断会把「文件其实在」误判成要出图。"""
        directory = tmpdir()
        brief = make_brief(directory)
        plan = judge_deck(brief, POLICY, base_dir=directory)
        self.assertEqual(plan["pages"][0]["decision"]["media"]["decision"], "existing")

    def test_manifest_lists_only_what_must_be_produced(self):
        directory = tmpdir()
        brief = make_brief(directory, path="assets/not_yet_shot.jpg")
        _, deck = build_deck(brief, directory)
        result = manifest(deck, POLICY)
        self.assertEqual(result["images"], 1)
        self.assertEqual(result["to_produce"], 1)
        entry = result["entries"][0]
        self.assertEqual(entry["status"], "to-produce")
        self.assertTrue(entry["prompt"])
        self.assertTrue(entry["negative"])
        self.assertEqual(entry["ratio"], "4:5")
        self.assertIn("assets/not_yet_shot.jpg", entry["expected_path"])

    def test_manifest_marks_ready_images_ready(self):
        directory = tmpdir()
        brief = make_brief(directory)
        _, deck = build_deck(brief, directory)
        result = manifest(deck, POLICY)
        self.assertEqual(result["to_produce"], 0)
        self.assertEqual(result["entries"][0]["status"], "ready")

    def test_deck_with_brief_passes_schema_and_leaves_no_undeclared_warning(self):
        from pvad import validate_deck
        from guard import check as guard_check
        directory = tmpdir()
        brief = make_brief(directory, path="assets/not_yet_shot.jpg")
        _, deck = build_deck(brief, directory)
        self.assertEqual(validate_deck(deck), [])
        result = guard_check(deck, POLICY, mode="release")
        codes = {e["code"] for e in result["errors"]}
        self.assertIn("ASSET_MISSING", codes)      # 缺图必须拦
        self.assertNotIn("FIELD_UNDECLARED", {w["code"] for w in result["warnings"]})


if __name__ == "__main__":
    unittest.main()
