#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端：plan → build → check 的退出码、清单证据与「不宣称未验证的事」。

清单（manifest）必须是可审计的凭证：
  * 先写 BLOCKED，全链完成才写 PASS；
  * target PowerPoint / 目标渲染器 / 独立盲评三行永远是「未验证」；
  * 输出 PPTX 的 sha256 与实际文件一致。
"""
import json
import unittest

import support
from support import ROOT, tmpdir, write_brief

from policy import load as load_policy
from compile import readback

EXIT_OK, EXIT_BLOCK, EXIT_USAGE = 0, 1, 2


class TestHappyPath(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        cls.plan, cls.deck, cls.manifest, cls.artifacts = support.pipeline(cls.tmp, mode="release")

    def test_release_exit_is_blocked_until_external_mandatory_gates_are_verified(self):
        self.assertEqual(self.artifacts["exit"], EXIT_BLOCK, self.artifacts["stderr"])
        self.assertEqual(self.manifest["delivery_status"], "BLOCK — Delivery")

    def test_plan_records_judgment_and_policy(self):
        self.assertEqual(self.plan["schema"], "pvad-plan-v1")
        self.assertEqual(self.plan["policy_sha256"], load_policy().sha256)
        for page in self.plan["pages"]:
            self.assertIn(page["claim"]["source"], ("author", "derived", "absent"))
            self.assertIn("why", page["understanding"]["evidence"])
            self.assertIn("composition_why", page["decision"])

    def test_plan_matches_schema(self):
        import sys
        sys.path.insert(0, str(ROOT / "scripts"))
        from pvad import validate_against
        self.assertEqual(validate_against(self.plan, "plan.schema.json"), [])

    def test_deck_is_authorable(self):
        self.assertEqual(self.deck["schema"], "pvad-deck-v1")
        from pvad import validate_against
        self.assertEqual(validate_against(self.deck, "deck.schema.json"), [])
        self.assertTrue(all(e.get("layout_source") in ("plan", "author")
                            for s in self.deck["slides"] for e in s["elements"]))

    def test_manifest_claims_only_what_was_verified(self):
        self.assertEqual(self.manifest["status"], "PASS")
        verification = self.manifest["verification"]
        self.assertEqual(verification["structural"], "PASS")
        self.assertEqual(verification["pptx_readback"], "MATCH")
        self.assertTrue(verification["target_powerpoint"].startswith("NOT VERIFIED"))
        self.assertTrue(verification["target_renderer"].startswith("NOT VERIFIED"))
        self.assertTrue(verification["independent_visual_review"].startswith("NOT RUN"))
        self.assertEqual(self.manifest["quality_gates"]["G-TGT"]["status"], "NOT VERIFIED")
        self.assertEqual(self.manifest["quality_gates"]["G-COMPLY"]["status"], "NOT RUN")
        self.assertEqual(self.manifest["delivery_status"], "BLOCK — Delivery")

    def test_output_hash_matches_file(self):
        import hashlib
        digest = hashlib.sha256(self.artifacts["pptx"].read_bytes()).hexdigest()
        self.assertEqual(digest, self.manifest["output_sha256"])

    def test_previews_exist(self):
        report = self.artifacts["report"]
        self.assertTrue((report / "preview-spec.png").is_file())
        self.assertTrue((report / "preview-pptx.png").is_file())
        self.assertTrue((report / "pages").is_dir())

    def test_compiled_deck_opens(self):
        rb = readback(self.artifacts["pptx"])
        self.assertEqual(len(rb["slides"]), len(self.deck["slides"]))


class TestBlocking(unittest.TestCase):
    def test_build_without_provenance_is_stopped_before_layout(self):
        tmp = tmpdir()
        text = support.MINI_BRIEF.replace("period: 2026 Q1–Q3\n", "").replace("basis: 已开票收入\n", "")
        brief = write_brief(tmp, text=text)
        plan, deck = tmp / "plan.json", tmp / "deck.json"
        self.assertEqual(support.run_cli("plan", brief, "--out", plan).returncode, EXIT_OK)
        result = support.run_cli("build", plan, "--brief", brief, "--out", deck)
        self.assertEqual(result.returncode, EXIT_BLOCK)
        self.assertIn("BLOCK — BUILD", result.stdout)
        self.assertIn("PROVENANCE_MISSING", result.stdout)
        self.assertFalse(deck.exists(), "语义闸门失败后不应先生成 deck 再补理由")

    def test_strict_plan_blocks_on_absent_claim(self):
        tmp = tmpdir()
        brief = write_brief(tmp, text="audience: 评审\ndecision: 批准\nslides:\n"
                                      "  - id: s01\n    title: 渠道结构\n    content: |-\n"
                                      "      经销 51%\n      直营 34%\n      电商 15%\n")
        result = support.run_cli("plan", brief, "--out", tmp / "plan.json", "--strict")
        self.assertEqual(result.returncode, EXIT_BLOCK)
        self.assertIn("主张未决", result.stdout)

    def test_build_blocks_unresolved_claim_even_without_plan_strict_flag(self):
        tmp = tmpdir()
        brief = write_brief(tmp, text="audience: 评审\ndecision: 批准\ntension: 渠道选择\n"
                                      "source: 看板\nunit: '%'\nperiod: '2026'\nbasis: 收入口径\nslides:\n"
                                      "  - id: s01\n    title: 渠道结构\n    content: |-\n"
                                      "      经销 51%\n      直营 34%\n      电商 15%\n")
        plan = tmp / "plan.json"
        self.assertEqual(support.run_cli("plan", brief, "--out", plan).returncode, EXIT_OK)
        result = support.run_cli("build", plan, "--brief", brief, "--out", tmp / "deck.json")
        self.assertEqual(result.returncode, EXIT_BLOCK)
        self.assertIn("CLAIM_ABSENT", result.stdout)

    def test_policy_drift_blocks_build(self):
        tmp = tmpdir()
        brief = write_brief(tmp)
        self.assertEqual(support.run_cli("plan", brief, "--out", tmp / "plan.json").returncode, EXIT_OK)
        plan = support.load_json(tmp / "plan.json")
        plan["policy_sha256"] = "0" * 64
        (tmp / "plan.json").write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        result = support.run_cli("build", tmp / "plan.json", "--brief", brief, "--out", tmp / "deck.json")
        self.assertEqual(result.returncode, EXIT_USAGE)
        self.assertIn("指纹", result.stderr)

    def test_brief_changed_after_plan_blocks_build(self):
        tmp = tmpdir()
        brief = write_brief(tmp)
        self.assertEqual(support.run_cli("plan", brief, "--out", tmp / "plan.json").returncode, EXIT_OK)
        write_brief(tmp, text=support.MINI_BRIEF.replace("加店不再带来增长", "换个标题"))
        result = support.run_cli("build", tmp / "plan.json", "--brief", brief, "--out", tmp / "deck.json")
        self.assertEqual(result.returncode, EXIT_USAGE)
        self.assertIn("被改过", result.stderr)

    def test_bad_brief_is_usage_error(self):
        tmp = tmpdir()
        brief = write_brief(tmp, text="slides: []\n")
        result = support.run_cli("plan", brief, "--out", tmp / "plan.json")
        self.assertEqual(result.returncode, EXIT_USAGE)

    def test_invalid_deck_schema_blocks_without_crashing(self):
        tmp = tmpdir()
        deck = {"schema": "pvad-deck-v1", "canvas": {"width": 1280, "height": 720},
                "theme": {"colors": {"ink": "#000000"}, "fonts": {"cn": "a", "latin": "b"}},
                "slides": [{"id": "s01", "elements": [{"id": "x"}]}]}
        path = tmp / "deck.json"
        path.write_text(json.dumps(deck), encoding="utf-8")
        result = support.run_cli("check", path, tmp / "out.pptx", "--out-dir", tmp / "report")
        self.assertEqual(result.returncode, EXIT_BLOCK)
        manifest = support.load_json(tmp / "report" / "manifest.json")
        self.assertEqual(manifest["status"], "BLOCKED")
        self.assertEqual({e["code"] for e in manifest["guard"]["errors"]}, {"SCHEMA_FAIL"})


class TestDeterminism(unittest.TestCase):
    def test_same_inputs_same_plan_and_deck(self):
        tmp = tmpdir()
        brief = write_brief(tmp)
        digests = []
        for i in range(2):
            plan, deck = tmp / f"plan{i}.json", tmp / f"deck{i}.json"
            support.run_cli("plan", brief, "--out", plan)
            support.run_cli("build", plan, "--brief", brief, "--out", deck)
            digests.append((support.load_json(plan)["pages"], support.load_json(deck)["slides"]))
        self.assertEqual(digests[0], digests[1])

    def test_compiled_bytes_are_reproducible_across_runs(self):
        tmp = tmpdir()
        brief = write_brief(tmp)
        hashes = []
        for i in range(2):
            plan, deck, pptx = tmp / f"p{i}.json", tmp / f"d{i}.json", tmp / f"o{i}.pptx"
            support.run_cli("plan", brief, "--out", plan)
            support.run_cli("build", plan, "--brief", brief, "--out", deck)
            support.run_cli("check", deck, pptx, "--mode", "draft", "--out-dir", tmp / f"r{i}")
            hashes.append(support.load_json(tmp / f"r{i}" / "manifest.json")["output_sha256"])
        self.assertEqual(hashes[0], hashes[1])


if __name__ == "__main__":
    unittest.main(verbosity=2)
