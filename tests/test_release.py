#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import unittest

import support
from support import ROOT, tmpdir

from release import ReleaseEvidenceError, evaluate_evidence, load_evidence


SHA = "a" * 64


def complete_evidence(sha=SHA):
    return {
        "schema": "pvad-release-evidence-v1",
        "artifact_sha256": sha,
        "deck_review": {
            "status": "PASS", "reviewer": "Alice", "reviewed_at": "2026-10-22T12:00:00+08:00",
            "checklist": {
                "first_attention": True, "claim_evidence_focus": True, "adjacent_increment": True,
                "critical_page_weight": True, "climax_position": True, "visual_family": True,
                "whole_deck_rhythm": True,
            },
            "must_fix": [], "defects": [],
        },
        "target_environment": {
            "status": "PASS", "reviewer": "Bob", "reviewed_at": "2026-10-22T13:00:00+08:00",
            "application": "Microsoft PowerPoint", "application_version": "Version 2509",
            "operating_system": "Windows 11 24H2", "fonts": ["Noto Sans CJK SC", "DejaVu Sans"],
            "actions": {key: True for key in ("opened", "slide_show_checked", "edited", "saved_as",
                                                      "reopened", "charts_checked", "line_breaks_checked", "crops_checked")},
            "issues": [],
        },
        "compliance": {
            "status": "PASS", "reviewer": "Carol", "reviewed_at": "2026-10-22T14:00:00+08:00",
            "no_external_assets": True, "assets": [], "sensitive_content_reviewed": True, "issues": [],
        },
    }


class TestReleaseEvidence(unittest.TestCase):
    def test_complete_evidence_passes_three_external_gates(self):
        result = evaluate_evidence(complete_evidence(), SHA, {"slides": []})
        self.assertEqual({key: value["status"] for key, value in result.items()},
                         {"G-DECK": "PASS", "G-TGT": "PASS", "G-COMPLY": "PASS"})

    def test_artifact_hash_mismatch_is_rejected(self):
        with self.assertRaisesRegex(ReleaseEvidenceError, "不得跨版本"):
            evaluate_evidence(complete_evidence(), "b" * 64, {"slides": []})

    def test_deck_review_cannot_pass_with_must_fix(self):
        evidence = complete_evidence()
        evidence["deck_review"]["must_fix"] = ["修复第 2 页焦点"]
        result = evaluate_evidence(evidence, SHA, {"slides": []})
        self.assertEqual(result["G-DECK"]["status"], "FAIL")

    def test_target_cannot_pass_without_all_actions(self):
        evidence = complete_evidence()
        evidence["target_environment"]["actions"]["reopened"] = False
        result = evaluate_evidence(evidence, SHA, {"slides": []})
        self.assertEqual(result["G-TGT"]["status"], "FAIL")

    def test_no_external_assets_claim_fails_when_deck_has_image(self):
        evidence = complete_evidence()
        deck = {"slides": [{"id": "s1", "elements": [{"type": "image", "source": "asset.jpg"}]}]}
        result = evaluate_evidence(evidence, SHA, deck)
        self.assertEqual(result["G-COMPLY"]["status"], "FAIL")

    def test_asset_ledger_checks_intended_use_without_forcing_unused_rights(self):
        evidence = complete_evidence()
        evidence["compliance"].update({
            "no_external_assets": False,
            "assets": [{
                "slide_id": "s1", "source": "asset.jpg", "author_or_institution": "Owner",
                "acquisition": "licensed source", "license": "project-specific permission",
                "commercial_use": False, "modification_allowed": False,
                "intended_use_cleared": True, "attribution": "", "privacy_cleared": True,
                "generated": False, "professional_review": "not_required",
            }],
        })
        deck = {"slides": [{"id": "s1", "elements": [{"type": "image", "source": "asset.jpg"}]}]}
        result = evaluate_evidence(evidence, SHA, deck)
        self.assertEqual(result["G-COMPLY"]["status"], "PASS")

    def test_generated_asset_requires_positive_professional_review(self):
        evidence = complete_evidence()
        evidence["compliance"].update({
            "no_external_assets": False,
            "assets": [{
                "slide_id": "s1", "source": "generated.png", "author_or_institution": "AI tool",
                "acquisition": "generated", "license": "tool terms", "commercial_use": True,
                "modification_allowed": True, "intended_use_cleared": True, "attribution": "",
                "privacy_cleared": True, "generated": True, "professional_review": "not_required",
            }],
        })
        deck = {"slides": [{"id": "s1", "elements": [{"type": "image", "source": "generated.png"}]}]}
        result = evaluate_evidence(evidence, SHA, deck)
        self.assertEqual(result["G-COMPLY"]["status"], "FAIL")

    def test_schema_rejects_incomplete_target_record(self):
        tmp = tmpdir()
        evidence = complete_evidence()
        del evidence["target_environment"]["application_version"]
        path = tmp / "evidence.json"
        path.write_text(json.dumps(evidence), encoding="utf-8")
        with self.assertRaisesRegex(ReleaseEvidenceError, "schema"):
            load_evidence(path, ROOT / "schema")

    def test_release_cli_accepts_only_hash_bound_complete_external_evidence(self):
        tmp = tmpdir()
        plan, deck, manifest, artifacts = support.pipeline(tmp, mode="release")
        evidence = complete_evidence(manifest["output_sha256"])
        path = tmp / "release-evidence.json"
        path.write_text(json.dumps(evidence, ensure_ascii=False), encoding="utf-8")
        result = support.run_cli("check", tmp / "deck.json", artifacts["pptx"], "--mode", "release",
                                 "--out-dir", artifacts["report"], "--evidence", path)
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        final = support.load_json(artifacts["report"] / "manifest.json")
        self.assertEqual(final["delivery_status"], "READY TO RELEASE")
        self.assertTrue(all(gate["status"] == "PASS" for gate in final["quality_gates"].values()))
        release_dir = tmp / "release-package"
        packaged = support.run_cli("package", artifacts["report"] / "manifest.json",
                                   "--stage", "release", "--out", release_dir)
        self.assertEqual(packaged.returncode, 0, packaged.stderr)
        self.assertTrue((release_dir / "deck.pptx").is_file())
        self.assertTrue((release_dir / "SHA256SUMS.txt").is_file())


class TestDeliveryExperience(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        cls.plan, cls.deck, cls.manifest, cls.artifacts = support.pipeline(cls.tmp, mode="release")

    def test_check_writes_summary_consistent_with_manifest(self):
        summary = self.artifacts["report"] / "delivery-summary.md"
        self.assertTrue(summary.is_file())
        text = summary.read_text(encoding="utf-8")
        self.assertIn(self.manifest["status"], text)
        self.assertIn(self.manifest["delivery_status"], text)
        self.assertIn(self.manifest["output_sha256"], text)
        self.assertIn("G-TGT", text)
        self.assertIn("由 manifest 确定性生成", text)

    def test_draft_and_review_packages_are_distinct(self):
        manifest = self.artifacts["report"] / "manifest.json"
        draft = self.tmp / "draft-package"
        review = self.tmp / "review-package"
        self.assertEqual(support.run_cli("package", manifest, "--stage", "draft", "--out", draft).returncode, 0)
        self.assertEqual(support.run_cli("package", manifest, "--stage", "review", "--out", review).returncode, 0)
        self.assertFalse((draft / "REVIEW_REQUIRED.md").exists())
        self.assertTrue((review / "REVIEW_REQUIRED.md").is_file())
        self.assertTrue((review / "preview-pptx.png").is_file())

    def test_release_package_is_rejected_while_delivery_blocked(self):
        result = support.run_cli("package", self.artifacts["report"] / "manifest.json",
                                 "--stage", "release", "--out", self.tmp / "blocked-release")
        self.assertEqual(result.returncode, 2)
        self.assertIn("READY TO RELEASE", result.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
