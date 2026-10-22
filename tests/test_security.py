#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import copy
import json
import os
import unittest

import support
from support import tmpdir


class TestAuthorCodeBoundary(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        _, cls.deck, _, _ = support.pipeline(cls.tmp, mode="draft")
        cls.author = cls.tmp / "author_deck.py"
        cls.author.write_text(
            "DECK = " + repr(cls.deck) + "\n\ndef build_deck(plan, brief):\n    return DECK\n",
            encoding="utf-8",
        )

    def test_python_author_deck_is_rejected_by_default(self):
        result = support.run_cli("check", self.author, self.tmp / "blocked.pptx",
                                 "--mode", "spec", "--out-dir", self.tmp / "blocked-report")
        self.assertEqual(result.returncode, 2)
        self.assertIn("--allow-author-code", result.stderr)
        self.assertFalse((self.tmp / "blocked-report" / "manifest.json").exists())

    def test_release_mode_rejects_python_even_with_explicit_permission(self):
        result = support.run_cli("check", self.author, self.tmp / "release.pptx", "--mode", "release",
                                 "--out-dir", self.tmp / "release-report", "--allow-author-code")
        self.assertEqual(result.returncode, 2)
        self.assertIn("release 模式禁止", result.stderr)

    def test_explicit_author_code_permission_is_warned_and_recorded(self):
        report = self.tmp / "allowed-report"
        result = support.run_cli("check", self.author, self.tmp / "allowed.pptx",
                                 "--mode", "spec", "--out-dir", report, "--allow-author-code")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("SECURITY WARNING", result.stderr)
        manifest = support.load_json(report / "manifest.json")
        self.assertTrue(manifest["security"]["author_code_executed"])
        self.assertFalse(manifest["security"]["sandboxed"])
        self.assertEqual(manifest["security"]["author_code_path"], str(self.author.resolve()))


class TestPackageBoundary(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        _, cls.deck, cls.manifest, cls.artifacts = support.pipeline(cls.tmp, mode="release")
        cls.manifest_path = cls.artifacts["report"] / "manifest.json"

    def _write_manifest(self, value, name):
        path = self.tmp / name
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        return path

    def test_manifest_with_unknown_field_is_rejected(self):
        manifest = copy.deepcopy(self.manifest)
        manifest["forged"] = True
        path = self._write_manifest(manifest, "forged.json")
        result = support.run_cli("package", path, "--stage", "draft", "--out", self.tmp / "forged-pkg")
        self.assertEqual(result.returncode, 2)
        self.assertIn("manifest 不符合 schema", result.stderr)

    def test_ready_manifest_with_nonpassing_gate_is_rejected(self):
        manifest = copy.deepcopy(self.manifest)
        manifest["delivery_status"] = "READY TO RELEASE"
        manifest["quality_gates"]["G-TGT"]["status"] = "NOT VERIFIED"
        path = self._write_manifest(manifest, "false-ready.json")
        result = support.run_cli("package", path, "--stage", "release", "--out", self.tmp / "false-ready-pkg")
        self.assertEqual(result.returncode, 2)
        self.assertIn("manifest 不符合 schema", result.stderr)

    def test_stale_pptx_hash_is_rejected(self):
        copied = self.tmp / "tampered.pptx"
        copied.write_bytes(self.artifacts["pptx"].read_bytes() + b"tamper")
        manifest = copy.deepcopy(self.manifest)
        manifest["output_pptx"] = str(copied)
        path = self._write_manifest(manifest, "tampered-manifest.json")
        result = support.run_cli("package", path, "--stage", "draft", "--out", self.tmp / "tampered-pkg")
        self.assertEqual(result.returncode, 2)
        self.assertIn("SHA-256", result.stderr)

    def test_symlink_preview_is_rejected(self):
        real = self.tmp / "real-preview.png"
        real.write_bytes(b"png")
        link = self.tmp / "preview-link.png"
        try:
            os.symlink(real, link)
        except (OSError, NotImplementedError):
            self.skipTest("symlink unavailable")
        manifest = copy.deepcopy(self.manifest)
        manifest["previews"]["spec_sheet"] = str(link)
        path = self._write_manifest(manifest, "symlink-manifest.json")
        result = support.run_cli("package", path, "--stage", "review", "--out", self.tmp / "symlink-pkg")
        self.assertEqual(result.returncode, 2)
        self.assertIn("符号链接", result.stderr)

    def test_nonempty_output_directory_is_rejected(self):
        output = self.tmp / "nonempty"
        output.mkdir()
        (output / "existing.txt").write_text("do not overwrite", encoding="utf-8")
        result = support.run_cli("package", self.manifest_path, "--stage", "draft", "--out", output)
        self.assertEqual(result.returncode, 2)
        self.assertIn("拒绝覆盖", result.stderr)
        self.assertEqual((output / "existing.txt").read_text(), "do not overwrite")


if __name__ == "__main__":
    unittest.main(verbosity=2)
