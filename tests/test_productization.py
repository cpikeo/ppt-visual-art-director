#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
import unittest

from support import ROOT


class TestVersionAndChangelog(unittest.TestCase):
    def test_version_is_semver_and_documented(self):
        version = (ROOT / "VERSION").read_text().strip()
        self.assertRegex(version, r"^\d+\.\d+\.\d+$")
        changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
        self.assertIn(f"## {version}", changelog)
        for migration in ("--allow-author-code", "generated: yes", "manifest.schema.json",
                          "旧 evidence", "Python 最低支持版本"):
            self.assertIn(migration, changelog)


class TestCiContract(unittest.TestCase):
    def test_ci_covers_minimum_and_current_python_engineering_checks(self):
        text = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
        self.assertIn('python-version: ["3.10", "3.13"]', text)
        for command in ("unittest discover", "judgment_eval.py --ablate", "typography_eval.py",
                        "prompt_eval.py", "package_skill.py", "Draft202012Validator"):
            self.assertIn(command, text)

    def test_ci_does_not_claim_powerpoint_or_external_gates(self):
        text = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
        self.assertNotIn("READY TO RELEASE", text)
        self.assertIn("require external evidence", text)
        self.assertNotRegex(text, r"G-TGT.{0,40}PASS")


if __name__ == "__main__":
    unittest.main(verbosity=2)
