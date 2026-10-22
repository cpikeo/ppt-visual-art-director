#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import hashlib
import json
import unittest
from pathlib import Path

from support import tmpdir  # 先导入 support：它把 scripts/ 放进 sys.path

from package_skill import PackageSkillError, build_deterministic_zip, build_package


class TestSkillPackage(unittest.TestCase):
    def test_runtime_package_is_allowlisted_and_smoke_checked(self):
        tmp = tmpdir()
        output = tmp / "runtime"
        result = build_package(output)
        self.assertEqual(result["smoke"], "PASS")
        self.assertFalse((output / "audit").exists())
        self.assertFalse((output / "tests").exists())
        self.assertFalse((output / "evals").exists())
        for required in ("SKILL.md", "VERSION", "CHANGELOG.md", "requirements.txt",
                         "scripts/pvad.py", "schema/manifest.schema.json"):
            self.assertTrue((output / required).is_file(), required)
        manifest = json.loads((output / "PACKAGE-MANIFEST.json").read_text())
        self.assertEqual(manifest["version"], (Path(__file__).resolve().parent.parent / "VERSION")
                         .read_text(encoding="utf-8").strip(),
                         "包内版本必须等于 VERSION：版本号只能有一个真值")
        for item in manifest["files"]:
            path = output / item["path"]
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), item["sha256"])

    def test_reviewable_package_can_include_tests_and_evals_without_generated_reports(self):
        output = tmpdir() / "reviewable"
        result = build_package(output, include_tests=True)
        self.assertEqual(result["smoke"], "PASS")
        self.assertTrue((output / "tests" / "test_pipeline.py").is_file())
        self.assertTrue((output / "evals" / "cases.yaml").is_file())
        self.assertFalse((output / "evals" / "reports").exists())

    def test_deterministic_zip_has_stable_hash(self):
        tmp = tmpdir()
        output = tmp / "runtime"
        build_package(output)
        first = build_deterministic_zip(output, tmp / "first.zip")
        second = build_deterministic_zip(output, tmp / "second.zip")
        self.assertEqual(first["zip_sha256"], second["zip_sha256"])
        self.assertEqual((tmp / "first.zip").read_bytes(), (tmp / "second.zip").read_bytes())

    def test_nonempty_output_is_rejected(self):
        output = tmpdir() / "existing"
        output.mkdir()
        (output / "keep.txt").write_text("keep")
        with self.assertRaisesRegex(PackageSkillError, "不存在或为空"):
            build_package(output)
        self.assertEqual((output / "keep.txt").read_text(), "keep")


if __name__ == "__main__":
    unittest.main(verbosity=2)
