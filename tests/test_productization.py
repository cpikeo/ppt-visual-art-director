#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import sys
import tempfile
import unittest
from pathlib import Path

from support import ROOT

sys.path.insert(0, str(ROOT / "scripts"))
import release_check


class TestVersionAndChangelog(unittest.TestCase):
    def test_version_is_semver_and_documented(self):
        version = (ROOT / "VERSION").read_text().strip()
        self.assertRegex(version, r"^\d+\.\d+\.\d+$")
        changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
        self.assertIn(f"## {version}", changelog)
        for migration in ("--allow-author-code", "generated: yes", "manifest.schema.json",
                          "旧 evidence", "Python 最低支持版本"):
            self.assertIn(migration, changelog)


class TestEvidenceMatchesCurrentVersion(unittest.TestCase):
    """报告必须属于当前代码与策略：旧指纹的评测结果不能冒充当前证据。"""

    REPORTS = {"judgment-baseline.json": "judgment_eval.py",
               "prompt-baseline.json": "prompt_eval.py",
               "typography-baseline.json": "typography_eval.py",
               "stress-report.json": "stress_eval.py"}

    def test_every_stored_report_is_bound_to_the_current_policy(self):
        import sys
        sys.path.insert(0, str(ROOT / "scripts"))
        from policy import load as load_policy
        expected = load_policy().sha256
        stale = []
        for name, generator in self.REPORTS.items():
            path = ROOT / "evals" / "reports" / name
            self.assertTrue(path.is_file(), f"缺少评测报告：{name}")
            recorded = __import__("json").loads(path.read_text(encoding="utf-8")).get("policy_sha256")
            if recorded != expected:
                stale.append(f"{name}: {recorded} != {expected}（重新运行 evals/{generator}）")
        self.assertEqual(stale, [], "评测报告与当前策略不一致：旧证据不能当新证据用")

    def test_judgment_report_covers_every_case(self):
        import json
        import yaml
        cases = yaml.safe_load((ROOT / "evals" / "cases.yaml").read_text(encoding="utf-8"))["cases"]
        report = json.loads((ROOT / "evals" / "reports" / "judgment-baseline.json").read_text(encoding="utf-8"))
        self.assertEqual(report["candidate"]["total"], len(cases),
                         "报告里的用例数与 cases.yaml 不一致：报告过期了")
        self.assertEqual(report["candidate"]["passed"], len(cases))


class TestReleasePreflightRepeatable(unittest.TestCase):
    """预检本身也要可重复：同一目录重跑不得因上一轮残留而失败，证据不得串轮。"""

    def test_reset_removes_only_own_artifacts(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            (out / "package").mkdir()
            (out / "package" / "SKILL.md").write_text("stale", encoding="utf-8")
            for name in ("pvad-skill.zip", "manifest.json", "tests.stdout.txt", "tests.stderr.txt"):
                (out / name).write_text("stale", encoding="utf-8")
            keep = out / "notes.md"
            keep.write_text("keep", encoding="utf-8")
            removed = release_check.reset_outputs(out)
            self.assertFalse((out / "package").exists())
            self.assertFalse((out / "manifest.json").exists())
            self.assertFalse((out / "tests.stdout.txt").exists())
            self.assertTrue(keep.is_file(), "清场不得删除使用者自己放进来的文件")
            self.assertIn("package/", removed)

    def test_git_state_is_bound_or_honestly_unknown(self):
        state = release_check.git_state()
        self.assertEqual(set(state), {"commit", "dirty"})
        if state["commit"] != "UNKNOWN":
            self.assertRegex(state["commit"], r"^[0-9a-f]{40}$")
            self.assertIsInstance(state["dirty"], bool)

    @staticmethod
    def _stub(name, command, out):
        """替身阶段：不跑真实评测，只验证预检自身的编排与证据绑定。"""
        if name == "package":
            (out / "pvad-skill.zip").write_bytes(b"stub-zip")
        (out / f"{name}.stdout.txt").write_text("stub", encoding="utf-8")
        (out / f"{name}.stderr.txt").write_text("", encoding="utf-8")
        return {"name": name, "command": command, "exit_code": 0, "status": "PASS"}

    def _run_stubbed(self, out):
        import contextlib
        import io
        import unittest.mock
        buffer = io.StringIO()
        with unittest.mock.patch.object(release_check, "run", self._stub), contextlib.redirect_stdout(buffer):
            code = release_check.main(["--out", str(out)])
        return code, json.loads((out / "manifest.json").read_text(encoding="utf-8"))

    def test_manifest_binds_git_and_zip(self):
        import hashlib
        with tempfile.TemporaryDirectory() as tmp:
            code, manifest = self._run_stubbed(Path(tmp))
            self.assertEqual(code, 0, manifest.get("status"))
            self.assertEqual(manifest["status"], "PASS")
            self.assertEqual(set(manifest["git"]), {"commit", "dirty"})
            digest = hashlib.sha256((Path(tmp) / "pvad-skill.zip").read_bytes()).hexdigest()
            self.assertEqual(manifest["package_zip_sha256"], digest,
                             "manifest 记录的包指纹与本次产物不一致")

    def test_rerunning_into_the_same_directory_resets_instead_of_failing(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            (out / "package").mkdir()
            (out / "package" / "leftover.md").write_text("上一轮残留", encoding="utf-8")
            (out / "tests.stdout.txt").write_text("上一轮日志", encoding="utf-8")
            code, manifest = self._run_stubbed(out)
            self.assertEqual(code, 0, "同一目录重跑必须成功，而不是被上一轮残留卡住")
            self.assertIn("package/", manifest["reset"])
            self.assertIn("tests.stdout.txt", manifest["reset"])
            self.assertFalse((out / "package" / "leftover.md").exists())

    def test_package_stage_without_zip_fails_the_run(self):
        import contextlib
        import io
        import unittest.mock

        def no_zip(name, command, out):
            (out / f"{name}.stdout.txt").write_text("stub", encoding="utf-8")
            return {"name": name, "command": command, "exit_code": 0, "status": "PASS"}

        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            with unittest.mock.patch.object(release_check, "run", no_zip), contextlib.redirect_stdout(io.StringIO()):
                code = release_check.main(["--out", str(out)])
            manifest = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(code, 1, "没有产物却返回成功，等于用退出码冒充证据")
            self.assertEqual(manifest["status"], "FAIL")


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
