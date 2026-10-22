#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Active protocol references must use stable semantic anchors, not section numbers."""
import re
import unittest
from pathlib import Path

import support
from support import ROOT, tmpdir, write_brief


class TestProtocolReferences(unittest.TestCase):
    def test_skill_declares_unique_stable_anchors(self):
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        anchors = re.findall(r'<a id="([a-z0-9-]+)"></a>', text)
        self.assertEqual(len(anchors), len(set(anchors)))
        required = {"workflow", "task-routing", "think-handoff", "authority-boundaries",
                    "design-judgment", "chinese-typography", "asset-direction",
                    "evidence-levels", "quality-gates"}
        self.assertTrue(required <= set(anchors))

    def test_active_files_do_not_reference_bare_skill_section_numbers(self):
        offenders = []
        for root in ("scripts", "evals", "tests", "references", "policy"):
            for path in (ROOT / root).rglob("*"):
                if not path.is_file() or "reports" in path.parts or path.suffix == ".pyc":
                    continue
                try:
                    text = path.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    continue
                if re.search(r"SKILL(?:\.md)?\s*§\s*\d", text):
                    offenders.append(str(path.relative_to(ROOT)))
        self.assertEqual(offenders, [])

    def test_first_semantic_reference_name_exists_as_skill_heading(self):
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        raw_headings = re.findall(r"^#{1,6}\s+([^｜\n]+)", skill, re.M)
        headings = {re.sub(r"^\d+(?:\.\d+)*\.?\s+", "", heading).strip() for heading in raw_headings}
        names = set()
        for root in ("scripts", "evals", "tests", "policy"):
            for path in (ROOT / root).rglob("*"):
                if not path.is_file() or "reports" in path.parts or path.suffix == ".pyc":
                    continue
                try:
                    text = path.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    continue
                names.update(re.findall(r"SKILL(?:\.md)?「([^」]+)」", text))
        missing = sorted(name for name in names if name not in headings)
        self.assertEqual(missing, [])


class TestDesignGuidance(unittest.TestCase):
    def test_failure_case_library_has_at_least_eight_complete_cases(self):
        text = (ROOT / "references" / "failure-cases.md").read_text(encoding="utf-8")
        cases = re.split(r"^## \d+\. ", text, flags=re.M)[1:]
        self.assertGreaterEqual(len(cases), 8)
        for case in cases:
            for field in ("**输入**", "**错误设计**", "**为什么错**", "**修正**", "**必须保留**", "**不适用条件**"):
                self.assertIn(field, case)

    def test_visual_concept_is_not_accepted_as_style_label_only(self):
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        section = text.split("### 3.6 视觉概念不是风格标签", 1)[1].split("---", 1)[0]
        for requirement in ("帮助观众理解", "影响焦点", "哪些页面需要变化", "不能被模板", "删除这个概念"):
            self.assertIn(requirement, section)
        self.assertIn("不是可执行概念", section)


class TestHandoff(unittest.TestCase):
    def test_handoff_is_derived_from_plan_and_covers_minimum_decisions(self):
        tmp = tmpdir()
        brief = write_brief(tmp)
        plan = tmp / "plan.json"
        self.assertEqual(support.run_cli("plan", brief, "--out", plan).returncode, 0)
        output = tmp / "handoff.md"
        result = support.run_cli("handoff", plan, "--out", output)
        self.assertEqual(result.returncode, 0, result.stderr)
        text = output.read_text(encoding="utf-8")
        for label in ("目标决定", "整稿主张", "页面任务", "核心主张", "证据状态", "主导焦点",
                      "构图理由", "必须保留", "禁止推断", "未决事项", "验收条件", "失败回退"):
            self.assertIn(label, text)
        self.assertIn("不是第二份机器真值", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
