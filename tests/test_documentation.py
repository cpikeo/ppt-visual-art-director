#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Active protocol references must use stable semantic anchors, not section numbers."""
import re
import unittest

import support
from support import ROOT, tmpdir, write_brief


class TestProtocolReferences(unittest.TestCase):
    def test_skill_declares_unique_stable_anchors(self):
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        anchors = re.findall(r'<a id="([a-z0-9-]+)"></a>', text)
        self.assertEqual(len(anchors), len(set(anchors)))
        required = {"workflow", "task-routing", "think-handoff", "authority-boundaries",
                    "design-judgment", "chinese-typography", "visual-order", "editorial-typography",
                    "asset-direction", "evidence-levels", "quality-gates"}
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
                # 一段引用可能连写多个书名号名字，全部都要能在 SKILL 里找到同名标题
                for chain in re.findall(r"SKILL(?:\.md)?((?:「[^」]+」)+)", text):
                    names.update(re.findall(r"「([^」]+)」", chain))
        missing = sorted(name for name in names if name not in headings)
        self.assertEqual(missing, [])


class TestDocumentationMatchesCode(unittest.TestCase):
    """文档说的和代码做的必须是同一件事：错误码、路径、命令都不能只在一边存在。"""

    def test_every_guard_error_code_is_documented(self):
        scripts = "\n".join(path.read_text(encoding="utf-8")
                            for path in sorted((ROOT / "scripts").glob("*.py")))
        emitted = set(re.findall(r'rep\.(?:error|warn)\(\s*"([A-Z_]+)"', scripts))
        # 动态码：orphan_issues 的 kind 会以大写形式进入报告（ORPHAN_LINE / LEAD_FRAGMENT …）
        typography = (ROOT / "scripts" / "typography.py").read_text(encoding="utf-8")
        emitted |= {kind.upper() for kind in re.findall(r'WrapIssue\("([a-z_]+)"', typography)}
        contract = (ROOT / "references" / "contract.md").read_text(encoding="utf-8")
        undocumented = sorted(code for code in emitted if code not in contract)
        self.assertEqual(undocumented, [], f"guard 会报这些码，但 contract.md 没有登记：{undocumented}")

    # 运行期生成的产物不在仓库里，文档提到它们是正常的。
    GENERATED = {"plan.json", "deck.json", "manifest.json", "brief.yaml", "inspection.json",
                 "impact.json", "media_briefs.json", "release-evidence.yaml",
                 "delivery-summary.md", "REVIEW_REQUIRED.md", "SHA256SUMS.txt"}

    @staticmethod
    def _repo_index() -> tuple[set, set]:
        paths, basenames = set(), set()
        for path in ROOT.rglob("*"):
            if not path.is_file() or ".git" in path.parts or "__pycache__" in path.parts:
                continue
            relative = path.relative_to(ROOT).as_posix()
            paths.add(relative)
            basenames.add(path.name)
        return paths, basenames

    def test_referenced_files_exist(self):
        """失效引用是交付缺陷：文档里以代码格式写出的仓库文件必须真实存在。"""
        paths, basenames = self._repo_index()
        missing = []
        sources = [ROOT / "SKILL.md"] + sorted((ROOT / "references").glob("*.md"))
        for source in sources:
            text = source.read_text(encoding="utf-8")
            for name in re.findall(r'`([A-Za-z0-9_./-]+\.(?:md|json|yaml|yml|py|txt))`', text):
                if name.startswith("http") or name in self.GENERATED:
                    continue
                # 生成物（每次评测/每份 deck 重新产出）不是仓库文件；文档提到它们是正常的。
                if name.startswith("evals/reports/"):
                    continue
                # 运行包有意不带 tests/ 与 evals/：只有在开发工作区才检查它们的引用。
                top = name.split("/", 1)[0]
                if top in {"tests", "evals"} and not (ROOT / top).is_dir():
                    continue
                if name in paths or name.lstrip("./") in paths:
                    continue
                # 允许省略目录的简写（如 `render.py`、`strategy.yaml`），但必须在仓库里唯一存在
                if name.rsplit("/", 1)[-1] in basenames:
                    continue
                missing.append(f"{source.name}: {name}")
        self.assertEqual(sorted(set(missing)), [])

    def test_documented_commands_have_entrypoints(self):
        """文档里写 `make xxx` 之类的命令，必须有对应的可执行入口。"""
        offenders = []
        for source in [ROOT / "SKILL.md"] + sorted((ROOT / "references").glob("*.md")):
            text = source.read_text(encoding="utf-8")
            for command in re.findall(r'`make ([a-z][a-z0-9_-]*)', text):
                if not (ROOT / "Makefile").exists():
                    offenders.append(f"{source.name}: make {command}")
        self.assertEqual(offenders, [], "文档引用了不存在的构建入口")

    def test_pipeline_help_lists_every_documented_subcommand(self):
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        documented = set(re.findall(r'python scripts/pvad\.py ([a-z-]+)', skill))
        result = support.run_cli("--help")
        usage = result.stdout + result.stderr
        missing = sorted(name for name in documented if name not in usage)
        self.assertEqual(missing, [], f"SKILL.md 写了这些子命令，但 CLI 没有：{missing}")


class TestDesignGuidance(unittest.TestCase):
    def test_failure_case_library_has_at_least_eight_complete_cases(self):
        text = (ROOT / "references" / "failure-cases.md").read_text(encoding="utf-8")
        cases = re.split(r"^## \d+\. ", text, flags=re.M)[1:]
        self.assertGreaterEqual(len(cases), 8)
        for case in cases:
            for field in ("**输入**", "**错误设计**", "**为什么错**", "**修正**", "**必须保留**", "**不适用条件**"):
                self.assertIn(field, case)

    def test_visual_order_reference_preserves_execution_and_conflict_priority(self):
        text = (ROOT / "references" / "visual-art-direction.md").read_text(encoding="utf-8")
        self.assertIn("理解内容 → 明确核心信息 → 确立视觉概念 → 建立阅读层级 → 组织构图与排版 → 选择和处理素材 → 检查实际呈现效果", text)
        self.assertIn("真实性与合规", text)
        self.assertIn("删减 → 重组 → 调整比例与空间 → 优化排版 → 处理素材 → 添加必要效果", text)
        for question in ("传达什么", "突出什么", "如何阅读", "为何这样表达", "什么可以删除"):
            self.assertIn(question, text)

    def test_typography_protocol_preserves_decision_order_and_editing_priority(self):
        text = (ROOT / "references" / "typography-editorial-composition.md").read_text(encoding="utf-8")
        for stage in ("THINK｜理解与编辑", "COMPOSE｜建立文字构图", "TYPESET｜字体与排版", "VERIFY｜实际渲染校验"):
            self.assertIn(stage, text)
        self.assertIn("删除重复 → 合并相近观点 → 缩短冗余表达 → 调整层级 → 重组构图 → 必要时拆页", text)
        self.assertIn("语义与事实准确 ＞ 阅读清晰 ＞ 信息层级 ＞ 构图与空间 ＞ 视觉表现 ＞ 装饰细节", text)

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
