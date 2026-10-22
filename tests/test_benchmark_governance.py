#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import copy
import json
import unittest

import yaml
from jsonschema import Draft202012Validator

from support import ROOT


class TestBenchmarkGovernance(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.schema = json.loads((ROOT / "schema" / "benchmark-case.schema.json").read_text())
        cls.template = yaml.safe_load((ROOT / "evals" / "benchmarks" / "benchmark-case.template.yaml").read_text())
        cls.validator = Draft202012Validator(cls.schema)

    def test_draft_template_is_valid_but_explicitly_not_independent_evidence(self):
        self.assertEqual(list(self.validator.iter_errors(self.template)), [])
        self.assertEqual(self.template["status"], "DRAFT")
        self.assertEqual(self.template["expert_labels"], [])
        self.assertIn("do not claim", self.template["notes"])

    def test_frozen_case_requires_two_independent_experts_and_fixed_artifacts(self):
        value = copy.deepcopy(self.template)
        value["status"] = "FROZEN"
        errors = list(self.validator.iter_errors(value))
        self.assertTrue(errors)
        joined = " ".join(error.message for error in errors)
        self.assertIn("too short", joined)

    def test_protocol_covers_required_project_families_and_exposure_rule(self):
        text = (ROOT / "evals" / "benchmarks" / "README.md").read_text(encoding="utf-8")
        for name in ("经营复盘", "融资路演", "项目立项", "产品发布", "培训材料", "数据报告", "既有 PPTX 改稿"):
            self.assertIn(name, text)
        self.assertIn("EXPOSED_REGRESSION", text)
        self.assertIn("不能继续宣称独立泛化证据", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
