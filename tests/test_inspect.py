#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import unittest

import support
from support import ROOT, tmpdir


class TestPptxInspection(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        _, cls.deck, _, cls.artifacts = support.pipeline(cls.tmp, mode="draft")
        cls.output = cls.tmp / "inspection.json"
        cls.result = support.run_cli("inspect", cls.artifacts["pptx"], "--out", cls.output)
        cls.inspection = support.load_json(cls.output)

    def test_inspection_is_schema_valid_and_bound_to_pptx(self):
        self.assertEqual(self.result.returncode, 0, self.result.stderr)
        from pvad import sha256_file, validate_against
        self.assertEqual(validate_against(self.inspection, "inspection.schema.json"), [])
        self.assertEqual(self.inspection["artifact_sha256"], sha256_file(self.artifacts["pptx"]))

    def test_inspection_contains_observable_objects_without_binary_blobs(self):
        self.assertEqual(self.inspection["stats"]["slides"], len(self.deck["slides"]))
        self.assertGreater(self.inspection["stats"]["objects"], 0)
        self.assertGreaterEqual(self.inspection["stats"]["charts"], 1)
        for slide in self.inspection["slides"]:
            for obj in slide["objects"]:
                self.assertNotIn("image_blob", obj)
                for key in ("name", "x", "y", "width", "height", "type"):
                    self.assertIn(key, obj)

    def test_inspection_refuses_to_infer_design_intent(self):
        joined = " ".join(self.inspection["limitations"])
        self.assertIn("不恢复作者主张", joined)
        self.assertIn("不自动解释为因果", joined)
        self.assertNotIn("claim", self.inspection)
        self.assertNotIn("evidence", self.inspection)

    def test_contact_sheet_uses_policy_thumbnail_width(self):
        from PIL import Image
        sheet = self.artifacts["report"] / "preview-pptx.png"
        with Image.open(sheet) as image:
            self.assertLessEqual(image.width, 2 * 480 + 3 * 16)

    def test_resource_budget_rejects_before_deep_inspection(self):
        import copy
        from policy import Policy, load as load_policy
        from pptx_inspect import InspectError, inspect_pptx
        original = load_policy()
        strategy = copy.deepcopy(original.strategy)
        strategy["security"]["pptx"]["max_file_bytes"] = 1
        constrained = Policy(strategy=strategy, composition=original.composition,
                             sha256=original.sha256, paths=original.paths)
        with self.assertRaisesRegex(InspectError, "PPTX_RESOURCE_BUDGET"):
            inspect_pptx(self.artifacts["pptx"], constrained)

    def test_inspection_v1_migrates_to_v2_and_validates(self):
        import copy
        from migrate import migrate_inspection
        current = json.loads((self.tmp / "inspection.json").read_text())
        legacy = copy.deepcopy(current)
        legacy["schema"] = "pvad-inspection-v1"
        legacy.pop("accessibility")
        legacy.pop("font_diagnostics")
        for slide in legacy["slides"]:
            slide.pop("layout")
            slide.pop("notes_present")
        migrated = migrate_inspection(legacy)
        self.assertEqual(migrated["schema"], "pvad-inspection-v2")
        from jsonschema import Draft202012Validator
        Draft202012Validator(support.load_json(ROOT / "schema" / "inspection.schema.json")).validate(migrated)

    def test_non_pptx_is_rejected(self):
        fake = self.tmp / "fake.txt"
        fake.write_text("not pptx", encoding="utf-8")
        result = support.run_cli("inspect", fake, "--out", self.tmp / "bad.json")
        self.assertEqual(result.returncode, 2)
        self.assertIn(".pptx", result.stderr)


class TestModificationProtocol(unittest.TestCase):
    def test_protocol_contains_all_required_stages_and_boundaries(self):
        text = (ROOT / "references" / "modification-workflow.md").read_text(encoding="utf-8")
        for stage in ("INSPECT", "人工补充最小交接", "EDIT", "IMPACT", "VERIFY", "CHANGE SUMMARY"):
            self.assertIn(stage, text)
        self.assertIn("不能恢复作者主张", text)
        self.assertIn("旧 evidence 和旧 Release 包立即失效", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
