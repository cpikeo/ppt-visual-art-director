#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import copy
import json
import unittest

import support
from support import tmpdir
from workflows import impact


class TestModificationImpact(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tmpdir()
        _, cls.deck, _, _ = support.pipeline(cls.tmp, mode="draft")

    def test_changed_page_expands_to_adjacent_and_same_composition(self):
        after = copy.deepcopy(self.deck)
        slide = after["slides"][1]
        text = next(element for element in slide["elements"] if element.get("type") == "text")
        text["text"] += "（修订）"
        result = impact(self.deck, after)
        self.assertEqual(result["changed"], [slide["id"]])
        self.assertTrue(result["adjacent_review"])
        self.assertIn(slide["id"], result["verification_scope"])
        self.assertIn("不能从任意既有 PPTX", result["limitation"])

    def test_cli_writes_deterministic_change_summary(self):
        before = self.tmp / "before.json"
        after_path = self.tmp / "after.json"
        before.write_text(json.dumps(self.deck, ensure_ascii=False), encoding="utf-8")
        changed = copy.deepcopy(self.deck)
        last_text = next(element for element in changed["slides"][-1]["elements"] if element.get("type") == "text")
        last_text["text"] += "。"
        after_path.write_text(json.dumps(changed, ensure_ascii=False), encoding="utf-8")
        output = self.tmp / "impact.json"
        result = support.run_cli("impact", before, after_path, "--out", output)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = support.load_json(output)
        self.assertEqual(data["schema"], "pvad-impact-v1")
        self.assertEqual(data["changed"], [changed["slides"][-1]["id"]])
        self.assertIn("schema_and_guard_for_entire_deck", data["required_checks"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
