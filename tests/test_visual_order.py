#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import unittest

import support  # adds scripts/ to sys.path for direct module tests
from compose import Frame, _groups
from policy import load


class TestVisualOrderComposition(unittest.TestCase):
    def test_left_image_narrative_changes_direction_without_changing_roles(self):
        policy = load()
        right = policy.recipe("image_narrative")
        left = policy.recipe("image_narrative_left")
        right_slots = {slot["role"]: slot for slot in right["slots"]}
        left_slots = {slot["role"]: slot for slot in left["slots"]}
        self.assertGreater(right_slots["media"]["zone"][0], right_slots["claim"]["zone"][0])
        self.assertLess(left_slots["media"]["zone"][0], left_slots["claim"]["zone"][0])
        self.assertEqual(set(right_slots), set(left_slots))

    def test_four_semantic_groups_render_as_four_visible_groups(self):
        policy = load()
        frame = Frame({"width": 1280, "height": 720}, policy)
        elements = []

        def add_text_box(role, text, box, style, color, **kwargs):
            elements.append({"role": role, "text": text, "box": box, "style": style, "color": color})

        atoms = {"bullets": ["风土：地块", "时间：年份", "酿造：克制", "仪式：体验"], "sentences": []}
        slots = {"groups": {"zone": [0.0, 0.3, 1.0, 0.55]}}
        _groups(elements, slots, atoms, frame, policy, None, add_text_box, False)
        titles = [e for e in elements if e["role"].startswith("group_title")]
        bodies = [e for e in elements if e["role"].startswith("group_body")]
        self.assertEqual([e["text"] for e in titles], ["风土", "时间", "酿造", "仪式"])
        self.assertEqual(len(bodies), 4)
        self.assertEqual(len({(e["box"][0], e["box"][1]) for e in titles}), 4)


if __name__ == "__main__":
    unittest.main(verbosity=2)
