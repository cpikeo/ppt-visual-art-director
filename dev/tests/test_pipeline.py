# -*- coding: utf-8 -*-
"""Skill 回归测试（开发材料，不进入正式 Skill 包）。运行：python -m unittest discover -s dev/tests -v
每个“应拦”用例都先确认“应过”的对照存在，避免只测到拒绝。"""
import hashlib, json, os, subprocess, sys, tempfile, unittest, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
S = ROOT / "scripts"
CASES = ROOT / "dev" / "cases"


def run(*args, cwd=None):
    p = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True, encoding="utf-8", cwd=cwd)
    return p.returncode, p.stdout + p.stderr


def base_plan(**over):
    page = {"id": "p1", "role": "open", "focus": "标题", "design_state": "judged",
            "composition_rationale": "标题先被看见：尺度最大、位于左上",
            "elements": [{"id": "t", "type": "title", "x": 0.8, "y": 0.8, "w": 10, "h": 1.2, "content": "测试标题"}]}
    plan = {"deck_id": "t", "title": "T", "audience": "a", "goal": "g", "core_claim": "c",
            "brief": {"insight": "i", "concept": "c", "veto": ["v"]}, "pages": [page]}
    plan.update(over)
    return plan


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, name, obj):
        p = self.d / name
        p.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")
        return p

    def compile(self, plan, out="o.pptx", *extra):
        p = self.write("plan.json", plan)
        return run(S / "compile.py", p, self.d / out, *extra)


class TestScaffoldCompile(Base):
    def test_full_deck_all_recipes(self):
        rc, out = run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        self.assertEqual(rc, 0, out)
        rc, out = run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        self.assertEqual(rc, 0, out)
        rc, out = run(S / "verify.py", self.d / "plan.json", self.d / "o.pptx")
        self.assertIn(rc, (0, 4), out)  # 无 LibreOffice 时 PARTIAL
        self.assertIn("对象核对：PASS", out)

    def test_eyebrow_aligned_with_title_when_no_motif(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        plan = json.loads((self.d / "plan.json").read_text(encoding="utf-8"))
        for pg in plan["pages"]:
            ids = {e["id"]: e for e in pg["elements"]}
            if "eyebrow" in ids:
                self.assertAlmostEqual(ids["eyebrow"]["x"], ids["t"]["x"], places=2, msg=pg["id"])

    def test_close_items_stay_inside_canvas(self):
        run(S / "scaffold.py", "init", CASES / "c2_energy_outline.json", self.d / "plan.json")
        (self.d / "assets").mkdir()
        (self.d / "assets" / "plate.png").write_bytes((CASES / "c2_plate.png").read_bytes())
        rc, out = run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        self.assertEqual(rc, 0, out)
        self.assertNotIn("W-GEOM", out)

    def test_deterministic_bytes(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        h = []
        for n in ("a.pptx", "b.pptx"):
            run(S / "compile.py", self.d / "plan.json", self.d / n)
            h.append(hashlib.sha256((self.d / n).read_bytes()).hexdigest())
        self.assertEqual(h[0], h[1])

    def test_production_blocks_draft(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline_draft.json", self.d / "plan.json")
        rc, out = run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx", "--production")
        self.assertEqual(rc, 2, out)
        self.assertFalse((self.d / "o.pptx").exists())

    def test_minimal_plan_ok_and_native_title(self):
        rc, out = self.compile(base_plan())
        self.assertEqual(rc, 0, out)
        from pptx import Presentation
        prs = Presentation(str(self.d / "o.pptx"))
        self.assertEqual(prs.slides[0].shapes.title.text_frame.text, "测试标题")

    def test_4_3_canvas(self):
        rc, out = self.compile(base_plan(aspect_ratio="4:3"))
        self.assertEqual(rc, 0, out)
        from pptx import Presentation
        prs = Presentation(str(self.d / "o.pptx"))
        self.assertAlmostEqual(prs.slide_width / 914400, 10.0, places=2)

    def test_no_overwrite_without_force(self):
        self.compile(base_plan())
        rc, out = self.compile(base_plan())
        self.assertNotEqual(rc, 0, out)
        rc, out = self.compile(base_plan(), "o.pptx", "--force")
        self.assertEqual(rc, 0, out)

    def test_chinese_and_space_in_path(self):
        sub = self.d / "我的 稿子"
        sub.mkdir()
        p = sub / "计划 plan.json"
        p.write_text(json.dumps(base_plan(), ensure_ascii=False), encoding="utf-8")
        rc, out = run(S / "compile.py", p, sub / "成品 v1.pptx")
        self.assertEqual(rc, 0, out)


class TestCost(Base):
    def test_fields_compact(self):
        rc, out = run(S / "scaffold.py", "fields", "element")
        self.assertEqual(rc, 0, out)
        self.assertIn("type*:", out)
        self.assertLess(len(out.encode()), 3000)

    def test_repeated_diagnostics_collapsed(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline_draft.json", self.d / "plan.json")
        rc, out = run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        self.assertEqual(rc, 0, out)
        self.assertEqual(out.count("W-PENDING"), 1, out)


class TestCharts(Base):
    def chart_plan(self, chart):
        pl = base_plan()
        pl["pages"][0]["elements"].append({"id": "ch", "type": "chart", "x": 0.8, "y": 2.2, "w": 11, "h": 4, "chart": {"title": "测试图", "period": "2025 Q3", "basis": "n=100", "unit": "%", **chart}, "alt": "图"})
        return pl

    def test_kinds_compile(self):
        cats = ["A", "B", "C"]
        ser = [{"name": "s1", "values": [1, 2, 3]}]
        kinds = {
            "bar": {"kind": "bar", "data": {"categories": cats, "series": ser}},
            "line": {"kind": "line", "data": {"categories": cats, "series": ser}},
            "area": {"kind": "area", "data": {"categories": cats, "series": ser}},
            "pie": {"kind": "pie", "data": {"categories": cats, "series": [{"name": "s", "values": [50, 30, 20]}]}, "data_labels": True, "legend": True},
            "donut": {"kind": "donut", "data": {"categories": cats, "series": [{"name": "s", "values": [50, 30, 20]}]}, "data_labels": True, "legend": True},
            "scatter": {"kind": "scatter", "data": {"series": [{"name": "s", "points": [[1, 2], [2, 3], [3, 5]]}]}},
            "table": {"kind": "table", "data": {"headers": ["a", "b"], "rows": [["1", "2"]]}},
        }
        for k, ch in kinds.items():
            with self.subTest(kind=k):
                rc, out = self.compile(self.chart_plan(ch), f"{k}.pptx")
                self.assertNotIn("Traceback", out)
                self.assertEqual(rc, 0, out)

    def test_nan_rejected_not_crash(self):
        ch = {"kind": "bar", "data": {"categories": ["A", "B"], "series": [{"name": "s", "values": [1, "NaN"]}]}}
        rc, out = self.compile(self.chart_plan(ch))
        self.assertNotIn("Traceback", out)
        self.assertEqual(rc, 2, out)


class TestSafety(Base):
    def test_malformed_json_exit2(self):
        p = self.d / "bad.json"
        p.write_text("{not json", encoding="utf-8")
        rc, out = run(S / "compile.py", p, self.d / "o.pptx")
        self.assertEqual(rc, 2, out)
        self.assertNotIn("Traceback", out)

    def test_missing_plan_exit2(self):
        rc, out = run(S / "compile.py", self.d / "nope.json", self.d / "o.pptx")
        self.assertEqual(rc, 2, out)
        self.assertNotIn("Traceback", out)

    def test_unknown_element_key_rejected(self):
        pl = base_plan()
        pl["pages"][0]["elements"][0]["typo_key"] = 1
        rc, out = self.compile(pl)
        self.assertEqual(rc, 2, out)
        self.assertIn("E-SCHEMA", out)

    def test_asset_path_traversal_blocked(self):
        (self.d / "secret.png").write_bytes(b"x")
        sub = self.d / "proj"
        sub.mkdir()
        pl = base_plan(assets={"../secret": {"source": "x", "rights": "cleared"}})
        pl["pages"][0]["elements"].append({"id": "im", "type": "image", "x": 1, "y": 3, "w": 3, "h": 2, "image_id": "../secret"})
        p = sub / "plan.json"
        p.write_text(json.dumps(pl, ensure_ascii=False), encoding="utf-8")
        rc, out = run(S / "compile.py", p, sub / "o.pptx")
        self.assertEqual(rc, 2, out)
        self.assertFalse((sub / "o.pptx").exists())

    def test_output_cannot_overwrite_plan(self):
        p = self.write("plan.json", base_plan())
        rc, out = run(S / "compile.py", p, p)
        self.assertNotEqual(rc, 0, out)
        self.assertEqual(json.loads(p.read_text(encoding="utf-8"))["deck_id"], "t")


class TestRevise(Base):
    def test_text_edit_preserves_everything_else(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        orig = (self.d / "o.pptx").read_bytes()
        edit = {"deck_id": "e", "intent": "改标题", "ops": [{"op": "text", "page": 1, "target": "p1.t", "content": "续费率在第九十天决定"}]}
        self.write("edit.json", edit)
        rc, out = run(S / "revise.py", self.d / "o.pptx", self.d / "edit.json", self.d / "n.pptx")
        self.assertEqual(rc, 0, out)
        self.assertEqual((self.d / "o.pptx").read_bytes(), orig)
        from pptx import Presentation
        self.assertEqual(Presentation(str(self.d / "n.pptx")).slides[0].shapes.title.text_frame.text, "续费率在第九十天决定")
        with zipfile.ZipFile(self.d / "o.pptx") as a, zipfile.ZipFile(self.d / "n.pptx") as b:
            changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
        self.assertEqual(changed, ["ppt/slides/slide1.xml"])

    def test_chart_data_edit(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        edit = {"deck_id": "e", "intent": "改数", "ops": [{"op": "chart_data", "page": 3, "target_index": 0,
                "data": {"categories": ["已激活", "未激活"], "series": [{"name": "续费率", "values": [92, 55]}]}}]}
        self.write("edit.json", edit)
        rc, out = run(S / "revise.py", self.d / "o.pptx", self.d / "edit.json", self.d / "n.pptx")
        self.assertNotIn("Traceback", out)

    def test_unknown_op_rejected(self):
        run(S / "scaffold.py", "init", CASES / "c1_retention_outline.json", self.d / "plan.json")
        run(S / "compile.py", self.d / "plan.json", self.d / "o.pptx")
        self.write("edit.json", {"deck_id": "e", "intent": "x", "ops": [{"op": "explode", "page": 1, "target": "p1.t"}]})
        rc, out = run(S / "revise.py", self.d / "o.pptx", self.d / "edit.json", self.d / "n.pptx")
        self.assertEqual(rc, 2, out)
        self.assertFalse((self.d / "n.pptx").exists())


if __name__ == "__main__":
    unittest.main()
