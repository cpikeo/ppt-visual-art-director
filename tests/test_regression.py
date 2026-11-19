"""回归验证：定点编辑的保留核对、渲染文字归属、图表直标。

每条负向用例都先人为破坏一个属性，再确认核对真的会响。没有反向构造过的绿灯不算证据。
运行：python -m unittest discover -s tests -v   （渲染用例需要 LibreOffice，缺失时跳过）
"""
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import compile as C  # noqa: E402
import inspect_pptx as I  # noqa: E402
import revise as R  # noqa: E402
import verify as V  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE  # noqa: E402
from pptx.util import Inches  # noqa: E402

PLAN = ROOT / "tests" / "fixtures" / "min_plan.json"


def run(script: str, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)],
                          capture_output=True, text=True, encoding="utf-8")


class Base(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="pptreg-"))
        self.src = self.tmp / "src.pptx"
        done = run("compile.py", PLAN, self.src)
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def edit(self, ops: list, src: Path = None) -> tuple:
        spec = self.tmp / "edit.json"
        spec.write_text(json.dumps({"deck_id": "t", "ops": ops}, ensure_ascii=False), encoding="utf-8")
        out = self.tmp / "out.pptx"
        done = run("revise.py", src or self.src, spec, out)
        return done, out

    def ids(self, path: Path, page: int = 1) -> dict:
        return {o["name"]: o["id"] for o in I.inspect(path)["pages"][page - 1]["objects"]}

    def compare_paths(self, after: Path, grants=None, deleted=None) -> R.C.Report:
        rep = R.C.Report()
        R.compare(I.inspect(self.src), I.inspect(after), grants or {}, set(), deleted or set(), rep)
        return rep

    def mutate(self, fn) -> Path:
        prs = Presentation(str(self.src))
        fn(prs)
        path = self.tmp / "mut.pptx"
        prs.save(str(path))
        return path


class RevisePreserve(Base):
    def test_clean_edit_is_ready_and_leaves_others_alone(self) -> None:
        done, out = self.edit([
            {"op": "text", "page": 1, "target": "p1.t", "content": "留存才是增长（改）"},
            {"op": "box", "page": 1, "target": "p1.t", "y": 4.5},
            {"op": "style", "page": 1, "target": "p1.sub", "font_size": 22},
        ])
        self.assertEqual(done.returncode, 0, done.stdout)
        self.assertIn("保留核对：READY", done.stdout)
        self.assertTrue(out.exists())
        grants = {}
        for page_obj, name, grant in (
            (1, "p1.t", {"top": {"text", "box_in"}, "sig": {"font_size", "color"}, "frame": set(), "text": True}),
            (1, "p1.sub", {"top": set(), "sig": {"font_size"}, "frame": set(), "text": False}),
        ):
            grants[(page_obj, self.ids(self.src)[name])] = grant
        self.assertFalse(self.compare_paths(out, grants).blocked, "授权范围内的改动不应被判为漂移")

    def test_unknown_key_is_rejected_not_ignored(self) -> None:
        done, out = self.edit([{"op": "box", "page": 1, "target": "p1.t", "y": 4.5, "rotation": 90}])
        self.assertEqual(done.returncode, 2)
        self.assertIn("不接受", done.stdout)
        self.assertFalse(out.exists())

    def test_text_op_rejects_frame_keys_it_cannot_apply(self) -> None:
        done, out = self.edit([{"op": "text", "page": 1, "target": "p1.t", "content": "x", "pad_in": 0}])
        self.assertEqual(done.returncode, 2)
        self.assertFalse(out.exists())

    def test_unauthorized_fill_change_on_untouched_object_blocks(self) -> None:
        def paint(prs):
            shape = [s for s in prs.slides[0].shapes if s.name == "p1.sub"][0]
            shape.fill.solid()
            from pptx.dml.color import RGBColor
            shape.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
        rep = self.compare_paths(self.mutate(paint))
        self.assertTrue(rep.blocked)

    def test_touched_object_may_only_change_granted_attributes(self) -> None:
        def paint(prs):
            shape = [s for s in prs.slides[0].shapes if s.name == "p1.sub"][0]
            from pptx.dml.color import RGBColor
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
        after = self.mutate(paint)
        sid = self.ids(self.src)["p1.sub"]
        only_box = {(1, sid): {"top": {"box_in"}, "sig": set(), "frame": set(), "text": False}}
        only_fill = {(1, sid): {"top": {"fill"}, "sig": set(), "frame": set(), "text": False}}
        self.assertTrue(self.compare_paths(after, only_box).blocked, "box 授权不应放行填色改动")
        self.assertFalse(self.compare_paths(after, only_fill).blocked, "fill 授权应放行填色改动")

    def test_hidden_slide_flag_is_preserved(self) -> None:
        def hide(prs):
            prs.slides[0]._element.set("show", "0")
        self.assertTrue(self.compare_paths(self.mutate(hide)).blocked)

    def test_chart_value_beyond_first_twelve_points_is_seen(self) -> None:
        def build(path: Path, value_13: float) -> Path:
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            data = CategoryChartData()
            data.categories = [f"c{i}" for i in range(14)]
            data.add_series("s", [float(i) for i in range(14)])
            slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(8),
                                   Inches(5), data)
            prs.save(str(path))
            prs = Presentation(str(path))
            chart = [s for s in prs.slides[0].shapes if s.has_chart][0].chart
            values = [float(i) for i in range(14)]
            values[12] = value_13
            cd = CategoryChartData()
            cd.categories = [f"c{i}" for i in range(14)]
            cd.add_series("s", values)
            chart.replace_data(cd)
            out = path.with_name(path.stem + f"_{value_13}.pptx")
            prs.save(str(out))
            return out

        base = build(self.tmp / "c.pptx", 12.0)
        changed = build(self.tmp / "c2.pptx", 99.0)
        before, after = I.inspect(base), I.inspect(changed)
        self.assertNotEqual(before["pages"][0]["objects"][0].get("payload"),
                            after["pages"][0]["objects"][0].get("payload"))

    def test_slide_order_mapping_follows_sldidlst_not_filenames(self) -> None:
        def swap(prs):
            prs.slides  # 先按当前顺序完成 python-pptx 的部件编号
            lst = prs.slides._sldIdLst
            items = list(lst)
            lst.remove(items[0])
            lst.append(items[0])
        prs = Presentation(str(self.src))
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "第二页"
        two = self.tmp / "two.pptx"
        prs.save(str(two))
        prs = Presentation(str(two))
        swap(prs)
        reordered = self.tmp / "reordered.pptx"
        prs.save(str(reordered))
        names = I.slide_part_names(reordered)
        self.assertEqual(len(names), 2)
        self.assertNotEqual(names, sorted(names), "应能构造出页序与部件名不一致的文件")

    def test_duplicate_names_are_targeted_by_index(self) -> None:
        def dup(prs):
            slide = prs.slides[0]
            twin = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(3), Inches(0.5))
            twin.name = "p1.sub"
            twin.text_frame.text = "重名对象"
        dup_path = self.tmp / "dup.pptx"
        prs = Presentation(str(self.src))
        dup(prs)
        prs.save(str(dup_path))
        done, _ = self.edit([{"op": "style", "page": 1, "target_index": 2, "font_size": 14}], src=dup_path)
        self.assertEqual(done.returncode, 0, done.stdout)
        done, _ = self.edit([{"op": "style", "page": 1, "target": "p1.sub", "font_size": 14}], src=dup_path)
        self.assertEqual(done.returncode, 2, "重名时按名字定位必须阻断，不猜")


class RenderAttribution(Base):
    def setUp(self) -> None:
        super().setUp()
        if V.find_soffice() is None:
            self.skipTest("未安装 LibreOffice，渲染归属用例跳过")

    def test_text_missing_in_element_but_present_elsewhere_fails(self) -> None:
        prs = Presentation(str(self.src))
        target = [s for s in prs.slides[0].shapes if s.name == "p1.t"][0]
        target.text_frame.text = ""  # 本元素的字没了；页面另一处 dup 仍有同样的字
        broken = self.tmp / "broken.pptx"
        prs.save(str(broken))
        plan = json.loads(PLAN.read_text(encoding="utf-8"))
        render_dir = self.tmp / "render"
        render_dir.mkdir()
        rep = C.Report()
        pdf = V.render(broken, render_dir, rep)
        self.assertIsNotNone(pdf)
        V.verify_render(plan, broken, pdf, render_dir, rep)
        hits = [str(item) for item in rep.items if "V-TEXT-RENDER" in str(item) and "p1/t" in str(item)]
        self.assertTrue(hits, f"本元素缺字必须被归属到 p1/t，实际：{[str(i) for i in rep.items]}")

    def test_clean_render_has_no_element_text_miss(self) -> None:
        plan = json.loads(PLAN.read_text(encoding="utf-8"))
        render_dir = self.tmp / "render"
        render_dir.mkdir()
        rep = C.Report()
        pdf = V.render(self.src, render_dir, rep)
        V.verify_render(plan, self.src, pdf, render_dir, rep)
        self.assertFalse([i for i in rep.items if "V-TEXT-RENDER" in str(i)], str(rep.items))


class ChartLabels(unittest.TestCase):
    def test_default_label_pieces_carry_unit(self) -> None:
        chart = {"kind": "bar", "data_labels": True, "unit": "亿元",
                 "data": {"series": [{"name": "s", "values": [1234.5, 2]}]}}
        self.assertEqual(V._label_pieces(chart), ["1,234.5亿元", "2亿元"])

    def test_custom_number_format_is_not_claimed(self) -> None:
        chart = {"kind": "bar", "data_labels": True, "unit": "亿元", "value_axis": {"number_format": "0.00"},
                 "data": {"series": [{"name": "s", "values": [1.0]}]}}
        self.assertEqual(V._label_pieces(chart), [])

    def test_long_unit_is_not_silently_dropped(self) -> None:
        fmt = C._label_format({"unit": "人民币万元"})
        self.assertIn("人民币万元", fmt)


class RealWorldFiles(Base):
    """非 python-pptx 生成的既有稿：LibreOffice 另存后再改稿，不能被排版差异误判成改动。"""

    def test_libreoffice_resaved_deck_edit_is_ready(self) -> None:
        if V.find_soffice() is None:
            self.skipTest("未安装 LibreOffice")
        out_dir = self.tmp / "lo"
        out_dir.mkdir()
        subprocess.run([V.find_soffice(), "--headless", "--convert-to", "pptx", "--outdir", str(out_dir),
                        str(self.src)], capture_output=True, timeout=240)
        resaved = out_dir / self.src.name
        self.assertTrue(resaved.exists())
        done, out = self.edit([{"op": "text", "page": 1, "target": "p1.t", "content": "留存才是增长（改）"}],
                              src=resaved)
        self.assertEqual(done.returncode, 0, done.stdout)
        self.assertIn("保留核对：READY", done.stdout)


class ChartRender(Base):
    def setUp(self) -> None:
        super().setUp()
        if V.find_soffice() is None:
            self.skipTest("未安装 LibreOffice，图表直标渲染用例跳过")

    def test_label_unit_mismatch_with_render_is_reported(self) -> None:
        chart_plan = ROOT / "tests" / "fixtures" / "chart_plan.json"
        pptx = self.tmp / "chart.pptx"
        self.assertEqual(run("compile.py", chart_plan, pptx).returncode, 0)
        plan = json.loads(chart_plan.read_text(encoding="utf-8"))

        def check(spec: dict) -> list:
            render_dir = self.tmp / ("r" + str(abs(hash(json.dumps(spec)))))
            render_dir.mkdir()
            rep = C.Report()
            pdf = V.render(pptx, render_dir, rep)
            V.verify_render(spec, pptx, pdf, render_dir, rep)
            return [str(i) for i in rep.items if "V-TEXT-RENDER" in str(i)]

        self.assertEqual(check(plan), [], "正确声明的直标必须通过")
        wrong = json.loads(json.dumps(plan))
        wrong["pages"][0]["elements"][1]["chart"]["unit"] = "万元"
        hits = check(wrong)
        self.assertTrue(any("p1/c1" in h for h in hits), f"单位不一致必须报出：{hits}")


if __name__ == "__main__":
    unittest.main()
