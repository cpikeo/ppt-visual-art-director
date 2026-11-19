"""回归验证：定点编辑的保留核对（富文本、超链接、OOXML 结构）、包格式、素材证据边界、渲染归属与状态诚实度。

每条负向用例都先人为破坏一个属性，再确认核对真的会响。没有反向构造过的绿灯不算证据。
运行：python -m unittest discover -s tests -v   （渲染用例需要 LibreOffice，缺失时跳过）
"""
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import compile as C  # noqa: E402
import revise as R  # noqa: E402
import verify as V  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn as pqn  # noqa: E402
from lxml import etree  # noqa: E402
from PIL import Image  # noqa: E402

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
        return {o["name"]: o["id"] for o in R.inspect_deck(path)["pages"][page - 1]["objects"]}

    def compare_paths(self, after: Path, grants=None, deleted=None) -> R.C.Report:
        rep = R.C.Report()
        R.compare(R.inspect_deck(self.src), R.inspect_deck(after), grants or {}, set(), deleted or set(), rep)
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
        before, after = R.inspect_deck(base), R.inspect_deck(changed)
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
        names = R.slide_part_names(reordered)
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



class RichPreserve(Base):
    """既有稿的富文本、超链接与 OOXML 结构：改一处，其余逐项保留。"""

    def make_rich(self) -> Path:
        prs = Presentation(str(self.src))
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
        box.name = "rich"
        tf = box.text_frame
        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = "访问"
        r1.font.size = Pt(20)
        r1.hyperlink.address = "https://example.com/docs"
        rpr = r1._r.get_or_add_rPr()
        rpr.set("u", "sng")
        ea = etree.Element(pqn("a:ea"))
        ea.set("typeface", "等线")
        rpr.find(pqn("a:hlinkClick")).addprevious(ea)  # 东亚字族须位于 hlinkClick 之前
        r2 = p1.add_run()
        r2.text = "官网"
        r2.font.bold = True
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r3 = p2.add_run()
        r3.text = "第二段"
        path = self.tmp / "rich.pptx"
        prs.save(str(path))
        return path

    def test_text_edit_keeps_hyperlink_underline_ea_and_paragraph_align(self) -> None:
        src = self.make_rich()
        done, out = self.edit([{"op": "text", "page": 2, "target": "rich",
                                "content": "访问（改）官网\n第二段（改）",
                                "runs": [{"text": "访问（改）"}, {"text": "官网", "bold": True},
                                         {"text": "\n"}, {"text": "第二段（改）"}]}], src=src)
        self.assertEqual(done.returncode, 0, done.stdout)
        self.assertIn("保留核对：READY", done.stdout)
        shape = [s for s in Presentation(str(out)).slides[1].shapes if s.name == "rich"][0]
        first = shape.text_frame.paragraphs[0].runs[0]
        self.assertEqual(first.text, "访问（改）")
        self.assertEqual(first.hyperlink.address, "https://example.com/docs", "超链接必须保留")
        rpr = first._r.find(pqn("a:rPr"))
        self.assertEqual(rpr.get("u"), "sng", "下划线丢失")
        self.assertEqual(rpr.find(pqn("a:ea")).get("typeface"), "等线", "东亚字族丢失")
        self.assertEqual(shape.text_frame.paragraphs[1].alignment, PP_ALIGN.CENTER, "段落对齐丢失")

    def test_hyperlink_is_refused_when_runs_do_not_map(self) -> None:
        src = self.make_rich()
        done, out = self.edit([{"op": "text", "page": 2, "target": "rich", "content": "唯一的一段"}], src=src)
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("超链接", done.stdout)
        self.assertFalse(out.exists())

    def test_untouched_shape_xml_change_is_blocked(self) -> None:
        def tweak(prs):
            run = [s for s in prs.slides[0].shapes if s.name == "p1.sub"][0].text_frame.paragraphs[0].runs[0]
            run._r.get_or_add_rPr().set("u", "sng")
        self.assertTrue(self.compare_paths(self.mutate(tweak)).blocked, "未点名对象的 XML 改动必须被拦下")

    def test_slide_feature_added_outside_shape_tree_is_blocked(self) -> None:
        def add_transition(prs):
            el = etree.SubElement(prs.slides[0]._element, pqn("p:transition"))
            el.set("spd", "fast")
        self.assertTrue(self.compare_paths(self.mutate(add_transition)).blocked)

    def test_relationship_loss_is_blocked(self) -> None:
        linked = self.tmp / "linked.pptx"
        prs = Presentation(str(self.src))
        shape = [s for s in prs.slides[0].shapes if s.name == "p1.sub"][0]
        shape.click_action.hyperlink.address = "https://example.com/x"
        prs.save(str(linked))
        broken = self.tmp / "broken.pptx"
        with zipfile.ZipFile(linked) as zin, zipfile.ZipFile(broken, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("ppt/slides/_rels/slide") and b"example.com/x" in data:
                    root = etree.fromstring(data)
                    for rel in list(root):
                        if "example.com/x" in (rel.get("Target") or ""):
                            root.remove(rel)  # 形状上的引用仍在，只删关系本身
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
        rep = C.Report()
        R.compare(R.inspect_deck(linked), R.inspect_deck(broken), {}, set(), set(), rep)
        self.assertTrue(rep.blocked, "关系丢失必须被拦下")
        self.assertTrue(any("关系丢失" in str(i) for i in rep.items), str(rep.items))

    def test_clean_edit_keeps_other_shapes_xml_identical(self) -> None:
        done, out = self.edit([{"op": "style", "page": 1, "target": "p1.sub", "font_size": 22}])
        self.assertEqual(done.returncode, 0, done.stdout)
        sub_id = self.ids(self.src)["p1.sub"]
        before = R.slide_xml_facts(self.src, R.slide_part_names(self.src))[0]
        after = R.slide_xml_facts(out, R.slide_part_names(out))[0]
        after_by = {s["id"]: s["xml"] for s in after["shapes"]}
        others = [s for s in before["shapes"] if s["id"] != sub_id]
        self.assertTrue(others)
        self.assertTrue(all(after_by[s["id"]] == s["xml"] for s in others))


class PackageFormat(Base):
    def test_pptm_source_is_refused(self) -> None:
        renamed = self.tmp / "macro.pptm"
        shutil.copy(self.src, renamed)
        done, out = self.edit([{"op": "box", "page": 1, "target": "p1.t", "y": 4.5}], src=renamed)
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("E-FORMAT", done.stdout)
        self.assertFalse(out.exists())

    def test_macro_content_type_and_part_are_detected(self) -> None:
        bad = self.tmp / "macro.pptx"
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(b"presentationml.presentation.main+xml",
                                        b"presentationml.presentation.macroEnabled.main+xml")
                zout.writestr(item.filename, data)
            zout.writestr("ppt/vbaProject.bin", b"x")
        self.assertTrue(C.package_problems(bad), "宏启用内容类型必须被识别")
        self.assertEqual(C.package_problems(self.src), [], "正常 .pptx 不得被误报")

    def test_compile_never_writes_pptm(self) -> None:
        out = self.tmp / "deck.pptm"
        self.assertEqual(run("compile.py", PLAN, out).returncode, 2)
        self.assertFalse(out.exists())

    def test_compile_refuses_overwrite_without_force(self) -> None:
        self.assertEqual(run("compile.py", PLAN, self.src).returncode, 2)
        self.assertEqual(run("compile.py", PLAN, self.src, "--force").returncode, 0)


class AssetEvidence(unittest.TestCase):
    """素材元数据：generated 字段、generation 记录与证据边界必须一致，不得靠字段组合绕过。"""

    PNG_BYTES = None

    def setUp(self) -> None:
        import hashlib
        self.tmp = Path(tempfile.mkdtemp(prefix="asset-"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.assets = self.tmp / "assets"
        self.assets.mkdir()
        Image.new("RGB", (40, 30), (20, 20, 20)).save(self.assets / "hero.png")
        self.sha = hashlib.sha256((self.assets / "hero.png").read_bytes()).hexdigest()

    def gen(self, sha: str = None, drop: str = None) -> dict:
        rec = {"tool": "图像生成工具", "prompt": "概念产品摄影", "file": "hero.png",
               "status": "generated", "sha256": sha or self.sha}
        if drop:
            del rec[drop]
        return rec

    def compile_with(self, asset: dict) -> subprocess.CompletedProcess:
        plan = json.loads(PLAN.read_text(encoding="utf-8"))
        plan["pages"][0]["elements"].append({"id": "img", "type": "image", "image_id": "hero",
                                             "x": 11.0, "y": 0.5, "w": 1.5, "h": 1.0,
                                             "asset_reason": "测试用素材，只验证元数据核对"})
        plan["assets"] = {"hero": asset}
        plan_path = self.tmp / "plan.json"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        out = self.tmp / "out.pptx"
        if out.exists():
            out.unlink()
        return run("compile.py", plan_path, out)

    def test_consistent_generated_asset_with_real_hash_is_accepted(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "pending", "generated": True,
                                  "generation": self.gen()})
        self.assertEqual(done.returncode, 0, done.stdout)

    def test_generated_false_with_generation_record_is_blocked(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "pending", "generated": False,
                                  "generation": self.gen()})
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("E-ASSET", done.stdout)

    def test_generated_true_without_record_is_blocked(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "pending", "generated": True})
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("E-ASSET", done.stdout)

    def test_generated_image_cannot_be_evidence(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "cleared", "reference": "无",
                                  "used_as_evidence": True, "generated": True, "generation": self.gen()})
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("不能作证据", done.stdout)

    def test_hash_mismatch_is_blocked(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "pending", "generated": True,
                                  "generation": self.gen(sha="1" * 64)})
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("SHA-256", done.stdout)

    def test_sha256_is_mandatory_for_generated_images(self) -> None:
        done = self.compile_with({"source": "自建生成", "rights": "pending", "generated": True,
                                  "generation": self.gen(drop="sha256")})
        self.assertEqual(done.returncode, 2, done.stdout)
        self.assertIn("sha256", done.stdout)


class StatusHonesty(Base):
    def test_status_only_ready_when_render_passed_and_fonts_present(self) -> None:
        report = self.tmp / "verify.json"
        done = run("verify.py", PLAN, self.src, "--report", report, "--render-dir", self.tmp / "render")
        data = json.loads(report.read_text(encoding="utf-8"))
        layers = data["layers"]
        if layers["render"] != "PASS":
            self.assertEqual(data["status"], "PARTIAL", "没有渲染依据时不能是 READY")
        if layers.get("fonts") != "OK":
            self.assertNotEqual(data["status"], "READY", "字体被替换或未检查时不能是 READY")
        if data["status"] == "READY":
            self.assertEqual(done.returncode, 0)


if __name__ == "__main__":
    unittest.main()
