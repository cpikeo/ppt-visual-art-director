#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 程序正确性与契约回归测试

只验证实际存在的程序行为与数据契约（scripts/compile_pptx.py <-> references/contract.md），
不用于替代 AI 设计质量判断。
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.oxml.ns import qn

# 保证能直接从仓库根运行 pytest 并导入 scripts
REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from scripts.compile_pptx import (  # noqa: E402
    ErrorLevel,
    REQUIRED_DECK_FIELDS,
    REQUIRED_ELEMENT_FIELDS,
    REQUIRED_PAGE_FIELDS,
    _extract_numbers,
    compile_plan,
    load_plan,
    main,
    validate_plan,
)

# ---------------------------------------------------------------------------
# 测试基准 plan
# ---------------------------------------------------------------------------

VALID_PLAN = {
    "deck_id": "deck-001",
    "title": "季度业绩回顾",
    "audience": "领导团队",
    "goal": "批准下一季度预算",
    "core_claim": "核心业务增长 25%，建议预算增加 15%",
    "evidence": ["Q3 收入达 1.2 亿元", "用户数增长 40%"],
    "visual_concept": "克制、数据驱动、中性色调",
    "palette": ["#1A1A1A", "#FFFFFF", "#0078D4"],
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "aspect_ratio": "16:9",
    "pages": [
        {
            "id": "p1",
            "role": "open",
            "claim": "本季度核心业务增长 25%",
            "evidence": ["收入 1.2 亿元"],
            "layout": "large_white",
            "density": "sparse",
            "emphasis": "high",
            "elements": [
                {"id": "e1", "type": "title", "content": "核心业务增长 25%",
                 "x": 0.5, "y": 0.5, "w": 10, "h": 1.2, "font_size": 36, "color": "#1A1A1A"},
                {"id": "e2", "type": "body", "content": "收入达 1.2 亿元，用户数增长 40%",
                 "x": 0.5, "y": 2.0, "w": 8, "h": 1.0, "font_size": 20, "color": "#333333"},
            ],
        },
        {
            "id": "p2",
            "role": "close",
            "claim": "建议预算增加 15%",
            "layout": "editorial",
            "density": "moderate",
            "emphasis": "medium",
            "elements": [
                {"id": "e3", "type": "title", "content": "预算建议",
                 "x": 0.5, "y": 0.5, "w": 10, "h": 1.2, "font_size": 32, "color": "#1A1A1A"},
            ],
        },
    ],
}


def _write_plan(tmpdir: str, plan: dict) -> str:
    p = os.path.join(tmpdir, "plan.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False)
    return p


# ---------------------------------------------------------------------------
# 原有 20 项回归测试（完整保留）
# ---------------------------------------------------------------------------

def test_required_deck_fields_present():
    for field in REQUIRED_DECK_FIELDS:
        assert field in VALID_PLAN, f"缺少顶层必填字段：{field}"


def test_required_page_fields_present():
    for page in VALID_PLAN["pages"]:
        for field in REQUIRED_PAGE_FIELDS:
            assert field in page, f"页面 {page.get('id')} 缺少 {field}"


def test_required_element_fields_present():
    for page in VALID_PLAN["pages"]:
        for el in page["elements"]:
            for field in REQUIRED_ELEMENT_FIELDS:
                assert field in el, f"元素 {el.get('id')} 缺少 {field}"


def test_valid_plan_compiles_to_pptx(tmp_path):
    _write_plan(str(tmp_path), VALID_PLAN)
    out_path = str(tmp_path / "out.pptx")
    success, errors = compile_plan(VALID_PLAN, out_path, asset_root=str(tmp_path), strict=False)
    assert success, f"合法 plan 应编译成功，但得到错误：{errors}"
    assert Path(out_path).exists(), "输出 PPTX 文件应生成"
    assert Path(out_path).stat().st_size > 0, "PPTX 文件不应为空"


def test_missing_deck_field_is_hard_error():
    bad = dict(VALID_PLAN)
    del bad["core_claim"]
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD]
    assert any("core_claim" in e.location for e in hard), "缺少 core_claim 应产生硬错误"


def test_empty_pages_is_hard_error():
    bad = dict(VALID_PLAN)
    bad["pages"] = []
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD]
    assert any("pages" in e.location for e in hard), "空 pages 应产生硬错误"


def test_element_out_of_bounds_is_hard_error():
    bad = json.loads(json.dumps(VALID_PLAN))
    bad["pages"][0]["elements"][0]["x"] = 100
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert any("E002" in e.code for e in hard), "元素超出边界应产生 E002 硬错误"


def test_element_overlap_is_hard_error():
    bad = json.loads(json.dumps(VALID_PLAN))
    page = bad["pages"][0]
    page["elements"] = [
        {"id": "a", "type": "body", "content": "A", "x": 0, "y": 0, "w": 5, "h": 5,
         "font_size": 20, "color": "#000000"},
        {"id": "b", "type": "body", "content": "B", "x": 2, "y": 2, "w": 5, "h": 5,
         "font_size": 20, "color": "#000000"},
    ]
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert any("overlap" in e.location for e in hard), "元素重叠应产生 E002 硬错误"


def test_missing_element_field_is_hard_error():
    bad = json.loads(json.dumps(VALID_PLAN))
    del bad["pages"][0]["elements"][0]["x"]
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD]
    assert hard, "缺少元素必填字段应产生硬错误"


def test_load_plan_missing_file():
    plan, errors = load_plan("/nonexistent/plan.json")
    assert plan is None
    assert any(e.level == ErrorLevel.HARD for e in errors)


def test_load_plan_invalid_json(tmp_path):
    bad = str(tmp_path / "bad.json")
    with open(bad, "w", encoding="utf-8") as f:
        f.write("{ not json")
    plan, errors = load_plan(bad)
    assert plan is None
    assert any(e.level == ErrorLevel.HARD for e in errors)


def test_deterministic_output(tmp_path):
    p1 = str(tmp_path / "a.pptx")
    p2 = str(tmp_path / "b.pptx")
    compile_plan(VALID_PLAN, p1, asset_root=str(tmp_path))
    compile_plan(VALID_PLAN, p2, asset_root=str(tmp_path))
    s1 = Path(p1).read_bytes()
    s2 = Path(p2).read_bytes()
    assert s1 == s2, "相同 plan 应产出字节一致的 PPTX（确定性）"


def test_chart_missing_data_is_hard_error():
    bad = json.loads(json.dumps(VALID_PLAN))
    bad["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 0, "y": 0, "w": 5, "h": 3,
         "chart": {"kind": "bar"}}
    ]
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E004"]
    assert any("chart" in e.location for e in hard), "图表缺少 data 应产生 E004 硬错误"


def test_4_3_boundary_uses_actual_dimensions():
    plan_43 = json.loads(json.dumps(VALID_PLAN))
    plan_43["aspect_ratio"] = "4:3"
    el = plan_43["pages"][0]["elements"][0]
    el["x"] = 9.0
    el["w"] = 2.0
    # 不显式传参时 validate_plan 也必须自动按 plan["aspect_ratio"] == "4:3" 校验
    errors = validate_plan(plan_43)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert any("E002" in e.code for e in hard), "4:3 页面元素超出 10 英寸宽度应产生 E002"


def test_16_9_boundary_allows_wider_elements():
    plan_169 = json.loads(json.dumps(VALID_PLAN))
    el = plan_169["pages"][0]["elements"][0]
    el["x"] = 9.0
    el["w"] = 2.0
    errors = validate_plan(plan_169)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert not hard, "16:9 页面元素在 13.333 英寸宽度内应合法"


def test_overlap_allowed_is_permitted():
    plan = json.loads(json.dumps(VALID_PLAN))
    page = plan["pages"][0]
    page["elements"] = [
        {"id": "bg", "type": "shape", "content": "", "x": 0, "y": 0, "w": 13.333, "h": 7.5,
         "color": "#F5F1EA", "z_index": 0, "overlap_allowed": True},
        {"id": "fg", "type": "body", "content": "文字叠图", "x": 2, "y": 2, "w": 6, "h": 3,
         "font_size": 20, "color": "#000000", "z_index": 1, "overlap_allowed": True},
    ]
    errors = validate_plan(plan)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert not hard, "overlap_allowed=true 的图层融合不应产生 E002"


def test_overlap_without_allowance_is_hard_error():
    plan = json.loads(json.dumps(VALID_PLAN))
    page = plan["pages"][0]
    page["elements"] = [
        {"id": "a", "type": "body", "content": "A", "x": 0, "y": 0, "w": 5, "h": 5,
         "font_size": 20, "color": "#000000"},
        {"id": "b", "type": "body", "content": "B", "x": 2, "y": 2, "w": 5, "h": 5,
         "font_size": 20, "color": "#000000"},
    ]
    errors = validate_plan(plan)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert any("overlap" in e.location for e in hard), "未设置 overlap_allowed 的实质重叠应产生 E002"


def test_chart_render_failure_records_e007(tmp_path):
    """无需 monkeypatch：categories 与 series 长度失配或非法 kind 必须真实触发 E007 HARD。"""
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 0, "y": 0, "w": 5, "h": 3,
         "chart": {"kind": "bar", "data": {"categories": ["A"], "series": [{"name": "S", "values": [1, 2, 3]}]}}}
    ]
    out_path = str(tmp_path / "chart_fail.pptx")
    success, errors = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=False)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E007"]
    assert not success, "图表数据长度不匹配应阻断编译"
    assert hard, "图表渲染失败应产生真实 E007"


def test_missing_image_asset_records_e003(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"] = [
        {"id": "img1", "type": "image", "image_id": "nonexistent_asset", "x": 0, "y": 0, "w": 3, "h": 2}
    ]
    out_path = str(tmp_path / "img_fail.pptx")
    success, errors = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=False)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E003"]
    assert not success, "素材缺失应阻断编译"
    assert any("E003" in e.code for e in hard), "素材缺失应产生 E003"


def test_success_reflects_hard_errors(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"][0]["x"] = 100
    out_path = str(tmp_path / "fail.pptx")
    success, errors = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=False)
    assert not success, "存在硬错误时 success 应为 False"
    assert any(e.level == ErrorLevel.HARD for e in errors), "应有硬错误记录"


# ---------------------------------------------------------------------------
# 本轮审计新增回归测试（契约对齐、东亚字体、等比裁切、原生表格、重叠漏洞、E008）
# ---------------------------------------------------------------------------

def test_east_asian_font_written_as_valid_ooxml_ea(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["fonts"] = {"heading": "Noto Serif CJK SC", "body": "Noto Sans CJK SC"}
    out_path = str(tmp_path / "fonts.pptx")
    ok, errs = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=True)
    assert ok and not errs
    prs = Presentation(out_path)
    rPr = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]._r.rPr
    assert "typeface" not in rPr.attrib, "ECMA-376 <a:rPr> 不得写入非法 typeface 属性"
    ea = rPr.find(qn("a:ea"))
    latin = rPr.find(qn("a:latin"))
    assert ea is not None and ea.get("typeface") == "Noto Serif CJK SC"
    assert latin is not None and latin.get("typeface") == "Noto Serif CJK SC"


def test_empty_claim_open_question_follows_contract(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["claim"] = ""
    plan["pages"][0]["notes"] = "open question: 原标题为话题标签且缺少比较数据，保持未决"
    ok, errs = compile_plan(plan, str(tmp_path / "open_q.pptx"), strict=True)
    assert ok and not errs, f"按契约留空 claim 并记录 notes 应严格通过，实际：{errs}"

    # 若未写 notes，则给出 SOFT 提示但不硬阻断常规编译
    plan["pages"][0]["notes"] = ""
    soft_errs = validate_plan(plan)
    assert any(e.code == "E001" and e.level == ErrorLevel.SOFT for e in soft_errs)
    assert not any(e.level == ErrorLevel.HARD for e in soft_errs)


def test_zero_negative_size_and_float_edge_bounds():
    bad = json.loads(json.dumps(VALID_PLAN))
    bad["pages"][0]["elements"][0].update({"x": 4.0, "y": 4.0, "w": -1.0, "h": 0.0})
    errs = validate_plan(bad)
    assert any(e.code == "E002" and e.level == ErrorLevel.HARD for e in errs)

    # IEEE-754 1.1 + 12.233 == 13.333000000000002 以及精确 16:9 宽 13.3333 不应误判越界
    good = json.loads(json.dumps(VALID_PLAN))
    good["pages"][0]["elements"] = [
        {"id": "bg", "type": "shape", "x": 0.0, "y": 0.0, "w": 13.3333, "h": 7.5, "color": "#111827"},
        {"id": "t1", "type": "title", "content": "全宽标题", "x": 1.1, "y": 1.0, "w": 12.233, "h": 1.2,
         "font_size": 34, "color": "#FFFFFF"},
    ]
    assert not validate_plan(good)


def test_z_index_zero_and_small_stacked_captions_do_not_bypass_overlap():
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"] = [
        {"id": "a", "type": "body", "content": "A", "x": 1, "y": 1, "w": 5, "h": 3, "z_index": 0},
        {"id": "b", "type": "body", "content": "B", "x": 1, "y": 1, "w": 5, "h": 3, "z_index": 1},
        {"id": "c1", "type": "caption", "content": "注1", "x": 8, "y": 1, "w": 1.2, "h": 0.4},
        {"id": "c2", "type": "caption", "content": "注2", "x": 8, "y": 1, "w": 1.2, "h": 0.4},
    ]
    e002 = [e for e in validate_plan(plan) if e.code == "E002" and e.level == ErrorLevel.HARD]
    assert len(e002) == 2, f"z_index=0 与小尺寸叠字均应被拦截，实际：{e002}"


def test_image_preserves_aspect_ratio_via_crop_and_accepts_extension(tmp_path):
    img = Image.new("RGB", (1600, 900), (200, 180, 160))
    ImageDraw.Draw(img).ellipse((500, 150, 1100, 750), fill=(120, 40, 40))
    img.save(tmp_path / "hero.png")

    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"] = [
        {"id": "pic", "type": "image", "image_id": "hero.png", "x": 7.5, "y": 0.5, "w": 5.0, "h": 6.5}
    ]
    out_path = str(tmp_path / "img_crop.pptx")
    ok, errs = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=True)
    assert ok and not errs
    prs = Presentation(out_path)
    pic = prs.slides[0].shapes[0]
    assert pic.crop_left > 0.25 and pic.crop_right > 0.25
    effective_ratio = (1600 / 900) * (1.0 - pic.crop_left - pic.crop_right)
    assert abs(effective_ratio - (5.0 / 6.5)) < 1e-3


def test_table_kind_renders_real_editable_table_and_rejects_unknown_kind(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["claim"] = "Q3 营收达 1.2 亿元，复购率升至 42%"
    plan["pages"][0]["elements"] = [
        {
            "id": "tbl1", "type": "chart", "x": 0.8, "y": 1.2, "w": 11.5, "h": 4.5,
            "chart": {
                "kind": "table",
                "title": "Q3 关键运营指标",
                "unit": "多口径",
                "period": "2026 Q3",
                "data": {
                    "headers": ["核心指标", "数值", "同比/变化"],
                    "rows": [["季度营收", "1.2 亿元", "+25%"], ["客户复购率", "42%", "+6 pct"]],
                },
            },
        }
    ]
    out_path = str(tmp_path / "table_ok.pptx")
    ok, errs = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=True)
    assert ok and not errs, f"原生表格编译不应报错：{errs}"
    prs = Presentation(out_path)
    assert any(sh.has_table for sh in prs.slides[0].shapes), "kind=table 必须生成原生可编辑表格而非柱状图"

    # 非法 chart.kind 必须报 E007 HARD，不得静默画成柱状图
    plan["pages"][0]["elements"][0]["chart"]["kind"] = "radar"
    ok_bad, errs_bad = compile_plan(plan, str(tmp_path / "bad_kind.pptx"), asset_root=str(tmp_path))
    assert not ok_bad and any(e.code == "E007" and e.level == ErrorLevel.HARD for e in errs_bad)


def test_e008_catches_conflict_without_penalizing_derived_rates_or_commas():
    assert _extract_numbers("ARR 已破 1,000 万，Q3 同比增长 25%") == [1000.0, 25.0]

    # Case 5 原稿冲突：800 -> 900 -> 1000（+25%），但主张写“下降 30%” → 必须报 E008
    conflict_plan = json.loads(json.dumps(VALID_PLAN))
    conflict_plan["pages"][0]["claim"] = "本季度获客成本下降 30%，获客效率大幅提升"
    conflict_plan["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 1, "y": 1.5, "w": 8, "h": 4.5,
         "chart": {"kind": "bar", "data": {"categories": ["Q1", "Q2", "Q3"],
                                           "series": [{"name": "获客人数", "values": [800, 900, 1000]}]}}}
    ]
    assert any(e.code == "E008" for e in validate_plan(conflict_plan))

    # 改写为真实派生结论“Q1–Q3 获客增长 25%” + 13pt 来源注释 → 零警告通过
    fixed_plan = json.loads(json.dumps(conflict_plan))
    fixed_plan["pages"][0]["claim"] = "Q1–Q3 获客人数持续增长 25%"
    fixed_plan["pages"][0]["elements"].append(
        {"id": "src", "type": "caption", "content": "数据来源：增长数据仓库（2026 Q1–Q3）",
         "x": 1, "y": 6.4, "w": 8, "h": 0.5, "font_size": 13}
    )
    assert validate_plan(fixed_plan) == []

    # 多系列按类别纵向合计（118+86=204）与占比（118/204≈58%）不应误报 E008
    multi_plan = json.loads(json.dumps(conflict_plan))
    multi_plan["pages"][0]["claim"] = "Q4 单季营收达 204 万元，会员复购与礼赠占比升至 58%"
    multi_plan["pages"][0]["elements"][0]["chart"]["data"] = {
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            {"name": "礼赠", "values": [52, 78, 105, 118]},
            {"name": "堂食", "values": [74, 78, 78, 86]},
        ],
    }
    assert validate_plan(multi_plan) == []


def test_structure_over_containers_table_and_dark_chart(tmp_path):
    """验证结构优先于组件（场域 > 对齐 > 发丝线 > 字体层级 > 容器）：无竖框三线表、透明场域与深色背景自适应。"""
    plan = json.loads(json.dumps(VALID_PLAN))
    # 深色页专用调色板：首个可用色必须真的可辨，编译器不应静默换色（换色时另有 E009 披露，见下）
    plan["palette"] = ["#C88D67", "#8C8278"]
    plan["pages"] = [
        {
            "id": "p_light_tbl",
            "role": "build",
            "claim": "Q3 营收达 1.2 亿元，复购率 42%",
            "background": "#F5F1EB",
            "elements": [
                {
                    "id": "tbl_clean",
                    "type": "chart",
                    "x": 0.8,
                    "y": 1.2,
                    "w": 11.6,
                    "h": 4.2,
                    "chart": {
                        "kind": "table",
                        "title": "核心经营指标",
                        "unit": "多口径",
                        "period": "2026 Q3",
                        "data": {
                            "headers": ["指标", "数值", "变动"],
                            "rows": [["营收", "1.2 亿元", "+25%"], ["复购率", "42%", "+6 pct"]],
                        },
                    },
                }
            ],
        },
        {
            "id": "p_dark_chart",
            "role": "turn",
            "claim": "Q3 销量升至 150 万件",
            "background": "#151210",
            "elements": [
                {
                    "id": "bar_dark",
                    "type": "chart",
                    "x": 0.8,
                    "y": 1.2,
                    "w": 11.6,
                    "h": 4.5,
                    "chart": {
                        "kind": "bar",
                        "title": "季度销量趋势",
                        "unit": "万件",
                        "data": {
                            "categories": ["Q1", "Q2", "Q3"],
                            "series": [{"name": "销量", "values": [100, 120, 150]}],
                        },
                    },
                }
            ],
        },
    ]
    out_path = str(tmp_path / "structural_clean.pptx")
    ok, errs = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=True)
    assert ok and not errs, f"编译应零警告通过：{errs}"

    prs = Presentation(out_path)
    tbl = [sh.table for sh in prs.slides[0].shapes if sh.has_table][0]
    tcPr = tbl.cell(0, 0)._tc.tcPr
    children_tags = [el.tag.split("}")[-1] for el in tcPr]
    # 1) OOXML XSD 顺序：lnL/lnR/lnT/lnB 必须位于 noFill 之前
    assert children_tags.index("lnB") < children_tags.index("noFill")
    # 2) 左右竖向边框为 noFill，无封闭网格笼子；单元格背景为 noFill 融入页面场域
    assert tcPr.find(qn("a:lnL")).find(qn("a:noFill")) is not None
    assert tcPr.find(qn("a:lnR")).find(qn("a:noFill")) is not None
    assert tcPr.find(qn("a:lnB")).find(qn("a:solidFill")) is not None
    # 3) 深色页柱状图消费调色板强调色 #C88D67，且标题为浅色字
    dark_chart = [sh.chart for sh in prs.slides[1].shapes if sh.has_chart][0]
    title_rgb = str(dark_chart.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb)
    assert title_rgb == "F3F4F6"
    assert str(dark_chart.series[0].format.fill.fore_color.rgb) == "C88D67"


def test_cli_main_exit_codes(tmp_path):
    plan_file = _write_plan(str(tmp_path), VALID_PLAN)
    out_file = str(tmp_path / "nested" / "cli_out.pptx")
    assert main([plan_file, out_file]) == 0
    assert Path(out_file).is_file()


if __name__ == "__main__":
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))


# ---------------------------------------------------------------------------
# P1 审计：E008 主张—数据语义核对（方向 / 单位 / 期间 / 派生值 / 逐项判定）
# ---------------------------------------------------------------------------

def _e008(claim, chart, page_extra=None):
    """构造单图表页面并返回 E008 结论（None 表示未触发）。"""
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"] = [{
        "id": "p_sem", "role": "build", "claim": claim,
        "elements": [{"id": "c1", "type": "chart", "x": 1, "y": 1, "w": 8, "h": 5, "chart": chart}],
        **(page_extra or {}),
    }]
    hits = [e for e in validate_plan(plan) if e.code == "E008"]
    assert len(hits) <= 1, f"同一页面应聚合为单条 E008 告警，实际 {len(hits)} 条"
    return hits[0] if hits else None


BAR_800_900_1000 = {
    "kind": "bar", "unit": "人",
    "data": {"categories": ["Q1", "Q2", "Q3"],
             "series": [{"name": "获客人数", "values": [800, 900, 1000]}]},
}


def test_e008_direction_mismatch_is_reported_not_passed():
    """数据上升 25%，主张写"下降 25%"：方向相反，数值碰巧一致也必须报出。"""
    err = _e008("获客人数下降 25%", BAR_800_900_1000)
    assert err is not None, "方向相反的主张不得因数值一致而通过"
    assert "direction" in err.actual
    assert err.level == ErrorLevel.SOFT, "语义核对是 SOFT 复核提示，不得擅自阻断设计判断"


def test_e008_single_matching_number_does_not_clear_whole_claim():
    """主张含两个定量表述：900 可核对，30% 无来源 → 仍须报出，且逐项列出。"""
    err = _e008("Q2 获客 900 人，同时获客成本下降 30%", BAR_800_900_1000)
    assert err is not None, "任一数字匹配不得使整条主张通过"
    assert "可核对 1 项" in err.actual and "待核 1 项" in err.actual
    assert "30%" in err.actual and "900" not in err.actual.split("待核")[1]


def test_e008_unit_scale_mismatch_is_reported():
    """同一数字不同量级（万元 vs 亿元）是量纲错配，不得视为对应。"""
    chart = {"kind": "bar", "unit": "万元",
             "data": {"categories": ["Q1"], "series": [{"name": "营收", "values": [1.2]}]}}
    err = _e008("本季度收入 1.2 亿元", chart)
    assert err is not None and "unit" in err.actual


def test_e008_unit_scale_conversion_is_accepted():
    """真实换算（12000 万元 = 1.2 亿元；万元/亿元互转）应正确通过，不产生噪声。"""
    chart = {"kind": "bar", "unit": "万元",
             "data": {"categories": ["Q1", "Q2", "Q3"],
                      "series": [{"name": "营收", "values": [9000, 10500, 12000]}]}}
    assert _e008("Q3 营收 12000 万元", chart) is None
    assert _e008("Q3 营收 1.2 亿元", chart) is None


def test_e008_period_mismatch_is_reported_even_when_value_matches():
    """Q4 主张引用了 Q3 的数值：数值一致但期间不符，必须提示复核。"""
    err = _e008("Q4 获客达 1000 人", BAR_800_900_1000)
    assert err is not None and "period" in err.actual


def test_e008_cross_period_range_is_accepted():
    """Q1–Q3 跨期累计主张不应因"匹配到 Q3"而误报。"""
    assert _e008("Q1–Q3 获客人数累计增长 25%", BAR_800_900_1000) is None


def test_e008_derived_magnitude_pollution_does_not_pass():
    """不得靠 ×100 之类的派生候选让任意数字"碰巧匹配"。"""
    assert _e008("规模达到 80,000 人", BAR_800_900_1000) is not None


def test_e008_fraction_vs_percent_scale_is_flagged_for_review():
    """数据以 0.25 存储而主张写 25%：尺度歧义应提示复核而非静默通过。"""
    chart = {"kind": "bar", "unit": "",
             "data": {"categories": ["Q1"], "series": [{"name": "复购率", "values": [0.25]}]}}
    err = _e008("复购率 25%", chart)
    assert err is not None and "assumption" in err.actual


def test_e008_percentage_point_difference_is_not_a_false_positive():
    """百分比数据的点差是百分点，不应被当成缺失的数值。"""
    chart = {"kind": "bar", "unit": "%",
             "data": {"categories": ["7日", "30日"], "series": [{"name": "留存", "values": [45, 28]}]}}
    assert _e008("留存下降 17 个百分点", chart) is None
    # 留存窗口（7 日/30 日）是修饰语而非数据量，同样不得误报
    assert _e008("7 日留存 45%，30 日留存 28%", chart) is None


def test_e008_from_to_order_is_checked_against_series():
    """「从 1000 降至 800」与上升序列冲突：先后顺序与方向都须核对。"""
    err = _e008("获客从 1000 降至 800 人", BAR_800_900_1000)
    assert err is not None and "direction" in err.actual


def test_e008_aggregates_per_page_and_never_blocks_compile(tmp_path):
    """E008 是复核提示：同页只报一条，且不改变 success（设计判断权仍在 AI/人）。"""
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["claim"] = "获客人数下降 25%，成本下降 30%"
    plan["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 0.6, "y": 1.2, "w": 6.0, "h": 4.5, "chart": BAR_800_900_1000},
    ]
    out = str(tmp_path / "soft.pptx")
    ok, errs = compile_plan(plan, out, asset_root=str(tmp_path), strict=False)
    assert ok, "E008 为 SOFT，常规编译不应被阻断"
    assert sum(1 for e in errs if e.code == "E008") == 1
    ok_strict, _ = compile_plan(plan, str(tmp_path / "soft_strict.pptx"), asset_root=str(tmp_path), strict=True)
    assert not ok_strict, "--strict 下 SOFT 必须阻断，不得被吞掉"


# ---------------------------------------------------------------------------
# P2 审计：图表与色彩契约（饼/环逐扇区着色、深底换色披露、图例与标注纪律）
# ---------------------------------------------------------------------------

def _chart_of(plan, path, tmp_path, page_idx=0):
    out = str(tmp_path / "chart_contract.pptx")
    ok, errs = compile_plan(plan, out, asset_root=str(tmp_path), strict=False)
    prs = Presentation(out)
    return ok, errs, [sh.chart for sh in prs.slides[page_idx].shapes if getattr(sh, "has_chart", False) and sh.has_chart][0]


def _single_chart_plan(chart, palette, bg=None):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["palette"] = palette
    page = {"id": "p_chart", "role": "build", "claim": "", "notes": "图表契约测试",
            "elements": [{"id": "c1", "type": "chart", "x": 1, "y": 1, "w": 8, "h": 5, "chart": chart}]}
    if bg:
        page["background"] = bg
    plan["pages"] = [page]
    return plan


def test_pie_and_donut_color_every_sector_not_only_first_series(tmp_path):
    """饼/环图类别靠扇区颜色区分：只给 series[0].format 上色等于整图单色。"""
    for kind in ("pie", "donut"):
        plan = _single_chart_plan(
            {"kind": kind, "title": "结构", "data": {"categories": ["A", "B", "C"],
                                                     "series": [{"name": "占比", "values": [50, 30, 20]}]}},
            ["#B08968", "#6B7A6C", "#1F3A5F"],
        )
        ok, errs, chart = _chart_of(plan, plan["pages"][0]["elements"][0]["chart"], tmp_path)
        assert ok, errs
        pts = chart.plots[0].series[0].points
        fills = [str(p.format.fill.fore_color.rgb) for p in pts]
        assert fills == ["B08968", "6B7A6C", "1F3A5F"], f"{kind} 每个扇区都应消费调色板：{fills}"


def test_pie_labels_show_percentage_only(tmp_path):
    """数值标注不得同时输出原值与百分比（渲染后会出现「58; 58%」这类重复）。"""
    plan = _single_chart_plan(
        {"kind": "pie", "data": {"categories": ["A", "B"], "series": [{"name": "占比", "values": [58, 42]}]}},
        ["#B08968", "#6B7A6C"],
    )
    ok, _, chart = _chart_of(plan, None, tmp_path)
    dl = chart.plots[0].data_labels
    assert dl.show_percentage is True
    assert dl.show_value in (False, None), "饼图不应同时显示原值"


def test_invisible_palette_color_is_disclosed_not_silently_substituted(tmp_path):
    """与底色同色的系列色被跳过时，必须给出 E009 披露，而不是悄悄换色。"""
    plan = _single_chart_plan(
        {"kind": "bar", "unit": "万元", "data": {"categories": ["Q1", "Q2"],
                                                 "series": [{"name": "营收", "values": [1, 2]}]}},
        ["#14110F", "#C88D67"], bg="#14110F",
    )
    ok, errs, chart = _chart_of(plan, None, tmp_path)
    assert ok, "换色不应阻断编译"
    e009 = [e for e in errs if e.code == "E009"]
    assert e009 and "#14110F" in e009[0].actual, "不可辨色必须显式披露"
    assert str(chart.series[0].format.fill.fore_color.rgb) == "C88D67"


def test_trailing_neutral_palette_colors_do_not_raise_noise(tmp_path):
    """调色板尾部的中性色（发丝线/底色）未被系列消费时不应产生告警。"""
    plan = _single_chart_plan(
        {"kind": "bar", "unit": "万元", "data": {"categories": ["Q1"], "series": [{"name": "营收", "values": [1]}]}},
        ["#B08968", "#111827", "#E5E7EB"], bg="#FFFFFF",
    )
    ok, errs, _ = _chart_of(plan, None, tmp_path)
    assert ok and not errs, f"未被消费的浅灰不应触发告警：{errs}"


def test_single_series_chart_hides_redundant_legend_but_pie_keeps_it(tmp_path):
    ok, _, bar = _chart_of(_single_chart_plan(
        {"kind": "bar", "data": {"categories": ["Q1"], "series": [{"name": "营收", "values": [1]}]}},
        ["#B08968"]), None, tmp_path)
    assert bar.has_legend is False, "单系列柱图图例与标题重复，应省略"
    ok, _, pie = _chart_of(_single_chart_plan(
        {"kind": "pie", "data": {"categories": ["A", "B"], "series": [{"name": "占比", "values": [50, 50]}]}},
        ["#B08968", "#6B7A6C"]), None, tmp_path)
    assert pie.has_legend is True, "饼图靠颜色区分类别，必须保留图例"


def test_brand_yellow_on_white_is_rejected_with_disclosure(tmp_path):
    """白底亮黄 #F5C518 作为数据填充对比度仅 1.63:1：应跳过并披露，不得静默使用。"""
    plan = _single_chart_plan(
        {"kind": "bar", "data": {"categories": ["Q1", "Q2"],
                                 "series": [{"name": "A", "values": [1, 2]}, {"name": "B", "values": [2, 1]}]}},
        ["#F5C518", "#111827"], bg="#FFFFFF",
    )
    ok, errs, chart = _chart_of(plan, None, tmp_path)
    assert ok
    e009 = [e for e in errs if e.code == "E009"]
    assert e009 and "F5C518" in e009[0].actual
    assert str(chart.series[0].format.fill.fore_color.rgb) == "111827"


def test_e008_flags_caption_period_that_contradicts_chart_period():
    """图注/来源期间与图表期间完全无交集时提示复核；有交集（如对比去年同期）不误报。"""
    chart = {"kind": "bar", "unit": "亿元", "period": "2025 Q2–2026 Q3",
             "data": {"categories": ["25Q2", "26Q3"],
                      "series": [{"name": "收入", "values": [0.52, 1.2]}]}}
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"] = [{
        "id": "p_cap", "role": "build", "claim": "26Q3 收入 1.2 亿元",
        "elements": [
            {"id": "c1", "type": "chart", "x": 1, "y": 1, "w": 8, "h": 4, "chart": chart},
            {"id": "src", "type": "caption", "content": "数据来源：财务系统 2024 Q1–Q4。",
             "x": 1, "y": 5.4, "w": 8, "h": 0.4, "font_size": 13},
        ],
    }]
    errs = [e for e in validate_plan(plan) if e.code == "E008"]
    assert errs and "src" in errs[0].actual and "无交集" in errs[0].actual

    # 与图表期间存在交集的图注（含"对比去年同期"式写法）不得误报
    for content in ("数据来源：财务系统 2025 Q2–2026 Q3，含去年同期对照。",
                    "数据来源：财务系统 2026 Q1–Q3。"):
        plan["pages"][0]["elements"][1]["content"] = content
        assert not [e for e in validate_plan(plan) if e.code == "E008"], f"不应误报：{content}"


def test_e007_reports_actionable_root_cause(tmp_path):
    """图表失败必须给出可定位原因，不能只写『数据非法』让 AI 猜。"""
    cases = [
        ({"kind": "radar", "data": {"categories": ["A"], "series": [{"name": "s", "values": [1]}]}}, "kind='radar'"),
        ({"kind": "bar", "data": {"categories": ["A"], "series": [{"name": "s", "values": [1, 2]}]}}, "长度 2"),
        ({"kind": "bar", "data": {"categories": ["A"], "series": [{"name": "s", "values": ["x"]}]}}, "不可解析"),
    ]
    for chart, needle in cases:
        plan = json.loads(json.dumps(VALID_PLAN))
        plan["pages"] = [{"id": "p", "role": "build", "claim": "x", "elements": [
            {"id": "c1", "type": "chart", "x": 1, "y": 1, "w": 8, "h": 5, "chart": chart}]}]
        ok, errs = compile_plan(plan, str(tmp_path / "bad.pptx"), asset_root=str(tmp_path))
        assert not ok
        e007 = [e for e in errs if e.code == "E007"]
        assert e007 and needle in e007[0].actual, f"E007 未给出根因：{e007}"


# ---------------------------------------------------------------------------
# P1 审计：文本装载能力提示（E010，SOFT，不阻断也不替人改文案）
# ---------------------------------------------------------------------------

def test_e010_flags_text_that_cannot_fit_its_box():
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"].append(
        {"id": "crowded", "type": "body", "content": "这是一段明显超出文本框容量的长正文。" * 12,
         "x": 0.5, "y": 4.2, "w": 5.0, "h": 1.0, "font_size": 20}
    )
    errs = validate_plan(plan)
    e010 = [e for e in errs if e.code == "E010"]
    assert e010 and e010[0].level == ErrorLevel.SOFT
    assert "估算需" in e010[0].actual and "勿压缩字号" in e010[0].message


def test_e010_stays_quiet_for_text_that_fits():
    plan = json.loads(json.dumps(VALID_PLAN))
    assert not [e for e in validate_plan(plan) if e.code == "E010"]
    # 显式换行且余量充足时也不应误报
    plan["pages"][0]["elements"].append(
        {"id": "ok", "type": "body", "content": "第一行\n第二行",
         "x": 0.5, "y": 4.2, "w": 6.0, "h": 1.6, "font_size": 18}
    )
    assert not [e for e in validate_plan(plan) if e.code == "E010"]


def test_e010_does_not_shrink_font_or_rewrite_content(tmp_path):
    """提示归提示：编译产物必须保持作者给定字号与原文，不得自动降级。"""
    plan = json.loads(json.dumps(VALID_PLAN))
    long_text = "排版不得靠缩字号硬塞原文。" * 10
    plan["pages"][0]["elements"] = [
        {"id": "t", "type": "body", "content": long_text, "x": 0.5, "y": 0.5,
         "w": 4.0, "h": 0.9, "font_size": 21, "color": "#111111"}
    ]
    out = str(tmp_path / "no_autoshrink.pptx")
    ok, errs = compile_plan(plan, out, asset_root=str(tmp_path))
    assert ok, "E010 为 SOFT，不应阻断"
    prs = Presentation(out)
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert run.font.size.pt == 21, "不得擅自缩小字号"
    assert run.text == long_text.split("\n")[0], "不得擅自改写文案"


def test_e008_bare_magnitude_unit_matches_scaled_unit():
    """「1,000 万」与图表 unit=万元 的 1000 是同一量级，不应误报。"""
    chart = {"kind": "bar", "unit": "万元",
             "data": {"categories": ["Q3"], "series": [{"name": "ARR", "values": [1000]}]}}
    assert _e008("ARR 已破 1,000 万", chart) is None
    # 但量级差 100 倍（万 vs 亿）必须报出
    assert _e008("ARR 已破 1,000 亿元", chart) is not None


def test_pie_label_ink_maximises_worst_case_contrast(tmp_path):
    """扇区深浅不一时，标注墨色必须按"最差扇区也可读"选择，不能固定深色。"""
    from scripts.compile_pptx import _contrast_ratio, _label_ink

    fills = ["#B08968", "#6B7A6C", "#1F3A5F"]
    ink = _label_ink(fills, "#FFFFFF")
    worst = min(_contrast_ratio(ink, f) for f in fills)
    for alt in ("#111827", "#FFFFFF"):
        assert worst >= min(_contrast_ratio(alt, f) for f in fills) - 1e-9

    plan = _single_chart_plan(
        {"kind": "pie", "data": {"categories": ["A", "B", "C"],
                                 "series": [{"name": "占比", "values": [45, 30, 25]}]}},
        fills, bg="#FFFFFF",
    )
    ok, _, chart = _chart_of(plan, None, tmp_path)
    assert ok
    assert str(chart.plots[0].data_labels.font.color.rgb) == ink.lstrip("#")
