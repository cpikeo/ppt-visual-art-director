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
    CompileError,
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
    plan["palette"] = ["#151210", "#C88D67", "#8C8278"]
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
    # 3) 深色页柱状图自动跳过与暗底撞色的 #151210，选用高对比强调色 #C88D67，且标题为浅色字
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
