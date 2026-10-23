#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ppt-visual-art-director · 程序正确性测试

只验证实际存在的程序行为与契约（compile_pptx.py），不用于替代设计质量判断。
覆盖：数据契约、合法/非法输入处理、确定性计算、边界条件、异常路径、PPTX 编译与文件完整性。
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

# 保证能直接从仓库根运行 pytest 并导入 scripts
REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from scripts.compile_pptx import (  # noqa: E402
    load_plan,
    validate_plan,
    compile_plan,
    CompileError,
    ErrorLevel,
    REQUIRED_DECK_FIELDS,
    REQUIRED_PAGE_FIELDS,
    REQUIRED_ELEMENT_FIELDS,
)

# ---------------------------------------------------------------------------
# 测试用 plan
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
# 契约与合法输入
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
    plan_path = _write_plan(str(tmp_path), VALID_PLAN)
    out_path = str(tmp_path / "out.pptx")
    success, errors = compile_plan(VALID_PLAN, out_path, asset_root=str(tmp_path), strict=False)
    assert success, f"合法 plan 应编译成功，但得到错误：{errors}"
    assert Path(out_path).exists(), "输出 PPTX 文件应生成"
    assert Path(out_path).stat().st_size > 0, "PPTX 文件不应为空"


# ---------------------------------------------------------------------------
# 非法输入处理
# ---------------------------------------------------------------------------

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
    bad["pages"][0]["elements"][0]["x"] = 100  # 超出边界
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
    with open(bad, "w") as f:
        f.write("{ not json")
    plan, errors = load_plan(bad)
    assert plan is None
    assert any(e.level == ErrorLevel.HARD for e in errors)


# ---------------------------------------------------------------------------
# 确定性：同一输入产出相同结果
# ---------------------------------------------------------------------------

def test_deterministic_output(tmp_path):
    p1 = str(tmp_path / "a.pptx")
    p2 = str(tmp_path / "b.pptx")
    compile_plan(VALID_PLAN, p1, asset_root=str(tmp_path))
    compile_plan(VALID_PLAN, p2, asset_root=str(tmp_path))
    s1 = Path(p1).read_bytes()
    s2 = Path(p2).read_bytes()
    assert s1 == s2, "相同 plan 应产出字节一致的 PPTX（确定性）"


# ---------------------------------------------------------------------------
# 图表数据缺失 → 硬错误
# ---------------------------------------------------------------------------

def test_chart_missing_data_is_hard_error():
    bad = json.loads(json.dumps(VALID_PLAN))
    bad["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 0, "y": 0, "w": 5, "h": 3,
         "chart": {"kind": "bar"}}
    ]
    errors = validate_plan(bad)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E004"]
    assert any("chart" in e.location for e in hard), "图表缺少 data 应产生 E004 硬错误"


# ---------------------------------------------------------------------------
# 页面比例：4:3 边界校验使用实际尺寸（非 16:9 硬编码）
# ---------------------------------------------------------------------------


def test_4_3_boundary_uses_actual_dimensions():
    """4:3 页面宽 10 英寸，x=9,w=2（x2=11）在 16:9 下合法，在 4:3 下应阻断。"""
    plan_43 = json.loads(json.dumps(VALID_PLAN))
    plan_43["aspect_ratio"] = "4:3"
    el = plan_43["pages"][0]["elements"][0]
    el["x"] = 9.0
    el["w"] = 2.0  # x2 = 11.0 > 10.0
    errors = validate_plan(plan_43, slide_w_in=10.0, slide_h_in=7.5)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert any("E002" in e.code for e in hard), "4:3 页面元素超出 10 英寸宽度应产生 E002"


def test_16_9_boundary_allows_wider_elements():
    """16:9 页面宽 13.333 英寸，x=9,w=2（x2=11）应合法。"""
    plan_169 = json.loads(json.dumps(VALID_PLAN))
    el = plan_169["pages"][0]["elements"][0]
    el["x"] = 9.0
    el["w"] = 2.0  # x2 = 11.0 < 13.333
    errors = validate_plan(plan_169, slide_w_in=13.333, slide_h_in=7.5)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E002"]
    assert not hard, "16:9 页面元素在 13.333 英寸宽度内应合法"


# ---------------------------------------------------------------------------
# 合法重叠：overlap_allowed=true 视为设计授权图层融合，不应阻断
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# 图表渲染失败：E007 并阻断
# ---------------------------------------------------------------------------


def test_chart_render_failure_records_e007(tmp_path, monkeypatch):
    """_add_chart 返回 False 时，应产生 E007 并阻断编译。"""
    import scripts.compile_pptx as m

    def _fake_add_chart(*args, **kwargs):
        return False

    monkeypatch.setattr(m, "_add_chart", _fake_add_chart)

    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"] = [
        {"id": "c1", "type": "chart", "x": 0, "y": 0, "w": 5, "h": 3,
         "chart": {"kind": "bar", "data": {"categories": ["A"], "series": [{"name": "S", "values": [1, 2, 3]}]}}}
    ]
    out_path = str(tmp_path / "chart_fail.pptx")
    success, errors = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=False)
    hard = [e for e in errors if e.level == ErrorLevel.HARD and e.code == "E007"]
    assert not success, "图表渲染失败应阻断编译"
    assert any("E007" in e.code for e in hard), "图表渲染失败应产生 E007"


# ---------------------------------------------------------------------------
# 素材缺失：E003 并阻断（编译阶段）
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# 错误状态：success 正确反映硬错误
# ---------------------------------------------------------------------------


def test_success_reflects_hard_errors(tmp_path):
    plan = json.loads(json.dumps(VALID_PLAN))
    plan["pages"][0]["elements"][0]["x"] = 100  # 超出边界
    out_path = str(tmp_path / "fail.pptx")
    success, errors = compile_plan(plan, out_path, asset_root=str(tmp_path), strict=False)
    assert not success, "存在硬错误时 success 应为 False"
    assert any(e.level == ErrorLevel.HARD for e in errors), "应有硬错误记录"


if __name__ == "__main__":
    # 允许直接 python dev/tests/test_compile.py 快速自检
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))
