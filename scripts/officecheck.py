#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Office Validation Engine（V23）：OPC / 关系 / ID / XSD / 图表 / 动画引用完整性，一处实现。

层级（只说结构层，不代表 PowerPoint 实机结果；LibreOffice 渲染不在此处，见 verify.py）：
  OPC      重复部件名、[Content_Types] 覆盖与悬空 Override、关系目标存在、关系 Id 唯一、XML 可解析
  引用     sldIdLst ↔ presentation.xml.rels、未登记幻灯片、同页 cNvPr@id 唯一
  独占     图表 / 内嵌工作簿 / 备注页 / tags / SmartArt 数据 / OLE 只能有一个属主（复制页不得共享）
  XSD      ISO/IEC 29500-4 的 pml.xsd（幻灯片、版式、母版、备注、presentation.xml）；mc:AlternateContent 先按 Fallback 展开
  图表     post.chart_problems（VAD 既有图表兼容性自检）
  动画     motion.integrity_problems（悬空目标、cTn 重复 id、click 锚点、段落/系列越界、子节点顺序）
  基线     --baseline 原稿：与原稿字节相同的部件不重新校验；原稿已有的 XSD 缺陷不算新增（增量验证）

借鉴 anthropics/skills pptx 校验器的检查项（唯一 id、关系、独占、基线对比、MC 预处理）；实现为独立重写，见 NOTICE。
退出码：0 PASS｜1 BLOCK｜2 输入错误。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from lxml import etree

import motion
from model import LIMITS, package_problems, parse_xml, utf8_stdio

HERE = Path(__file__).resolve().parent
SCHEMA_DIR = HERE.parent / "schemas" / "ooxml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
R_NS = motion.R_NS
MC_NS = motion.MC_NS
XSD_PARTS = re.compile(r"ppt/(slides|slideLayouts|slideMasters|notesSlides|notesMasters)/[^/]+\.xml|ppt/presentation\.xml")
# 独占属主：这些部件被多页引用意味着复制页时没有深拷贝，改一页会同步改另一页
EXCLUSIVE = ("/chart", "/package", "/notesSlide", "/tags", "/diagramData", "/diagramLayout", "/diagramQuickStyle", "/diagramColors",
             "/diagramDrawing", "/oleObject", "/chartUserShapes", "/comments", "/themeOverride")
_SCHEMA = None


def _schema() -> etree.XMLSchema:
    global _SCHEMA
    if _SCHEMA is None:
        _SCHEMA = etree.XMLSchema(etree.parse(str(SCHEMA_DIR / "pml.xsd")))
    return _SCHEMA


def resolve_target(source: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def rels_name(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def read_parts(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist() if not i.filename.endswith("/")}


def part_rels(parts: Dict[str, bytes], part: str) -> List[Dict[str, str]]:
    name = rels_name(part)
    if name not in parts:
        return []
    out = []
    for r in parse_xml(parts[name]):
        if not isinstance(r.tag, str):
            continue
        ext = (r.get("TargetMode") or "").lower() == "external"
        out.append({"id": r.get("Id"), "type": r.get("Type"), "target": r.get("Target"), "external": ext,
                    "part": None if ext else resolve_target(part, r.get("Target"))})
    return out


def _mc_expand(root) -> None:
    """XSD 前的 MC 预处理：AlternateContent 以 Fallback 展开（无 Fallback 则移除）；去掉可忽略命名空间的属性与元素。"""
    for ac in list(root.iter("{%s}AlternateContent" % MC_NS)):
        parent = ac.getparent()
        fb = ac.find("{%s}Fallback" % MC_NS)
        pos = parent.index(ac)
        for k, child in enumerate(list(fb) if fb is not None else []):
            parent.insert(pos + k, child)
        parent.remove(ac)
    ign = root.get("{%s}Ignorable" % MC_NS)
    if ign:
        nss = {root.nsmap.get(p) for p in ign.split()} - {None}
        for el in list(root.iter()):
            if isinstance(el.tag, str) and etree.QName(el).namespace in nss and el is not root:
                el.getparent().remove(el)
        for el in root.iter():
            for a in list(el.attrib):
                if a.startswith("{") and a[1:].split("}")[0] in nss:
                    del el.attrib[a]
        del root.attrib["{%s}Ignorable" % MC_NS]


def xsd_errors(data: bytes) -> List[str]:
    root = parse_xml(data)
    _mc_expand(root)
    sch = _schema()
    if sch.validate(etree.ElementTree(root)):
        return []
    return sorted({re.sub(r"^:\d+:\d+:", "", f":{e.line}:{e.column}:{e.message}") for e in sch.error_log})


def _chart_info(parts: Dict[str, bytes], slide: str, rels: List[Dict[str, str]]):
    by_id = {r["id"]: r["part"] for r in rels}
    plots = ("barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart")

    def info(rid: str):
        part = by_id.get(rid)
        if part not in parts:
            return 0, 0, None
        root = parse_xml(parts[part])
        ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
        sers = root.findall(".//c:ser", ns)
        plot = next((etree.QName(e).localname for e in root.iter() if isinstance(e.tag, str) and etree.QName(e).localname in plots), None)
        cats = sers[0].findall(".//c:cat//c:pt", ns) if sers else []
        return len(sers), len(cats), plot
    return info


def listed_slides(parts: Dict[str, bytes]) -> List[Dict[str, Any]]:
    pres = parse_xml(parts["ppt/presentation.xml"])
    rels = {r["id"]: r for r in part_rels(parts, "ppt/presentation.xml")}
    out = []
    for e in pres.iter(motion.pn("sldId")):
        r = rels.get(e.get("{%s}id" % R_NS))
        out.append({"sld_id": e.get("id"), "rid": e.get("{%s}id" % R_NS), "part": r["part"] if r else None})
    return out


def check_parts(parts: Dict[str, bytes], baseline: Optional[Dict[str, bytes]] = None, xsd: bool = True) -> Dict[str, Any]:
    """校验一个已读入的包。返回 {status, problems:[{code,where,message}], stats}。"""
    probs: List[Dict[str, str]] = []

    def bad(code: str, where: str, msg: str) -> None:
        probs.append({"code": code, "where": where, "message": msg})

    # ---- OPC
    ct = parse_xml(parts["[Content_Types].xml"]) if "[Content_Types].xml" in parts else None
    if ct is None:
        return {"status": "BLOCK", "problems": [{"code": "E-OPC", "where": "[Content_Types].xml", "message": "缺少 [Content_Types].xml"}], "stats": {}}
    defaults = {e.get("Extension", "").lower() for e in ct.findall("{%s}Default" % CT_NS)}
    overrides = {e.get("PartName", "").lstrip("/") for e in ct.findall("{%s}Override" % CT_NS)}
    for name in overrides:
        if name not in parts:
            bad("E-OPC", name, "[Content_Types] 有 Override 但部件不存在")
    xml_ok: Dict[str, Any] = {}
    for name, data in parts.items():
        if name == "[Content_Types].xml":
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if name not in overrides and ext not in defaults:
            bad("E-OPC", name, "部件没有 Content Type（既无 Override 也无同扩展名 Default）")
        if name.endswith((".xml", ".rels")):
            try:
                xml_ok[name] = parse_xml(data)
            except Exception as exc:
                bad("E-XML", name, f"XML 不可解析：{type(exc).__name__}")
    stats = {"parts": len(parts), "slides": 0, "xsd_checked": 0, "xsd_skipped_unchanged": 0}
    owners: Dict[str, Set[str]] = {}
    for name in [n for n in parts if n.endswith(".rels")]:
        src = posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5]) if name != "_rels/.rels" else ""
        try:
            rels = part_rels(parts, src) if src else [
                {"id": r.get("Id"), "type": r.get("Type"), "part": None if (r.get("TargetMode") or "").lower() == "external" else resolve_target("", r.get("Target"))}
                for r in parse_xml(parts[name])]
        except Exception:
            continue
        ids = [r["id"] for r in rels]
        for d in sorted({i for i in ids if ids.count(i) > 1}):
            bad("E-REL", name, f"关系 Id 重复：{d}")
        for r in rels:
            if r["part"] and r["part"] not in parts:
                bad("E-REL", name, f"关系 {r['id']} 指向不存在的部件 {r['part']}")
            elif r["part"] and any(r["type"].endswith(t) for t in EXCLUSIVE) and src.startswith("ppt/slides/"):
                owners.setdefault(r["part"], set()).add(src)
    for part, slides in sorted(owners.items()):
        if len(slides) > 1:
            bad("E-SHARED", part, f"被 {len(slides)} 页共同引用（{sorted(slides)[:3]}）：改一页会同步改另一页；复制页须深拷贝该部件")

    # ---- 引用
    try:
        slides = listed_slides(parts)
    except Exception as exc:
        slides = []
        bad("E-PRES", "ppt/presentation.xml", f"无法读取 sldIdLst：{type(exc).__name__}")
    stats["slides"] = len(slides)
    sld_ids = [s["sld_id"] for s in slides]
    for d in sorted({i for i in sld_ids if sld_ids.count(i) > 1}):
        bad("E-PRES", "ppt/presentation.xml", f"sldId 重复：{d}")
    for s in slides:
        if s["part"] is None or s["part"] not in parts:
            bad("E-PRES", "ppt/presentation.xml", f"sldId {s['sld_id']} 的关系 {s['rid']} 没有可用的幻灯片部件")
        if s["sld_id"] and not (256 <= int(s["sld_id"]) < 2147483648):
            bad("E-PRES", "ppt/presentation.xml", f"sldId {s['sld_id']} 超出合法范围 [256, 2^31)")
    listed = {s["part"] for s in slides}
    for name in parts:
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) and name not in listed:
            bad("W-ORPHAN", name, "幻灯片部件存在但未登记在 sldIdLst（clean 可清理）")

    # ---- 逐页：id 唯一 / 动画完整性 / XSD / 图表
    base = baseline or {}
    inherited: Dict[str, Set[str]] = {}
    for name in sorted(n for n in parts if XSD_PARTS.fullmatch(n) and n in xml_ok):
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
            root = xml_ok[name]
            ids = [c.get("id") for c in root.iter(motion.pn("cNvPr"))]
            for d in sorted({i for i in ids if ids.count(i) > 1}, key=str):
                bad("E-ID", name, f"同页 cNvPr@id 重复：{d}")
            try:
                cat = motion.catalog(root, _chart_info(parts, name, part_rels(parts, name)))
            except Exception:
                cat = None
            for m in motion.integrity_problems(root, cat):
                bad(m.split(" ", 1)[0] if m.startswith("E-") else "E-MOTION", name, m.split(" ", 1)[-1])
        if not xsd:
            continue
        if name in base and base[name] == parts[name]:
            stats["xsd_skipped_unchanged"] += 1
            continue
        errs = xsd_errors(parts[name])
        stats["xsd_checked"] += 1
        if errs and name in base:
            try:
                inherited[name] = set(xsd_errors(base[name]))
            except Exception:
                inherited[name] = set()
        for e in errs:
            if e not in inherited.get(name, set()):
                bad("E-XSD", name, e[:300])
    from post import chart_problems   # 延迟导入：post 依赖 motion，避免循环
    for name in sorted(n for n in parts if re.fullmatch(r"ppt/charts/chart\d+\.xml", n) and n in xml_ok):
        if name in base and base[name] == parts[name]:
            continue
        for m in chart_problems(parts[name], name):
            bad("E-CHART", name, m)
    hard = [p for p in probs if p["code"].startswith("E-")]
    return {"status": "BLOCK" if hard else "PASS", "problems": probs, "stats": stats}


def check(path: Path, baseline: Optional[Path] = None, xsd: bool = True) -> Dict[str, Any]:
    path = Path(path)
    pk = package_problems(path, check_suffix=False)
    if pk:
        return {"status": "BLOCK", "problems": [{"code": "E-FORMAT", "where": str(path), "message": m} for m in pk], "stats": {}}
    parts = read_parts(path)
    res = check_parts(parts, read_parts(Path(baseline)) if baseline else None, xsd)
    res["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    return res


def main(argv: Optional[List[str]] = None) -> int:
    utf8_stdio()
    ap = argparse.ArgumentParser(description="PPTX 结构层校验（OPC / 关系 / XSD / 图表 / 动画引用）")
    ap.add_argument("pptx")
    ap.add_argument("--baseline", help="原稿：相同部件不重验，原稿已有的 XSD 缺陷不算新增")
    ap.add_argument("--no-xsd", action="store_true")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)
    if not Path(args.pptx).is_file():
        print(f"[HARD] E-SRC 找不到 {args.pptx}")
        return 2
    res = check(Path(args.pptx), Path(args.baseline) if args.baseline else None, not args.no_xsd)
    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        for p in res["problems"]:
            print(f"[{'HARD' if p['code'].startswith('E-') else 'SOFT'}] {p['code']} {p['where']}: {p['message']}")
        print(f"结构层 {res['status']}（OPC/关系/XSD/图表/动画引用）｜{res['stats']}｜PowerPoint 实机：NOT VERIFIED")
    return 0 if res["status"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
