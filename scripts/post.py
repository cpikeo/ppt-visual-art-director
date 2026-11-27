#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OOXML 后处理（L4）：把 PptxGenJS 原始包补成 VAD 契约要求的成品。单次读 zip、单次写 zip，不经 python-pptx 重存。

补齐的能力（均为 PptxGenJS 无法直接表达或表达不合规的部分）：
  图片 srcRect 裁切 / 预设与多边形蒙版 / 双色调 / 软边；形状渐变与透明渐变、圆角 adj、线帽、箭头；
  原生分组 p:grpSp；替代文字；标题占位符命名；列表符号色；语言标记（zh-CN/en-US）；
  图表：series 子元素次序（PptxGenJS 把 dLbls 写在 dPt 之前，违反 schema）、逐系列直标、东亚/复杂文字字体、
        标题左对齐、饼图内缩、散点只留标记、组合折线标记；图表 XML 兼容性自检（轴引用、序号、dLblPos、数值、线型）；
  主题色/主题字体、母版标题样式、表格样式、文档属性、媒体去重、zip 条目时间与内嵌工作簿时间钉死（可复现）。
"""

from __future__ import annotations

import copy
import hashlib
import io
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from lxml import etree

from model import CANVAS_EMU, LIMITS, PT_EMU, allowed_label_positions, parse_xml, qn
from ooxml import stable_xlsx

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CJK = re.compile("[\u3400-\u9fff\u3000-\u303f\uff00-\uffef]")
FIXED = (2000, 1, 1, 0, 0, 0)


class PostError(Exception):
    """后处理发现包不合格（BLOCK）：交付路径不写。"""


def E(tag: str, **attrs) -> etree._Element:
    node = etree.Element(qn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def sub(parent, tag: str, **attrs):
    node = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def _flat(elements: List[Dict[str, Any]], out: Optional[Dict[str, Dict[str, Any]]] = None) -> Dict[str, Dict[str, Any]]:
    out = {} if out is None else out
    for e in elements:
        out[e["name"]] = e
        _flat(e.get("children") or [], out)
    return out


# ---------------------------------------------------------------- 幻灯片


def _cnvpr(node):
    hits = node.xpath("./*[1]/p:cNvPr", namespaces=NS)
    return hits[0] if hits else None


def _xfrm(node):
    for path in ("./p:spPr/a:xfrm", "./p:xfrm", "./p:grpSpPr/a:xfrm"):
        hit = node.xpath(path, namespaces=NS)
        if hit:
            return hit[0]
    return None


def _box(node):
    x = _xfrm(node)
    off, ext = x.find(qn("a:off")), x.find(qn("a:ext"))
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def _alpha(clr, alpha: Optional[float]) -> None:
    if alpha is not None and alpha < 1:
        sub(clr, "a:alpha", val=int(round(alpha * 100000)))


def _solid(hexv: str, alpha: Optional[float] = None):
    f = E("a:solidFill")
    c = sub(f, "a:srgbClr", val=hexv)
    _alpha(c, alpha)
    return f


def _fix_line(spPr, line: Optional[Dict[str, Any]]) -> None:
    ln = spPr.find(qn("a:ln"))
    if line is None:
        if ln is None:
            ln = E("a:ln")
            _insert_spPr(spPr, ln)
        for ch in list(ln):
            ln.remove(ch)
        sub(ln, "a:noFill")
        return
    if ln is None:
        return
    if line.get("cap"):
        ln.set("cap", {"flat": "flat", "round": "rnd", "square": "sq"}[line["cap"]])
    for tag in ("a:headEnd", "a:tailEnd"):
        for old in ln.findall(qn(tag)):
            ln.remove(old)
    arrow = line.get("arrow") or "none"
    if arrow in ("end", "both"):
        if arrow == "both":
            sub(ln, "a:headEnd", type="triangle", w="med", len="med")
        sub(ln, "a:tailEnd", type="triangle", w="med", len="med")


_SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln", "effectLst",
               "effectDag", "scene3d", "sp3d", "extLst"]


def _insert_spPr(spPr, node) -> None:
    """按 CT_ShapeProperties 的子元素次序插入。"""
    rank = _SPPR_ORDER.index(etree.QName(node).localname)
    for child in spPr:
        if _SPPR_ORDER.index(etree.QName(child).localname) > rank:
            child.addprevious(node)
            return
    spPr.append(node)


def _gradient(spPr, g: Dict[str, Any]) -> None:
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for old in spPr.findall(qn(tag)):
            spPr.remove(old)
    gf = E("a:gradFill", rotWithShape="1")
    lst = sub(gf, "a:gsLst")
    for pos, key, ak in ((0, "from", "from_alpha"), (100000, "to", "to_alpha")):
        gs = sub(lst, "a:gs", pos=pos)
        c = sub(gs, "a:srgbClr", val=g[key])
        _alpha(c, g.get(ak))
    sub(gf, "a:lin", ang=int(round((g["angle"] % 360) * 60000)), scaled="0")
    _insert_spPr(spPr, gf)


def _text_fix(sp, e: Dict[str, Any]) -> None:
    body = sp.find(qn("p:txBody"))
    if body is None:
        return
    bp = body.find(qn("a:bodyPr"))
    # 版式占位符会盖掉 JS 传入的几何与锚点：这里以 spec 为准，逐项写回
    if sp.xpath("./p:nvSpPr/p:nvPr/p:ph", namespaces=NS):
        xf = sp.find(qn("p:spPr")).find(qn("a:xfrm"))
        xf.find(qn("a:off")).set("x", str(int(round(e["box"][0] * 914400)))), xf.find(qn("a:off")).set("y", str(int(round(e["box"][1] * 914400))))
        xf.find(qn("a:ext")).set("cx", str(int(round(e["box"][2] * 914400)))), xf.find(qn("a:ext")).set("cy", str(int(round(e["box"][3] * 914400))))
    if bp is not None:
        bp.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[e["valign"]])
        bp.set("lIns", str(int(round(e["pad"][0] * 914400)))), bp.set("rIns", str(int(round(e["pad"][0] * 914400))))
        bp.set("tIns", str(int(round(e["pad"][1] * 914400)))), bp.set("bIns", str(int(round(e["pad"][1] * 914400))))
        bp.set("wrap", "square" if e["wrap"] else "none")
    if bp is not None and not any(etree.QName(c).localname in ("noAutofit", "normAutofit", "spAutoFit") for c in bp):
        sub(bp, "a:noAutofit")
    paras = body.findall(qn("a:p"))
    for i, p in enumerate(paras):
        for dup in p.findall(qn("a:pPr"))[1:]:   # PptxGenJS 在占位符里会为每个 run 重复写 a:pPr（CT_TextParagraph 只允许一个，且须在最前）
            p.remove(dup)
        spec_p = e["paras"][i] if i < len(e["paras"]) else None
        lang = "en-US"
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            t = "".join(x.text or "" for x in r.findall(qn("a:t")))
            if rpr is not None:
                lang = "zh-CN" if _CJK.search(t) else "en-US"
                rpr.set("lang", lang)
                rpr.attrib.pop("kern", None) if rpr.get("kern") == "0" and False else None
        end = p.find(qn("a:endParaRPr"))
        if end is not None:
            end.set("lang", lang)
        ppr = p.find(qn("a:pPr"))
        if spec_p and spec_p.get("list") and ppr is not None:
            for old in ppr.findall(qn("a:buClr")):
                ppr.remove(old)
            clr = E("a:buClr")
            sub(clr, "a:srgbClr", val=spec_p["list"]["color"])
            anchor = None
            for ch in ppr:
                if etree.QName(ch).localname in ("lnSpc", "spcBef", "spcAft"):
                    anchor = ch
            if anchor is not None:
                anchor.addnext(clr)
            else:
                ppr.insert(0, clr)
            if spec_p["list"]["kind"] == "number":
                for bc in ppr.findall(qn("a:buChar")):
                    ppr.remove(bc)
                if ppr.find(qn("a:buAutoNum")) is None:
                    sub(ppr, "a:buFont", typeface="+mj-lt")
                    sub(ppr, "a:buAutoNum", type="arabicPeriod")
            elif ppr.find(qn("a:buFont")) is None:
                buc = ppr.find(qn("a:buChar"))
                if buc is not None:
                    bf = E("a:buFont", typeface="Arial")
                    buc.addprevious(bf)


def _pic_fix(pic, e: Dict[str, Any]) -> None:
    ox = e.get("ooxml") or {}
    blipFill = pic.find(qn("p:blipFill"))
    blip = blipFill.find(qn("a:blip"))
    crop = ox.get("crop")
    for old in blipFill.findall(qn("a:srcRect")):
        blipFill.remove(old)
    if crop and any(v > 1e-9 for v in crop):
        l, r, t, b = crop
        blip.addnext(E("a:srcRect", l=int(round(l * 100000)), t=int(round(t * 100000)), r=int(round(r * 100000)), b=int(round(b * 100000))))
    if ox.get("duotone"):  # alphaModFix 先于 duotone（PptxGenJS 已先写透明度）
        d = E("a:duotone")
        for k in ("dark", "light"):
            sub(d, "a:srgbClr", val=ox["duotone"][k])
        blip.append(d)
    spPr = pic.find(qn("p:spPr"))
    mask = ox.get("mask")
    if mask:
        old = spPr.find(qn("a:prstGeom"))
        if isinstance(mask, str):
            old.set("prst", mask)
            for av in old.findall(qn("a:avLst")):
                av.clear()
        else:
            cust = E("a:custGeom")
            for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
                sub(cust, tag)
            sub(cust, "a:rect", l="0", t="0", r="r", b="b")
            path = sub(sub(cust, "a:pathLst"), "a:path", w=100000, h=100000)
            for i, (px, py) in enumerate(mask):
                node = sub(path, "a:moveTo" if i == 0 else "a:lnTo")
                sub(node, "a:pt", x=int(round(px * 100000)), y=int(round(py * 100000)))
            sub(path, "a:close")
            spPr.replace(old, cust)
    if ox.get("soft_edge_pt"):
        for old in spPr.findall(qn("a:effectLst")):
            spPr.remove(old)
        eff = E("a:effectLst")
        sub(eff, "a:softEdge", rad=int(round(ox["soft_edge_pt"] * PT_EMU)))
        _insert_spPr(spPr, eff)


def _as_connector(sp, by_name: Dict[str, Any], name: str):
    cxn = E("p:cxnSp")
    nv = sub(cxn, "p:nvCxnSpPr")
    nv.append(copy.deepcopy(sp.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))))
    sub(nv, "p:cNvCxnSpPr")
    sub(nv, "p:nvPr")
    cxn.append(sp.find(qn("p:spPr")))
    sp.addprevious(cxn)
    sp.getparent().remove(sp)
    by_name[name] = cxn
    return cxn


def _wrap_groups(tree, elements: List[Dict[str, Any]], by_name: Dict[str, Any]) -> None:
    next_id = [max([int(x) for x in tree.xpath("//p:cNvPr/@id", namespaces=NS)] + [1]) + 1]

    def node_of(e):
        if e["kind"] != "group":
            return by_name.get(e["name"])
        kids = [k for k in (node_of(c) for c in e["children"]) if k is not None]
        if not kids:
            return None
        xs = [_box(k) for k in kids]
        x0, y0 = min(b[0] for b in xs), min(b[1] for b in xs)
        x1, y1 = max(b[0] + b[2] for b in xs), max(b[1] + b[3] for b in xs)
        g = E("p:grpSp")
        nv = sub(g, "p:nvGrpSpPr")
        pr = sub(nv, "p:cNvPr", id=next_id[0], name=e["name"])
        if e.get("alt"):
            pr.set("descr", e["alt"])
        next_id[0] += 1
        sub(nv, "p:cNvGrpSpPr")
        sub(nv, "p:nvPr")
        gp = sub(g, "p:grpSpPr")
        xf = sub(gp, "a:xfrm")
        if e.get("rotation"):
            xf.set("rot", str(int(round(e["rotation"] * 60000))))
        sub(xf, "a:off", x=x0, y=y0)
        sub(xf, "a:ext", cx=x1 - x0, cy=y1 - y0)
        sub(xf, "a:chOff", x=x0, y=y0)
        sub(xf, "a:chExt", cx=x1 - x0, cy=y1 - y0)
        kids[0].addprevious(g)
        for k in kids:
            g.append(k)  # 移动（lxml 会从原父节点摘下）
        return g

    for e in elements:
        if e["kind"] == "group":
            node_of(e)


def _slide(xml: bytes, page: Dict[str, Any], rels: Dict[str, str]) -> (bytes, Dict[str, str]):
    root = parse_xml(xml)
    tree = root.find(".//" + qn("p:spTree"))
    by_spec = _flat(page["elements"] + ([page["page_number"]] if page.get("page_number") else []))
    by_name: Dict[str, Any] = {}
    charts: Dict[str, str] = {}
    title_spec = next((e for e in page["elements"] if e.get("placeholder") == "title"), None)
    for node in list(tree):
        if node.tag not in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:cxnSp")):
            continue
        pr = _cnvpr(node)
        if pr is None:
            continue
        if node.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='title']", namespaces=NS) and title_spec is not None:
            pr.set("name", title_spec["name"])
        e = by_spec.get(pr.get("name"))
        if e is None:
            continue
        by_name[e["name"]] = node
        if e.get("alt"):
            pr.set("descr", e["alt"])
        else:
            pr.attrib.pop("descr", None)
        k = e["kind"]
        spPr = node.find(qn("p:spPr"))
        if k == "text":
            _text_fix(node, e)
        elif k in ("shape", "path"):
            ox = e.get("ooxml") or {}
            if ox.get("gradient"):
                _gradient(spPr, ox["gradient"])
            if ox.get("adj") is not None:
                prst = spPr.find(qn("a:prstGeom"))
                if prst is not None:
                    for av in prst.findall(qn("a:avLst")):
                        prst.remove(av)
                    sub(sub(prst, "a:avLst"), "a:gd", name="adj", fmla=f"val {ox['adj']}")
            _fix_line(spPr, e.get("line"))
            if k == "path" and len(e["points"]) == 2 and node.tag == qn("p:sp"):
                node = _as_connector(node, by_name, e["name"])   # 两点路径 = 原生直线连接符（PptxGenJS 只写 p:sp）
        elif k == "image":
            _pic_fix(node, e)
        elif k == "table":
            tblPr = node.find(".//" + qn("a:tblPr"))
            for old in tblPr.findall(qn("a:tableStyleId")):
                tblPr.remove(old)
            tblPr.attrib.pop("firstRow", None)
            tblPr.attrib.pop("bandRow", None)
            sub(tblPr, "a:tableStyleId").text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
        elif k == "chart":
            rid = node.find(".//" + qn("c:chart")).get(qn("r:id"))
            charts[rels[rid]] = e["name"]
    _wrap_groups(tree, page["elements"], by_name)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), charts


# ---------------------------------------------------------------- 图表

_SER_ORDER = ["idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker", "explosion", "dPt", "dLbls", "trendline",
              "errBars", "cat", "val", "xVal", "yVal", "smooth", "shape", "extLst"]
_PLOTS = ("barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart")
_VALID_DASH = {"solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"}


def _txpr(ink: str, size_pt: float, family: str):
    tx = E("c:txPr")
    sub(tx, "a:bodyPr")
    sub(tx, "a:lstStyle")
    p = sub(tx, "a:p")
    rpr = sub(sub(p, "a:pPr"), "a:defRPr", sz=int(size_pt * 100))
    sub(sub(rpr, "a:solidFill"), "a:srgbClr", val=ink)
    for t in ("a:latin", "a:ea", "a:cs"):
        sub(rpr, t, typeface=family)
    sub(p, "a:endParaRPr", lang="zh-CN")
    return tx


def _label_body(parent, ink, fmt, pos, family, plate=None, show_cat=False, sep=None) -> None:
    """CT_DLbl / Group_DLbls 次序：numFmt, spPr, txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName, showPercent, showBubbleSize, separator。"""
    sub(parent, "c:numFmt", formatCode=fmt, sourceLinked="0")
    sp = sub(parent, "c:spPr")
    if plate:
        sub(sub(sp, "a:solidFill"), "a:srgbClr", val=plate)
    else:
        sub(sp, "a:noFill")
    sub(sub(sp, "a:ln"), "a:noFill")
    parent.append(_txpr(ink, LIMITS["chart"]["font_pt"]["label"], family))
    if pos:
        sub(parent, "c:dLblPos", val=pos)
    for tag, on in (("showLegendKey", 0), ("showVal", 1), ("showCatName", int(show_cat)), ("showSerName", 0), ("showPercent", 0), ("showBubbleSize", 0)):
        sub(parent, "c:" + tag, val=on)
    if sep is not None:
        sub(parent, "c:separator").text = sep


def _dlbls(ink, fmt, pos, family, plate=None, show_cat=False, sep=None, point_inks=None, leader=False):
    d = E("c:dLbls")
    for i, pink in enumerate(point_inks or []):
        dl = sub(d, "c:dLbl")
        sub(dl, "c:idx", val=i)
        _label_body(dl, pink, fmt, pos, family, plate, show_cat, sep)
    _label_body(d, ink, fmt, pos, family, plate, show_cat, sep)
    if leader:
        sub(d, "c:showLeaderLines", val=1)
    return d


def _ser_name(ser) -> str:
    v = ser.xpath("./c:tx//c:v", namespaces=NS)
    return v[0].text if v else ""


def _marker(ser, color: str) -> None:
    for old in ser.findall(qn("c:marker")):
        ser.remove(old)
    m = E("c:marker")
    sub(m, "c:symbol", val="circle")
    sub(m, "c:size", val=7)
    sp = sub(m, "c:spPr")
    sub(sub(sp, "a:solidFill"), "a:srgbClr", val=color)
    sub(sub(sub(sp, "a:ln"), "a:solidFill"), "a:srgbClr", val=color)
    ser.append(m)


def _chart_xml(xml: bytes, e: Dict[str, Any]) -> bytes:
    root = parse_xml(xml)
    chart = root.find(qn("c:chart"))
    plot_area = chart.find(qn("c:plotArea"))
    fam, kind = e["family"], e["chart_kind"]
    plots = [p for p in plot_area if etree.QName(p).localname in _PLOTS]
    # 0) PptxGenJS 缺陷修正：2D 绘图层带了指向不存在的轴的第三个 axId（CT_BarChart 等只允许 2 个）；
    #    非柱图 series 里混入 invertIfNegative（CT_LineSer / CT_AreaSer / CT_ScatterSer 不允许）。
    axis_ids = {ax.find(qn("c:axId")).get("val") for ax in plot_area if etree.QName(ax).localname in ("catAx", "valAx", "dateAx", "serAx")}
    for plot in plots:
        for ax in plot.findall(qn("c:axId")):
            if ax.get("val") not in axis_ids:
                plot.remove(ax)
        if etree.QName(plot).localname != "barChart":
            for ser in plot.findall(qn("c:ser")):
                for bad in ser.findall(qn("c:invertIfNegative")):
                    ser.remove(bad)
    # 1) series 子元素次序（PptxGenJS 的 dLbls 在 dPt 之前，PowerPoint 严格校验会要求修复）
    for ser in root.iter(qn("c:ser")):
        kids = sorted(list(ser), key=lambda c: _SER_ORDER.index(etree.QName(c).localname))
        for c in list(ser):
            ser.remove(c)
        for c in kids:
            ser.append(c)
    # 2) 逐系列直标
    plan_by_name = {x["name"]: x for x in (e.get("label_plan") or [])}
    combo = set(e.get("combo") or [])
    for plot in plots:
        tag = etree.QName(plot).localname
        for i, ser in enumerate(plot.findall(qn("c:ser"))):
            name = _ser_name(ser)
            if tag in ("pieChart", "doughnutChart"):
                ent = (e.get("label_plan") or [None])[0]
            elif tag == "lineChart" and name in combo:
                ent = plan_by_name.get(name) if e["combo_line_labels"] else None
            else:
                ent = plan_by_name.get(name) if e["data_labels"] else None
            for old in ser.findall(qn("c:dLbls")):
                ser.remove(old)
            if ent is None:
                continue
            d = _dlbls(ent["ink"], e["label_format"], ent["position"], fam, ent.get("plate"), show_cat=(tag == "pieChart"),
                       sep=" " if tag == "pieChart" else None, point_inks=ent.get("points"), leader=(tag == "pieChart"))
            nxt = next((c for c in ser if etree.QName(c).localname in _SER_ORDER[_SER_ORDER.index("dLbls") + 1:]), None)
            nxt.addprevious(d) if nxt is not None else ser.append(d)
        # 绘图层级的 dLbls 一律关闭，避免与系列级冲突
        for d in plot.findall(qn("c:dLbls")):
            for sv in d.findall(qn("c:showVal")):
                sv.set("val", "0")
    # 2b) 柱图系列底色以系列色为准（高亮柱的“退隐色”只落在 dPt 上，系列本身仍是主色）
    for ser in root.iter(qn("c:ser")):
        if ser.getparent().tag == qn("c:barChart") and _ser_name(ser) in [s["name"] for s in e["series"]]:
            col = e["colors"][[s["name"] for s in e["series"]].index(_ser_name(ser))]
            sp = ser.find(qn("c:spPr"))
            for ch in list(sp):
                sp.remove(ch)
            sub(sub(sp, "a:solidFill"), "a:srgbClr", val=col)
            sub(sub(sp, "a:ln"), "a:noFill")
    # 3) 组合折线、折线与散点：标记与平滑
    colors = e["colors"]
    names = [s["name"] for s in e["series"]]
    for plot in plots:
        tag = etree.QName(plot).localname
        for ser in plot.findall(qn("c:ser")):
            nm = _ser_name(ser)
            color = colors[names.index(nm)] if nm in names and kind not in ("pie", "donut") else None
            if tag == "lineChart":
                if kind == "line":
                    for old in ser.findall(qn("c:marker")):
                        ser.remove(old)
                    m = E("c:marker")
                    sub(m, "c:symbol", val="none")
                    ser.find(qn("c:spPr")).addnext(m)
                elif color:
                    _marker(ser, color)
                    # marker 必须排在 spPr 之后、dPt 之前
                    mk = ser.find(qn("c:marker"))
                    ser.remove(mk)
                    ser.find(qn("c:spPr")).addnext(mk)
                for sm in ser.findall(qn("c:smooth")):
                    sm.set("val", "0")
                if ser.find(qn("c:smooth")) is None:
                    sub(ser, "c:smooth", val=0)
            elif tag == "scatterChart" and color:
                sp = ser.find(qn("c:spPr"))
                if sp is None:
                    sp = E("c:spPr")
                    ser.find(qn("c:tx")).addnext(sp)
                for ch in list(sp):
                    sp.remove(ch)
                sub(sub(sp, "a:ln", w=19050), "a:noFill")
                for old in ser.findall(qn("c:marker")):
                    ser.remove(old)
                _marker(ser, color)
                mk = ser.find(qn("c:marker"))
                ser.remove(mk)
                sp.addnext(mk)
            elif tag == "areaChart" and color:
                sp = ser.find(qn("c:spPr"))
                for ch in list(sp):
                    sp.remove(ch)
                sub(sub(sp, "a:solidFill"), "a:srgbClr", val=color)
                sub(sub(sub(sp, "a:ln", w=9525), "a:solidFill"), "a:srgbClr", val=e["surface"])
        if tag == "scatterChart":
            ss = plot.find(qn("c:scatterStyle"))
            if ss is not None:
                ss.set("val", "marker")
    # 4) 饼图：内缩绘图区给外侧直标留空间
    if kind == "pie" and e["data_labels"]:
        layout = plot_area.find(qn("c:layout"))
        if layout is None:
            layout = E("c:layout")
            plot_area.insert(0, layout)
        for ch in list(layout):
            layout.remove(ch)
        man = sub(layout, "c:manualLayout")
        for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"), ("x", "0.24"), ("y", "0.2"), ("w", "0.52"), ("h", "0.66")):
            sub(man, "c:" + tag, val=val)
    # 5) 标题：左上固定、左对齐
    title = chart.find(qn("c:title"))
    if title is not None:
        for pp in title.iter(qn("a:pPr")):
            pp.set("algn", "l")
        layout = title.find(qn("c:layout"))
        if layout is None:
            layout = E("c:layout")
            title.find(qn("c:tx")).addnext(layout)
        for ch in list(layout):
            layout.remove(ch)
        man = sub(layout, "c:manualLayout")
        for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"), ("x", "0"), ("y", "0.02"), ("w", "1"), ("h", "0.14")):
            sub(man, "c:" + tag, val=val)
        if title.find(qn("c:overlay")) is None:
            layout.addnext(E("c:overlay", val="0"))
    # 6) 东亚/复杂文字字体：凡写了 latin 的文字属性，同步 ea 与 cs
    for latin in list(root.iter(qn("a:latin"))):
        face = latin.get("typeface")
        parent = latin.getparent()
        prev = latin
        for t in ("a:ea", "a:cs"):
            node = parent.find(qn(t))
            if node is None:
                node = E(t)
                prev.addnext(node)
            node.set("typeface", face)
            for att in ("pitchFamily", "charset"):
                node.attrib.pop(att, None)
            prev = node
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def chart_problems(xml: bytes, label: str) -> List[str]:
    """PowerPoint 兼容性自检（借鉴 Anthropic pptx skill 的图表 XML 校验思路）。LibreOffice 能画不等于 PowerPoint 能开。"""
    root = parse_xml(xml)
    out: List[str] = []
    plot_area = root.find(qn("c:chart")).find(qn("c:plotArea"))
    axes = {}
    for ax in plot_area:
        if etree.QName(ax).localname in ("catAx", "valAx", "dateAx", "serAx"):
            axes[ax.find(qn("c:axId")).get("val")] = ax
    for ax_id, ax in axes.items():
        ref = ax.find(qn("c:crossAx"))
        if ref is None or ref.get("val") not in axes:
            out.append(f"{label}: 轴 {ax_id} 的 crossAx 指向不存在的轴")
    idxs, orders = set(), set()
    for plot in plot_area:
        tag = etree.QName(plot).localname
        if tag not in _PLOTS:
            continue
        if tag not in ("pieChart", "doughnutChart"):
            ids = [a.get("val") for a in plot.findall(qn("c:axId"))]
            if len(ids) < 2 or any(i not in axes for i in ids):
                out.append(f"{label}: {tag} 引用了不存在的轴 {ids}")
        grouping = (plot.find(qn("c:grouping")).get("val") if plot.find(qn("c:grouping")) is not None else "clustered")
        allowed = allowed_label_positions(tag, grouping)
        for ser in plot.findall(qn("c:ser")):
            i, o = ser.find(qn("c:idx")).get("val"), ser.find(qn("c:order")).get("val")
            if i in idxs or o in orders:
                out.append(f"{label}: series idx/order 重复（{i}/{o}）")
            idxs.add(i), orders.add(o)
            names = [etree.QName(c).localname for c in ser]
            if names != sorted(names, key=_SER_ORDER.index):
                out.append(f"{label}: c:ser 子元素次序违反 schema：{names}")
            for pos in ser.iter(qn("c:dLblPos")):
                if allowed is not None and pos.get("val") not in allowed:
                    out.append(f"{label}: {tag} 不允许 dLblPos={pos.get('val')}（PowerPoint 会报修复）")
    for v in root.iter(qn("c:v")):
        if v.getparent().tag == qn("c:pt") and v.getparent().getparent().tag == qn("c:numCache"):
            try:
                x = float(v.text)
                if x != x or x in (float("inf"), float("-inf")):
                    raise ValueError
            except Exception:
                out.append(f"{label}: 数值缓存含无效值 {v.text!r}")
    for d in root.iter(qn("a:prstDash")):
        if d.get("val") not in _VALID_DASH:
            out.append(f"{label}: 非法线型 prstDash={d.get('val')}")
    return out


# ---------------------------------------------------------------- 包


def _rels_of(zf: Dict[str, bytes], part: str) -> Dict[str, str]:
    name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if name not in zf:
        return {}
    root = parse_xml(zf[name])
    def target(t: str) -> str:
        return t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), t))
    return {r.get("Id"): target(r.get("Target")) for r in root}


def _theme(xml: bytes, theme: Dict[str, Any]) -> bytes:
    root = parse_xml(xml)
    scheme = root.find(".//" + qn("a:clrScheme"))
    for name, hexv in theme["colors"].items():
        node = scheme.find(qn("a:" + name))
        if node is not None:
            for ch in list(node):
                node.remove(ch)
            sub(node, "a:srgbClr", val=hexv)
    scheme.set("name", "PPT Visual Art Director")
    for tag, fam in (("majorFont", theme["heading"]), ("minorFont", theme["body"])):
        node = root.find(f".//{qn('a:fontScheme')}/{qn('a:' + tag)}")
        for kind in ("latin", "ea"):
            ch = node.find(qn("a:" + kind))
            if ch is not None:
                ch.set("typeface", fam)
                ch.attrib.pop("panose", None)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _master(xml: bytes, theme: Dict[str, Any]) -> bytes:
    root = parse_xml(xml)
    ts = root.find(f".//{qn('p:txStyles')}/{qn('p:titleStyle')}")
    if ts is not None:
        for lvl in ts:
            if etree.QName(lvl).localname.startswith("lvl"):
                lvl.set("algn", "l")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _dedupe_media(parts: Dict[str, bytes]) -> None:
    seen: Dict[str, str] = {}
    alias: Dict[str, str] = {}
    for name in sorted(n for n in parts if n.startswith("ppt/media/")):
        h = hashlib.sha256(parts[name]).hexdigest()
        if h in seen:
            alias[name] = seen[h]
        else:
            seen[h] = name
    for dup, keep in alias.items():
        del parts[dup]
    if alias:
        for name in [n for n in parts if n.endswith(".rels")]:
            s = parts[name].decode("utf-8")
            for dup, keep in alias.items():
                s = s.replace("../media/" + posixpath.basename(dup), "../media/" + posixpath.basename(keep))
            parts[name] = s.encode("utf-8")


def _content_types(parts: Dict[str, bytes]) -> bytes:
    """PptxGenJS 会为不存在的 slideMaster2..N 写 Override：删掉指向缺失部件的条目（否则包不合格）。"""
    root = parse_xml(parts["[Content_Types].xml"])
    for ov in list(root):
        if etree.QName(ov).localname == "Override" and ov.get("PartName").lstrip("/") not in parts:
            root.remove(ov)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _app_props(n_pages: int, n_notes: int) -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f'<TotalTime>0</TotalTime><Application>ppt-visual-art-director</Application>'
            f'<PresentationFormat>Widescreen</PresentationFormat><Slides>{n_pages}</Slides>'
            f'<Notes>{n_notes}</Notes><HiddenSlides>0</HiddenSlides></Properties>').encode("utf-8")


def _core(xml: bytes, core: Dict[str, str]) -> bytes:
    root = parse_xml(xml)
    dc, dcterms = "http://purl.org/dc/elements/1.1/", "http://purl.org/dc/terms/"
    cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

    def put(ns, local, text):
        node = root.find(f"{{{ns}}}{local}")
        if node is None:
            node = etree.SubElement(root, f"{{{ns}}}{local}")
        node.text = text
        return node
    put(dc, "title", core["title"])
    put(dc, "creator", core["author"])
    put(dc, "description", core["comments"])
    for local in ("created", "modified"):
        n = put(dcterms, local, "2000-01-01T00:00:00Z")
        n.set("{http://www.w3.org/2001/XMLSchema-instance}type", "dcterms:W3CDTF")
    root.append(root.find(f"{{{cp}}}lastModifiedBy")) if root.find(f"{{{cp}}}lastModifiedBy") is not None else None
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def postprocess(raw: Path, spec: Dict[str, Any], out: Path, rendered_pages: Optional[List[str]] = None) -> Dict[str, Any]:
    """raw → out。返回统计，供报告使用。包不合格抛 PostError。"""
    with zipfile.ZipFile(raw) as z:
        order = [i.filename for i in z.infolist() if not i.filename.endswith("/")]
        parts = {n: z.read(n) for n in order}
    pages = [p for p in spec["pages"] if rendered_pages is None or p["id"] in rendered_pages]
    problems: List[str] = []

    # 画布：必须与 CANVAS_EMU 精确一致，不缩放、不补偿
    pres = parse_xml(parts["ppt/presentation.xml"])
    sz = pres.find(qn("p:sldSz"))
    if (int(sz.get("cx")), int(sz.get("cy"))) != tuple(CANVAS_EMU):
        raise PostError(f"画布 {sz.get('cx')}×{sz.get('cy')} ≠ CANVAS_EMU {CANVAS_EMU}")
    sz.attrib.pop("type", None)
    parts["ppt/presentation.xml"] = etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True)

    chart_specs: Dict[str, Dict[str, Any]] = {}
    for i, page in enumerate(pages, 1):
        part = f"ppt/slides/slide{i}.xml"
        xml, charts = _slide(parts[part], page, _rels_of(parts, part))
        parts[part] = xml
        by_name = _flat(page["elements"])
        for chart_part, name in charts.items():
            chart_specs[chart_part] = by_name[name]
    for chart_part, e in chart_specs.items():
        parts[chart_part] = _chart_xml(parts[chart_part], e)
        problems += chart_problems(parts[chart_part], chart_part)
    if problems:
        raise PostError("；".join(problems))

    parts["ppt/theme/theme1.xml"] = _theme(parts["ppt/theme/theme1.xml"], spec["theme"])
    parts["ppt/slideMasters/slideMaster1.xml"] = _master(parts["ppt/slideMasters/slideMaster1.xml"], spec["theme"])
    parts["[Content_Types].xml"] = _content_types(parts)
    parts["docProps/core.xml"] = _core(parts["docProps/core.xml"], spec["core"])
    parts["docProps/app.xml"] = _app_props(len(pages), sum(1 for p in pages if p["notes"]))
    _dedupe_media(parts)

    tmp = out.with_name(out.name + ".post")
    names = ["[Content_Types].xml"] + [n for n in order if n in parts and n != "[Content_Types].xml" and not n.endswith("/")]
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            data = parts[n]
            if n.startswith("ppt/embeddings/") and n.endswith(".xlsx"):
                data = stable_xlsx(data)
            info = zipfile.ZipInfo(n, FIXED)
            info.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(info, data)
    tmp.replace(out)
    return {"charts": len(chart_specs), "media": sum(1 for n in parts if n.startswith("ppt/media/")), "parts": len(parts)}
