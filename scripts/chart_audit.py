#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""图表成品的有效色审计：只读 chart XML，回答“读者实际看到的字，落在什么表面上、够不够读”。

和 verify 里“成品 == plan”的核对不同：那类核对用的是编译器自己的映射，两边一起错也会通过。
这里不读 plan 的任何颜色，只读成品：

  标签墨色（XML 里写了什么） × 标签落点（dLblPos 或该类型的默认位置）
      → 落点表面（承载面 / 该点的系列或扇区色 / 标签衬底）
      → WCAG 对比度，低于 4.5 即不可读。

同时核对 dLblPos 对图表类型是否合法（donut、area 不允许；堆叠柱不允许 outEnd）——
XSD 只认枚举，认不出这条，但 PowerPoint 会提示“需要修复”。
"""

from __future__ import annotations

from typing import List, Optional, Tuple

from lxml import etree

from model import INSIDE_POSITIONS, LABEL_INK_MIN, LIMITS, allowed_label_positions, contrast_ratio, gray_conflicts, qn

Finding = Tuple[str, str, str]  # (HARD|SOFT|INFO, code, message)

_PLOTS = ("barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart")
_DEFAULT_POS = {"lineChart": "r", "scatterChart": "r", "areaChart": "ctr", "doughnutChart": "ctr", "pieChart": "bestFit"}


def _srgb(node) -> Optional[str]:
    hit = None if node is None else node.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
    return None if hit is None else "#" + (hit.get("val") or "").upper()


def _val(node, tag: str, default: Optional[str] = None) -> Optional[str]:
    hit = None if node is None else node.find(qn(tag))
    return default if hit is None else hit.get("val")


def _ink(txpr) -> Optional[str]:
    """txPr 里的墨色：defRPr 优先，其次 run 属性。"""
    if txpr is None:
        return None
    for tag in ("a:defRPr", "a:rPr"):
        rpr = txpr.find(".//" + qn(tag))
        if rpr is not None and _srgb(rpr):
            return _srgb(rpr)
    return None


def _on(node, tag: str) -> bool:
    return _val(node, tag, "0") in ("1", "true")


def _shown(lbl) -> bool:
    return lbl is not None and _val(lbl, "c:delete", "0") not in ("1", "true") and any(
        _on(lbl, t) for t in ("c:showVal", "c:showCatName", "c:showPercent", "c:showSerName"))


def _ser_name(ser) -> str:
    tx = ser.find(qn("c:tx"))
    node = None if tx is None else tx.find(".//" + qn("c:v"))
    return node.text if node is not None and node.text else ""


def marker_of(ser) -> dict:
    """成品里一个系列的 marker：symbol、size、填充色、边线色（只读 XML；缺省返回 None 值）。"""
    m = ser.find(qn("c:marker"))
    if m is None:
        return {"symbol": None, "size": None, "fill": None, "line": None}
    sp = m.find(qn("c:spPr"))
    size = _val(m, "c:size")
    return {"symbol": _val(m, "c:symbol"), "size": int(size) if size and size.isdigit() else None,
            "fill": _srgb(sp), "line": _srgb(sp.find(qn("a:ln")) if sp is not None else None)}


def _fill(ser, plot_tag: str, idx: Optional[int]) -> Optional[str]:
    """点（或系列）实际的填充色：折线取线色，散点取 marker 填充（线被关掉了），其余取实心填充；逐点 dPt 优先。"""
    if idx is not None:
        for dpt in ser.findall(qn("c:dPt")):
            if _val(dpt, "c:idx") == str(idx):
                got = marker_of(dpt)["fill"] if plot_tag == "scatterChart" else _srgb(dpt.find(qn("c:spPr")))
                if got:
                    return got
    if plot_tag == "scatterChart":
        mk = marker_of(ser)
        if mk["symbol"] == "none":
            return None
        return mk["fill"] or mk["line"]
    sp = ser.find(qn("c:spPr"))
    if sp is None:
        return None
    if plot_tag == "lineChart":
        return _srgb(sp.find(qn("a:ln")))
    return _srgb(sp)


def _label_for(ser, plot_dlbls, idx: int):
    """某数据点上生效的标签：逐点 dLbl > 系列 dLbls > plot 级 dLbls。返回（标签节点, 容器节点）。"""
    sd = ser.find(qn("c:dLbls"))
    if sd is None:
        return plot_dlbls, plot_dlbls
    for d in sd.findall(qn("c:dLbl")):
        if _val(d, "c:idx") == str(idx):
            return d, sd
    return sd, sd


def _count(ser) -> int:
    for holder in ("c:val", "c:yVal"):
        node = ser.find(qn(holder))
        cache = None if node is None else node.find(".//" + qn("c:numCache"))
        if cache is not None:
            return len(cache.findall(qn("c:pt")))
    return 0


def _plots(chart_space):
    chart = chart_space.find(qn("c:chart"))
    area = None if chart is None else chart.find(qn("c:plotArea"))
    if area is None:
        return []
    return [p for p in area if etree.QName(p).localname in _PLOTS]


def audit_labels(chart_space, surface: str) -> List[Finding]:
    out: List[Finding] = []
    for plot in _plots(chart_space):
        tag = etree.QName(plot).localname
        grouping = _val(plot, "c:grouping", "clustered") or "clustered"
        allowed = allowed_label_positions(tag, grouping)
        plot_dlbls = plot.find(qn("c:dLbls"))
        # dLblPos 的合法性看每一处写入（plot / 系列 / 逐点），不看“谁覆盖了谁”：PowerPoint 逐个元素校验。
        illegal = sorted({n.get("val") for n in plot.iter(qn("c:dLblPos"))
                          if allowed is None or n.get("val") not in allowed})
        if illegal:
            legal = "该类型不允许写 dLblPos" if allowed is None else f"只允许 {sorted(allowed)}"
            out.append(("HARD", "V-LABEL-POS",
                        f"{tag}（{grouping}）里 dLblPos={illegal} 非法：{legal}；PowerPoint 打开会提示修复"))
        # 折线叠在柱/面之上时，t 位的标签可能压在色块里，不能只按承载面判读
        under = [_fill(s, etree.QName(p).localname, None) for p in _plots(chart_space)
                 if etree.QName(p).localname in ("barChart", "areaChart") and tag == "lineChart"
                 for s in p.findall(qn("c:ser"))]
        for ser in plot.findall(qn("c:ser")):
            name, seen = _ser_name(ser), set()
            for idx in range(max(_count(ser), 1)):
                lbl, owner = _label_for(ser, plot_dlbls, idx)
                if not _shown(lbl):
                    continue
                pos_node = lbl.find(qn("c:dLblPos"))
                pos = None if pos_node is None else pos_node.get("val")
                if pos is not None:
                    eff = pos
                elif tag == "barChart":
                    eff = "ctr" if grouping in ("stacked", "percentStacked") else "outEnd"
                else:
                    eff = _DEFAULT_POS[tag]
                ink = _ink(lbl.find(qn("c:txPr"))) or _ink(owner.find(qn("c:txPr")) if owner is not None else None)
                if ink is None:
                    if "ink" not in seen:
                        seen.add("ink")
                        out.append(("SOFT", "V-LABEL-INK", f"系列「{name}」直标没有显式墨色，实际取决于主题默认色"))
                    continue
                plate, fill = _srgb(lbl.find(qn("c:spPr"))), _fill(ser, tag, idx)
                if plate:
                    surfaces = [(plate, "标签衬底")]
                elif eff in INSIDE_POSITIONS:
                    surfaces = [(fill, "系列/扇区色块")]
                elif eff == "bestFit":  # 饼的 bestFit 里外由渲染器定：两种表面都要读得清
                    surfaces = [(surface, "承载面"), (fill, "扇区色块")]
                else:
                    surfaces = [(surface, "承载面")] + [(c, "其下的柱/面色块") for c in under if c]
                for color, where in surfaces:
                    if color is None:
                        if ("unk", where) not in seen:
                            seen.add(("unk", where))
                            out.append(("INFO", "I-LABEL", f"系列「{name}」直标落在{where}，其色不是显式 RGB，对比度未核对"))
                        continue
                    ratio = contrast_ratio(ink, color)
                    if ratio < LABEL_INK_MIN and ("c", ink, color) not in seen:
                        seen.add(("c", ink, color))
                        out.append(("HARD", "V-LABEL-INK",
                                    f"系列「{name}」直标墨色 {ink} 落在{where} {color} 上仅 {ratio:.2f}:1"
                                    f"（< {LABEL_INK_MIN}），读不清"))
    return out


def audit_structure(chart_space, surface: str) -> List[Finding]:
    """轴、图例、图题的墨色相对承载面。"""
    out: List[Finding] = []
    chart = chart_space.find(qn("c:chart"))
    if chart is None:
        return out
    items: List[Tuple[str, Optional[str]]] = []
    legend = chart.find(qn("c:legend"))
    if legend is not None:
        items.append(("图例", _ink(legend.find(qn("c:txPr")))))
    for ax in chart.iter(qn("c:catAx"), qn("c:valAx")):
        if _val(ax, "c:delete", "0") not in ("1", "true"):
            items.append(("轴标签", _ink(ax.find(qn("c:txPr")))))
    title = chart.find(qn("c:title"))
    if title is not None:
        rpr = title.find(".//" + qn("a:rPr"))
        items.append(("图题", _srgb(rpr)))
    for what, ink in items:
        if ink is None:
            out.append(("SOFT", "V-LABEL-INK", f"{what}没有显式墨色，实际取决于主题默认色"))
            continue
        ratio = contrast_ratio(ink, surface)
        if ratio < LABEL_INK_MIN:
            out.append(("HARD", "V-LABEL-INK", f"{what}墨色 {ink} 对承载面 {surface} 仅 {ratio:.2f}:1（< {LABEL_INK_MIN}）"))
    return out


def audit_series_surface(chart_space, surface: str, kind: str) -> List[Finding]:
    """系列色相对承载面；阈值按图表类型取自 limits.json。"""
    out: List[Finding] = []
    floor = LIMITS["chart"]["min_contrast"].get(kind, 2.0)
    seen = set()
    for plot in _plots(chart_space):
        tag = etree.QName(plot).localname
        for ser in plot.findall(qn("c:ser")):
            colors = [(_fill(ser, tag, None), _ser_name(ser))]
            colors += [(_fill(ser, tag, int(_val(d, "c:idx") or 0)), f"{_ser_name(ser)}#{_val(d, 'c:idx')}") for d in ser.findall(qn("c:dPt"))]
            for color, label in colors:
                if not color or color in seen:
                    continue
                seen.add(color)
                ratio = contrast_ratio(color, surface)
                if ratio < floor:
                    out.append(("SOFT", "V-CONTRAST",
                                f"「{label}」色 {color} 对承载面 {surface} 仅 {ratio:.2f}:1（< {floor}），图形会沉进底里"))
    return out


def audit_gray(chart_space, kind: str) -> List[Finding]:
    """成品里各系列的实际颜色，转灰度后两两比较（只读 XML，不读 plan）。

    饼/环的扇区有直标或图例，且彼此以承载面色描边隔开，不按系列灰度判；单系列图没有可比对象。
    """
    if kind in ("pie", "donut", "table"):
        return []
    named: List[Tuple[str, str]] = []
    for plot in _plots(chart_space):
        tag = etree.QName(plot).localname
        for ser in plot.findall(qn("c:ser")):
            color = _fill(ser, tag, None)
            if color:
                named.append((_ser_name(ser) or f"#{len(named) + 1}", color))
    if len(named) < 2:
        return []
    pairs = gray_conflicts([c for _, c in named])
    if not pairs:
        return []
    shown = "；".join(f"「{named[i][0]}」{named[i][1]} 与「{named[j][0]}」{named[j][1]}（ΔL* {d:.1f}）" for i, j, d in pairs[:3])
    return [("SOFT", "V-GRAY", f"成品里 {len(pairs)} 对系列转灰度后难以区分：{shown}；"
                               "需有颜色之外的区分（直标系列名、线型、标记），否则投影与黑白输出会读错")]


def audit_readout(chart_space, kind: str, legend_element: bool = False, legend_shows_value: bool = False) -> List[Finding]:
    """饼/环成品的读数闭环：只看 XML 里真实存在的图例与数据标签。

    类别来源 = 原生 c:legend、showCatName（仅饼；环内直标只有数值）或页上 legend 元素；数值来源 = showVal 或带读数的 legend 元素。
    """
    if kind not in ("pie", "donut"):
        return []
    chart = chart_space.find(qn("c:chart"))
    native_legend = chart is not None and chart.find(qn("c:legend")) is not None
    cat_on = val_on = False
    for plot in _plots(chart_space):
        for node in plot.iter(qn("c:dLbls"), qn("c:dLbl")):
            if _shown(node):
                cat_on = cat_on or _on(node, "c:showCatName")
                val_on = val_on or _on(node, "c:showVal")
    names_ok = native_legend or legend_element or (kind == "pie" and cat_on)
    out: List[Finding] = []
    if not names_ok:
        out.append(("HARD", "V-READOUT", f"成品{kind}没有任何类别读数（无图例、无类别名直标）：读者认不出每个扇区是什么"))
    if not (val_on or legend_shows_value):
        out.append(("SOFT", "V-READOUT", f"成品{kind}没有数值读数：只能凭扇区面积估数"))
    return out


def audit_markers(chart_space, kind: str) -> List[Finding]:
    """marker 的成品属性：散点的数据点就是 marker，看不见等于数据不存在；折线的 marker 须与线同色。

    只读 XML。散点 marker 色对承载面的对比与灰度由 audit_series_surface / audit_gray 经 _fill 覆盖；
    散点颜色等不等于 plan 的调色板由 verify 另核；这里回答“marker 在不在、够不够大、是否连了线”。
    """
    out: List[Finding] = []
    for plot in _plots(chart_space):
        tag = etree.QName(plot).localname
        if tag not in ("scatterChart", "lineChart"):
            continue
        for ser in plot.findall(qn("c:ser")):
            name, mk = _ser_name(ser), marker_of(ser)
            if tag == "scatterChart":
                color = mk["fill"] or mk["line"]
                if mk["symbol"] == "none" or color is None:
                    out.append(("HARD", "V-MARKER", f"散点系列「{name}」成品没有可见 marker（symbol={mk['symbol']}，填充={mk['fill']}）：数据点不会出现在图上"))
                    continue
                if mk["size"] is not None and mk["size"] < 5:
                    out.append(("SOFT", "V-MARKER", f"散点系列「{name}」marker 仅 {mk['size']}pt，投影下点会读不到"))
                ln = ser.find(qn("c:spPr") + "/" + qn("a:ln"))
                if ln is None or ln.find(qn("a:noFill")) is None:
                    out.append(("SOFT", "V-MARKER", f"散点系列「{name}」没有关闭连线：连线会暗示不存在的顺序"))
            elif mk["symbol"] not in (None, "none"):
                line = _srgb(ser.find(qn("c:spPr") + "/" + qn("a:ln")) if ser.find(qn("c:spPr")) is not None else None)
                if mk["fill"] and line and mk["fill"] != line:
                    out.append(("SOFT", "V-MARKER", f"折线「{name}」marker 色 {mk['fill']} ≠ 线色 {line}：读者会把点与线看成两个系列"))
    return out


def audit_chart(chart_space, surface: str, kind: str) -> List[Finding]:
    return (audit_labels(chart_space, surface) + audit_structure(chart_space, surface)
            + audit_series_surface(chart_space, surface, kind) + audit_gray(chart_space, kind)
            + audit_markers(chart_space, kind))
