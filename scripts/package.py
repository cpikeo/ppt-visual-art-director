#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPTX Package Engine（V23）：既有稿 / 模板的结构级、保真操作。不经 python-pptx 重存整份文件。

子命令（输出一律 .pptx，先写临时件、officecheck 通过后原子落盘；输入与输出不得同文件）：
  inspect  DECK [--text] [--json]          只读：页序、版式、对象 id/名称、动画/切换、媒体、备注、主题；--text 抽取文本/表格/图表数据/替代文字/备注
  unpack   DECK DIR                        安全解包（拒绝越界部件名、超限体积）
  pack     DIR OUT                         重新打包（[Content_Types].xml 居首、固定时间戳）
  to-pptx  TEMPLATE.potx OUT               模板 → 演示文稿（只改主部件内容类型）
  dup      DECK OUT --slide N [--after M]  复制页：图表/工作簿/备注/tags/SmartArt/OLE 深拷贝，图片/媒体/版式共享；动画随页保留
  delete   DECK OUT --slide N              删除页并清理孤立部件（含分节引用）
  move     DECK OUT --slide N --to M       重排页序（含分节的稿子 BLOCK）
  remove-shape DECK OUT --slide N --target NAME|#ID   删除对象，同步清理其动画目标与 bld 登记；孤立资源一并清理
  motion   DECK OUT --slide N (--set motion.json [--replace] | --clear)   既有稿定点写入/清除切换与动画
  clean    DECK OUT                        清理孤立页面、媒体、图表与关系，修正 Content Types

吸收自 anthropics/skills pptx（Apache-2.0）的思路：add_slide.py 的页注册与 tags 独占、clean.py 的可达性清理、
slide_text.py 的文本抽取范围、soffice/validate 的“先验后落盘”。均为独立重写，见 NOTICE。
画布固定 16:9（12192000×6858000 EMU）：非 16:9 稿的结构操作一律 BLOCK（inspect 只报告）。
退出码：0 成功｜1 校验未通过（交付路径未写）｜2 输入/契约错误。
"""

from __future__ import annotations

import argparse
import copy
import json
import posixpath
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

import motion
import officecheck as oc
from model import CANVAS_EMU, LIMITS, Report, guard_paths, package_problems, parse_xml, place_atomic, temp_beside, utf8_stdio

P, A, R = motion.P_NS, motion.A_NS, motion.R_NS
REL_NS, CT_NS = oc.REL_NS, oc.CT_NS
FIXED = (2000, 1, 1, 0, 0, 0)
SHARED_TYPES = ("/slideLayout", "/slideMaster", "/theme", "/notesMaster", "/image", "/video", "/audio", "/media", "/hyperlink", "/font",
                "/tableStyles", "/presProps", "/viewProps")
PRUNABLE = ("/image", "/chart", "/diagramData", "/diagramLayout", "/diagramQuickStyle", "/diagramColors", "/diagramDrawing", "/oleObject",
            "/package", "/video", "/audio", "/media", "/hyperlink", "/tags")
SLIDE_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
NOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
MAIN_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"


class PackageEngineError(Exception):
    """无法安全执行的结构操作：BLOCK，不写交付件。"""


def _ser(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Pkg:
    """内存中的 OPC 包：部件字典 + 原始顺序。所有结构操作都在这里，保证单一实现。"""

    def __init__(self, parts: Dict[str, bytes]) -> None:
        self.parts = dict(parts)
        self.order = list(parts)

    @classmethod
    def load(cls, path: Path, need_canvas: bool = True) -> "Pkg":
        problems = package_problems(path, check_suffix=False)
        if problems:
            raise PackageEngineError("包不合格：" + "；".join(problems))
        pkg = cls(oc.read_parts(path))
        main = pkg.content_type("ppt/presentation.xml")
        if main not in (MAIN_CT, TEMPLATE_CT):
            raise PackageEngineError(f"主部件内容类型 {main!r} 不是 .pptx/.potx（宏稿 .pptm 一律拒绝）")
        if need_canvas:
            sz = parse_xml(pkg.parts["ppt/presentation.xml"]).find(motion.pn("sldSz"))
            if (int(sz.get("cx")), int(sz.get("cy"))) != tuple(CANVAS_EMU):
                raise PackageEngineError(f"E-CANVAS 画布 {sz.get('cx')}×{sz.get('cy')} 不是 16:9（{CANVAS_EMU[0]}×{CANVAS_EMU[1]} EMU）；本 Skill 只支持 16:9")
        return pkg

    # ---- 内容类型
    def _ct(self):
        return parse_xml(self.parts["[Content_Types].xml"])

    def content_type(self, part: str) -> Optional[str]:
        root = self._ct()
        for e in root.findall("{%s}Override" % CT_NS):
            if e.get("PartName") == "/" + part:
                return e.get("ContentType")
        ext = part.rsplit(".", 1)[-1].lower()
        for e in root.findall("{%s}Default" % CT_NS):
            if e.get("Extension", "").lower() == ext:
                return e.get("ContentType")
        return None

    def set_override(self, part: str, ctype: Optional[str]) -> None:
        root = self._ct()
        for e in root.findall("{%s}Override" % CT_NS):
            if e.get("PartName") == "/" + part:
                root.remove(e)
        if ctype:
            etree.SubElement(root, "{%s}Override" % CT_NS, PartName="/" + part, ContentType=ctype)
        self.parts["[Content_Types].xml"] = _ser(root)

    # ---- 关系
    def rels(self, part: str) -> List[Dict[str, Any]]:
        return oc.part_rels(self.parts, part)

    def write_rels(self, part: str, rels: List[Dict[str, Any]]) -> None:
        name = oc.rels_name(part)
        root = etree.Element("{%s}Relationships" % REL_NS, nsmap={None: REL_NS})
        for r in rels:
            node = etree.SubElement(root, "{%s}Relationship" % REL_NS, Id=r["id"], Type=r["type"], Target=r["target"])
            if r.get("external"):
                node.set("TargetMode", "External")
        self.parts[name] = _ser(root)
        if name not in self.order:
            self.order.append(name)

    def rel_target(self, source: str, target_part: str) -> str:
        return posixpath.relpath(target_part, posixpath.dirname(source) or ".")

    # ---- 部件
    def new_name(self, part: str) -> str:
        d, base = posixpath.split(part)
        m = re.fullmatch(r"(.*?)(\d+)(\.[^.]+)", base)
        stem, ext = (m.group(1), m.group(3)) if m else (base.rsplit(".", 1)[0], "." + base.rsplit(".", 1)[-1])
        n = max([int(x.group(1)) for k in self.parts for x in [re.fullmatch(re.escape(posixpath.join(d, stem)) + r"(\d+)" + re.escape(ext), k)] if x] + [0]) + 1
        return posixpath.join(d, f"{stem}{n}{ext}")

    def clone_part(self, part: str, memo: Dict[str, str], keep_shared: bool = True) -> str:
        """深拷贝一个部件及其私有子部件（图表→工作簿/样式等）；共享类型（版式、母版、图片、媒体…）只引用不拷贝。"""
        if part in memo:
            return memo[part]
        new = self.new_name(part)
        memo[part] = new
        self.parts[new] = self.parts[part]
        self.order.append(new)
        ctype = self.content_type(part)
        if any(e.get("PartName") == "/" + part for e in self._ct().findall("{%s}Override" % CT_NS)):
            self.set_override(new, ctype)
        rels = self.rels(part)
        out = []
        for r in rels:
            if r["external"] or any(r["type"].endswith(t) for t in SHARED_TYPES):
                out.append({"id": r["id"], "type": r["type"], "target": r["target"], "external": r["external"]})
            else:
                child = self.clone_part(r["part"], memo)
                out.append({"id": r["id"], "type": r["type"], "target": self.rel_target(new, child), "external": False})
        if rels:
            self.write_rels(new, out)
        return new

    # ---- 页
    def slides(self) -> List[Dict[str, Any]]:
        return oc.listed_slides(self.parts)

    def _pres(self):
        return parse_xml(self.parts["ppt/presentation.xml"])

    def _save_pres(self, root) -> None:
        self.parts["ppt/presentation.xml"] = _ser(root)

    def slide(self, n: int) -> Dict[str, Any]:
        s = self.slides()
        if not 1 <= n <= len(s):
            raise PackageEngineError(f"页号 {n} 超出范围 1–{len(s)}")
        return s[n - 1]

    def has_sections(self) -> bool:
        return bool(self._pres().xpath(".//*[local-name()='sectionLst']"))

    def dup_slide(self, n: int, after: Optional[int] = None) -> str:
        src = self.slide(n)
        memo: Dict[str, str] = {}
        new_slide = self.new_name(src["part"])
        memo[src["part"]] = new_slide
        self.parts[new_slide] = self.parts[src["part"]]
        self.order.append(new_slide)
        self.set_override(new_slide, self.content_type(src["part"]))
        rels = []
        for r in self.rels(src["part"]):
            if r["external"] or any(r["type"].endswith(t) for t in SHARED_TYPES):
                rels.append({"id": r["id"], "type": r["type"], "target": r["target"], "external": r["external"]})
            elif r["type"].endswith("/notesSlide"):
                note = self.clone_part(r["part"], memo)
                rels.append({"id": r["id"], "type": r["type"], "target": self.rel_target(new_slide, note), "external": False})
                back = [{**x, "target": (self.rel_target(note, new_slide) if x["part"] == src["part"] else x["target"])} for x in self.rels(note)]
                self.write_rels(note, back)
            else:
                child = self.clone_part(r["part"], memo)
                rels.append({"id": r["id"], "type": r["type"], "target": self.rel_target(new_slide, child), "external": False})
        self.write_rels(new_slide, rels)
        # 注册：presentation.xml.rels + sldIdLst（新 id = 最大 id + 1）+ 分节
        prels = self.rels("ppt/presentation.xml")
        rid = "rId%d" % (max([int(re.sub(r"\D", "", r["id"]) or 0) for r in prels] + [0]) + 1)
        prels.append({"id": rid, "type": SLIDE_RT, "target": self.rel_target("ppt/presentation.xml", new_slide), "external": False})
        self.write_rels("ppt/presentation.xml", prels)
        pres = self._pres()
        lst = pres.find(motion.pn("sldIdLst"))
        new_id = str(max([int(e.get("id")) for e in lst] + [255]) + 1)
        node = etree.Element(motion.pn("sldId"), id=new_id)
        node.set("{%s}id" % R, rid)
        pos = (after if after is not None else len(lst))
        if after is None:
            pos = n   # 默认紧跟源页
        lst.insert(pos, node)
        for sec in pres.xpath(".//*[local-name()='sectionLst']//*[local-name()='sldId']"):
            if sec.get("id") == src["sld_id"]:
                clone = etree.Element(sec.tag, id=new_id)
                sec.addnext(clone)
        self._save_pres(pres)
        self._bump_app()
        return new_slide

    def delete_slide(self, n: int) -> None:
        s = self.slide(n)
        pres = self._pres()
        lst = pres.find(motion.pn("sldIdLst"))
        for e in list(lst):
            if e.get("id") == s["sld_id"]:
                lst.remove(e)
        for sec in pres.xpath(".//*[local-name()='sectionLst']//*[local-name()='sldId']"):
            if sec.get("id") == s["sld_id"]:
                sec.getparent().remove(sec)
        self._save_pres(pres)
        self.write_rels("ppt/presentation.xml", [r for r in self.rels("ppt/presentation.xml") if r["id"] != s["rid"]])
        self.clean()
        self._bump_app()

    def move_slide(self, n: int, to: int) -> None:
        if self.has_sections():
            raise PackageEngineError("E-SECTION 本稿含分节（p14:sectionLst）：移动页会改变分节归属，请在 PowerPoint 内完成")
        self.slide(n)
        pres = self._pres()
        lst = pres.find(motion.pn("sldIdLst"))
        if not 1 <= to <= len(lst):
            raise PackageEngineError(f"目标位置 {to} 超出范围 1–{len(lst)}")
        node = list(lst)[n - 1]
        lst.remove(node)
        lst.insert(to - 1, node)
        self._save_pres(pres)

    def _bump_app(self) -> None:
        if "docProps/app.xml" in self.parts:
            txt = self.parts["docProps/app.xml"].decode("utf-8")
            txt = re.sub(r"<Slides>\d+</Slides>", "<Slides>%d</Slides>" % len(self.slides()), txt)
            self.parts["docProps/app.xml"] = txt.encode("utf-8")

    # ---- 清理（可达性）
    def reachable(self) -> set:
        seen: set = set()
        stack = ["_rels/.rels"]
        roots = [r for r in oc.part_rels(self.parts, "")] if False else []
        root_rels = parse_xml(self.parts["_rels/.rels"])
        stack = [oc.resolve_target("", r.get("Target")) for r in root_rels if (r.get("TargetMode") or "").lower() != "external"]
        while stack:
            p = stack.pop()
            if p in seen or p not in self.parts:
                continue
            seen.add(p)
            for r in self.rels(p):
                if r["part"]:
                    stack.append(r["part"])
        return seen

    def clean(self) -> List[str]:
        keep = self.reachable()
        removed = []
        for name in list(self.parts):
            if name in ("[Content_Types].xml", "_rels/.rels"):
                continue
            if name.endswith(".rels"):
                owner = posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
                if owner not in keep:
                    del self.parts[name]
                    removed.append(name)
                continue
            if name not in keep:
                del self.parts[name]
                removed.append(name)
        root = self._ct()
        for e in root.findall("{%s}Override" % CT_NS):
            if e.get("PartName").lstrip("/") not in self.parts:
                root.remove(e)
        self.parts["[Content_Types].xml"] = _ser(root)
        self.order = [n for n in self.order if n in self.parts]
        return removed

    def prune_slide_rels(self, part: str) -> None:
        """页 XML 不再引用的资源关系（图片、图表、媒体…）摘掉，clean 随后清掉孤立部件。"""
        xml = self.parts[part].decode("utf-8")
        keep = [r for r in self.rels(part) if not any(r["type"].endswith(t) for t in PRUNABLE) or ('"%s"' % r["id"]) in xml]
        if len(keep) != len(self.rels(part)):
            self.write_rels(part, keep)

    # ---- 写出
    def write(self, path: Path) -> None:
        names = ["[Content_Types].xml"] + [n for n in self.order if n in self.parts and n != "[Content_Types].xml"]
        names += [n for n in self.parts if n not in names]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in names:
                info = zipfile.ZipInfo(n, FIXED)
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, self.parts[n])


# ---------------------------------------------------------------- 对象 / 动画 编辑


def _find_nodes(root, target: str):
    tree = root.find(".//" + motion.pn("spTree"))
    out = []
    for node in tree:
        pr = motion._cnvpr(node) if isinstance(node.tag, str) else None
        if pr is None:
            continue
        if (target.startswith("#") and pr.get("id") == target[1:]) or pr.get("name") == target:
            out.append(node)
    return out


def remove_shape(pkg: Pkg, n: int, target: str) -> Dict[str, int]:
    part = pkg.slide(n)["part"]
    root = parse_xml(pkg.parts[part])
    hits = _find_nodes(root, target)
    if not hits:
        raise PackageEngineError(f"第 {n} 页没有顶层对象 {target!r}（按 cNvPr@name 或 #id；group 内对象须先在 PowerPoint 中解组）")
    if len(hits) > 1:
        raise PackageEngineError(f"{target!r} 重名，无法唯一定位；改用 #id")
    ids = [int(c.get("id")) for c in hits[0].iter(motion.pn("cNvPr"))]
    removed = motion.remove_targets(root, ids)
    hits[0].getparent().remove(hits[0])
    pkg.parts[part] = _ser(root)
    pkg.prune_slide_rels(part)
    cleaned = pkg.clean()
    return {"animations_removed": removed, "parts_cleaned": len(cleaned)}


def _shim_top(cat: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """用实际 XML 目录给 motion.validate_page 造一个“plan 元素”视图（只含顺序/字段/构建规则需要的类型信息）。"""
    top = {}
    for name, info in cat.items():
        if name == "__nested__":
            continue
        if info["kind"] == "frame" and info["chart_kind"]:
            kind = {"pieChart": "pie", "doughnutChart": "donut"}.get(info["chart_kind"], "bar")
            top[name] = {"type": "chart", "chart": {"kind": kind, "data": {"series": [0] * info["series"], "categories": [0] * info["cats"]}}}
        elif info["kind"] == "sp" and info["has_text"]:
            top[name] = {"type": "body", "content": "\n".join("x" for _ in range(max(info["paras"], 1)))}
        else:
            top[name] = {"type": "shape" if info["kind"] != "frame" else "table"}
        top["#%d" % info["id"]] = top[name]
    return top


def set_motion(pkg: Pkg, n: int, spec: Dict[str, Any], replace: bool, rep: Report) -> Dict[str, int]:
    import jsonschema
    part = pkg.slide(n)["part"]
    root = parse_xml(pkg.parts[part])
    cat = motion.catalog(root, oc._chart_info(pkg.parts, part, pkg.rels(part)))
    if set(spec) - {"transition", "animations"}:
        raise PackageEngineError("motion.json 只能有 transition 与 animations 两个键")
    sch = {"$ref": "#/$defs/page_motion", "$defs": {**__import__("model").SCHEMA["$defs"],
           "page_motion": {"type": "object", "properties": {"transition": {"$ref": "#/$defs/transition"}, "animations": {"type": "array", "items": {"$ref": "#/$defs/animation"}}}}}}
    errs = sorted(jsonschema.Draft202012Validator(sch).iter_errors(spec), key=lambda e: list(e.path))
    if errs:
        raise PackageEngineError("motion.json 不符合契约：" + "；".join(f"{'/'.join(map(str, e.path))}: {e.message}" for e in errs[:4]))
    existing = motion.summarize(root, cat)
    timing = motion._timing(root)
    if timing is not None and (timing.xpath(".//p:cTn[@nodeType='interactiveSeq']", namespaces=motion._NS) or timing.xpath(".//p:cmd", namespaces=motion._NS)):
        raise PackageEngineError("E-MOTION-BLOCK 本页含交互触发序列（interactiveSeq）或媒体命令（p:cmd），本引擎无法安全改写；请在 PowerPoint 中修改")
    if (existing["effects"]) and not replace:
        raise PackageEngineError(f"E-MOTION-EXIST 本页已有 {len(existing['effects'])} 个动画效果；确认整体替换请加 --replace（不静默覆盖原稿动画）")
    page = {"id": f"slide{n}", **spec}
    tops = _shim_top(cat)
    motion.validate_page(page, tops, cat.get("__nested__", {}), rep)
    if rep.blocked:
        raise PackageEngineError("E-MOTION 动画契约未通过：" + "；".join(f"{i.code} {i.message}" for i in rep.items if i.level == "HARD"))
    norm = motion.normalize(page, lambda t: t, "full")
    try:
        stats = motion.compile_slide(root, norm, cat, replace=True)
    except motion.MotionError as exc:
        raise PackageEngineError(f"E-MOTION {exc}")
    pkg.parts[part] = _ser(root)
    return stats


def clear_motion(pkg: Pkg, n: int) -> None:
    part = pkg.slide(n)["part"]
    root = parse_xml(pkg.parts[part])
    motion.compile_slide(root, None, {}, replace=True)
    pkg.parts[part] = _ser(root)


# ---------------------------------------------------------------- inspect / text


def _txt(node) -> str:
    return "".join(t.text or "" for t in node.iter("{%s}t" % A))


def slide_text(pkg: Pkg, n: int) -> List[str]:
    """文本、表格、图表数据、图片替代文字、备注：抽取范围对齐 Anthropic slide_text.py，按声明顺序。"""
    s = pkg.slide(n)
    root = parse_xml(pkg.parts[s["part"]])
    rels = {r["id"]: r for r in pkg.rels(s["part"])}
    out: List[str] = []
    for node in root.find(".//" + motion.pn("spTree")).iter(motion.pn("sp"), motion.pn("pic"), motion.pn("graphicFrame"), motion.pn("cxnSp")):
        pr = motion._cnvpr(node)
        name = pr.get("name") if pr is not None else "?"
        if node.tag == motion.pn("sp"):
            paras = [_txt(p) for p in node.findall("./" + motion.pn("txBody") + "/{%s}p" % A)]
            if any(paras):
                out.append(f"[{name}] " + " ⏎ ".join(p for p in paras if p))
        elif node.tag == motion.pn("pic"):
            out.append(f"[{name}] 图片 alt={pr.get('descr') or '(无)'}")
        elif node.tag == motion.pn("graphicFrame"):
            tbl = node.find(".//{%s}tbl" % A)
            ch = node.find(".//{%s}chart" % motion.C_NS)
            if tbl is not None:
                for tr in tbl.findall("{%s}tr" % A):
                    out.append(f"[{name}] 表 | " + " | ".join(_txt(tc) for tc in tr.findall("{%s}tc" % A)))
            elif ch is not None and rels.get(ch.get("{%s}id" % R)) and rels[ch.get("{%s}id" % R)]["part"] in pkg.parts:
                cr = parse_xml(pkg.parts[rels[ch.get("{%s}id" % R)]["part"]])
                cns = {"c": motion.C_NS}
                for ser in cr.findall(".//c:ser", cns):
                    nm = "".join(v.text or "" for v in ser.findall("./c:tx//c:v", cns))
                    cats = [v.text for v in ser.findall("./c:cat//c:v", cns)]
                    vals = [v.text for v in ser.findall("./c:val//c:v", cns) + ser.findall("./c:yVal//c:v", cns)]
                    out.append(f"[{name}] 图表系列 {nm}: " + ", ".join(f"{c}={v}" for c, v in zip(cats or range(len(vals)), vals)))
    for r in pkg.rels(s["part"]):
        if r["type"].endswith("/notesSlide") and r["part"] in pkg.parts:
            nr = parse_xml(pkg.parts[r["part"]])
            body = [sp for sp in nr.iter(motion.pn("sp")) if sp.xpath(".//p:ph[@type='body']", namespaces=motion._NS)]
            txt = " ⏎ ".join(_txt(p) for sp in body for p in sp.findall(".//{%s}p" % A) if _txt(p))
            if txt:
                out.append("[备注] " + txt)
    return out


def inspect(path: Path, text: bool = False) -> Dict[str, Any]:
    pkg = Pkg.load(path, need_canvas=False)
    pres = pkg._pres()
    sz = pres.find(motion.pn("sldSz"))
    theme_name = next((n for n in pkg.parts if re.fullmatch(r"ppt/theme/theme\d+\.xml", n)), None)
    theme: Dict[str, Any] = {}
    if theme_name:
        t = parse_xml(pkg.parts[theme_name])
        ns = {"a": A}
        theme = {"name": t.get("name"), "major": (t.find(".//a:majorFont/a:latin", ns).get("typeface") if t.find(".//a:majorFont/a:latin", ns) is not None else None),
                 "minor": (t.find(".//a:minorFont/a:latin", ns).get("typeface") if t.find(".//a:minorFont/a:latin", ns) is not None else None),
                 "colors": {c.tag.split("}")[1]: (c[0].get("val") or c[0].get("lastClr")) for c in t.find(".//a:clrScheme", ns) if len(c)}}
    layouts = []
    for name in sorted(n for n in pkg.parts if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", n)):
        lr = parse_xml(pkg.parts[name])
        layouts.append({"part": name, "name": lr.find(motion.pn("cSld")).get("name"), "type": lr.get("type"),
                        "placeholders": [{"type": ph.get("type") or "body", "idx": ph.get("idx")} for ph in lr.iter(motion.pn("ph"))]})
    slides = []
    for i, s in enumerate(pkg.slides(), 1):
        root = parse_xml(pkg.parts[s["part"]])
        rels = pkg.rels(s["part"])
        cat = motion.catalog(root, oc._chart_info(pkg.parts, s["part"], rels))
        item = {"index": i, "part": s["part"], "sld_id": s["sld_id"], "hidden": root.get("show") == "0",
                "layout": next((r["part"] for r in rels if r["type"].endswith("/slideLayout")), None),
                "objects": [{"id": v["id"], "name": k, "kind": v["kind"], "paras": v["paras"], "series": v["series"]} for k, v in cat.items() if k != "__nested__"],
                "motion": motion.summarize(root, cat),
                "media": [r["part"] for r in rels if any(r["type"].endswith(t) for t in ("/video", "/audio", "/media"))],
                "smartart": any(r["type"].endswith("/diagramData") for r in rels),
                "has_notes": any(r["type"].endswith("/notesSlide") for r in rels)}
        if text:
            item["text"] = slide_text(pkg, i)
        slides.append(item)
    return {"file": str(path), "main_content_type": pkg.content_type("ppt/presentation.xml"), "is_template": pkg.content_type("ppt/presentation.xml") == TEMPLATE_CT,
            "canvas_emu": [int(sz.get("cx")), int(sz.get("cy"))], "canvas_16_9": (int(sz.get("cx")), int(sz.get("cy"))) == tuple(CANVAS_EMU),
            "masters": sum(1 for n in pkg.parts if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", n)), "layouts": layouts, "theme": theme,
            "sections": pkg.has_sections(), "slides": slides}


# ---------------------------------------------------------------- 解包 / 打包


def safe_unpack(src: Path, dest: Path) -> int:
    problems = package_problems(src, check_suffix=False)
    if problems:
        raise PackageEngineError("包不合格：" + "；".join(problems))
    if dest.exists() and any(dest.iterdir()):
        raise PackageEngineError(f"{dest} 非空，拒绝解包")
    dest.mkdir(parents=True, exist_ok=True)
    n = 0
    for name, data in oc.read_parts(src).items():
        target = (dest / name).resolve()
        if dest.resolve() not in target.parents:
            raise PackageEngineError(f"越界部件名：{name}")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        n += 1
    return n


def pack_dir(src: Path, out: Path, force: bool) -> None:
    parts = {p.relative_to(src).as_posix(): p.read_bytes() for p in sorted(src.rglob("*")) if p.is_file()}
    if "[Content_Types].xml" not in parts:
        raise PackageEngineError("目录里没有 [Content_Types].xml")
    tmp = temp_beside(out)
    try:
        Pkg(parts).write(tmp)
        _finish(tmp, out, None, force)
    finally:
        tmp.exists() and tmp.unlink()


def _finish(tmp: Path, out: Path, baseline: Optional[Path], force: bool) -> Dict[str, Any]:
    """先验后落盘：officecheck（含基线对比）不通过，交付路径不写。"""
    res = oc.check(tmp, baseline)
    if res["status"] != "PASS":
        raise PackageEngineError("校验未通过：" + "；".join(f"{p['code']} {p['where']}: {p['message']}" for p in res["problems"] if p["code"].startswith("E-"))[:900])
    place_atomic(tmp, out, force)
    return res


# ---------------------------------------------------------------- CLI


def main(argv: Optional[List[str]] = None) -> int:
    utf8_stdio()
    ap = argparse.ArgumentParser(description="PPTX Package Engine：既有稿 / 模板的结构级保真操作")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("inspect"); s.add_argument("deck"); s.add_argument("--text", action="store_true"); s.add_argument("--json", action="store_true")
    s = sub.add_parser("unpack"); s.add_argument("deck"); s.add_argument("dir")
    s = sub.add_parser("pack"); s.add_argument("dir"); s.add_argument("out")
    s = sub.add_parser("to-pptx"); s.add_argument("deck"); s.add_argument("out")
    for name in ("dup", "delete", "move", "remove-shape", "motion", "clean"):
        s = sub.add_parser(name); s.add_argument("deck"); s.add_argument("out")
        if name != "clean":
            s.add_argument("--slide", type=int, required=True)
    sub.choices["dup"].add_argument("--after", type=int)
    sub.choices["move"].add_argument("--to", type=int, required=True)
    sub.choices["remove-shape"].add_argument("--target", required=True)
    sub.choices["motion"].add_argument("--set"); sub.choices["motion"].add_argument("--clear", action="store_true")
    sub.choices["motion"].add_argument("--replace", action="store_true")
    for c in sub.choices.values():
        if c.prog.split()[-1] not in ("inspect", "unpack"):
            c.add_argument("--force", action="store_true")
    args = ap.parse_args(argv)
    try:
        if args.cmd == "inspect":
            data = inspect(Path(args.deck), args.text)
            if args.json:
                print(json.dumps(data, ensure_ascii=False, indent=2))
            else:
                print(f"{Path(args.deck).name}｜{'模板' if data['is_template'] else '演示文稿'}｜16:9={data['canvas_16_9']}｜{len(data['slides'])} 页｜母版 {data['masters']}｜版式 {len(data['layouts'])}｜分节 {data['sections']}")
                for sl in data["slides"]:
                    m = sl["motion"]
                    print(f"  {sl['index']:>2}. {sl['part']} id={sl['sld_id']} 对象 {len(sl['objects'])}｜切换 {m['transition']['type'] if m['transition'] else '-'}｜动画 {len(m['effects'])}"
                          f"{'｜SmartArt' if sl['smartart'] else ''}{'｜媒体 %d' % len(sl['media']) if sl['media'] else ''}{'｜备注' if sl['has_notes'] else ''}")
                    for line in sl.get("text", []):
                        print("      " + line)
            return 0
        if args.cmd == "unpack":
            print(f"已解包 {safe_unpack(Path(args.deck), Path(args.dir))} 个部件 → {args.dir}")
            return 0
        rep = Report()
        src = Path(args.deck if args.cmd != "pack" else args.dir)
        out = Path(args.out)
        if args.cmd == "pack":
            pack_dir(src, out, args.force)
            print(f"已打包 → {out}")
            return 0
        guard_paths(rep, {"输入": src}, [("输出", out, ".pptx")], args.force)
        if rep.blocked:
            for i in rep.items:
                print(f"[HARD] {i.code} {i.message}")
            return 2
        pkg = Pkg.load(src, need_canvas=args.cmd != "to-pptx")
        note = ""
        if args.cmd == "to-pptx":
            if pkg.content_type("ppt/presentation.xml") != TEMPLATE_CT:
                raise PackageEngineError("输入不是 .potx 模板")
            pkg.set_override("ppt/presentation.xml", MAIN_CT)
        elif args.cmd == "dup":
            note = pkg.dup_slide(args.slide, args.after)
        elif args.cmd == "delete":
            pkg.delete_slide(args.slide)
        elif args.cmd == "move":
            pkg.move_slide(args.slide, args.to)
        elif args.cmd == "remove-shape":
            note = json.dumps(remove_shape(pkg, args.slide, args.target), ensure_ascii=False)
        elif args.cmd == "clean":
            note = f"清理 {len(pkg.clean())} 个孤立部件"
        elif args.cmd == "motion":
            if bool(args.set) == bool(args.clear):
                raise PackageEngineError("motion 需要 --set motion.json 或 --clear 二选一")
            if args.clear:
                if pkg.slide(args.slide) and motion._timing(parse_xml(pkg.parts[pkg.slide(args.slide)['part']])) is not None and not args.replace:
                    raise PackageEngineError("E-MOTION-EXIST 清除会移除原稿动画；确认请加 --replace")
                clear_motion(pkg, args.slide)
            else:
                spec = json.loads(Path(args.set).read_text(encoding="utf-8-sig"))
                note = json.dumps(set_motion(pkg, args.slide, spec, args.replace, rep))
        tmp = temp_beside(out)
        try:
            pkg.write(tmp)
            res = _finish(tmp, out, src, args.force)
        finally:
            tmp.exists() and tmp.unlink()
        for i in rep.items:
            print(f"[{i.level}] {i.code} {i.where}: {i.message}")
        print(f"已写出 {out}｜{args.cmd}{('｜' + note) if note else ''}｜结构层 {res['status']}｜PowerPoint 实机：NOT VERIFIED")
        return 0
    except PackageEngineError as exc:
        print(f"[HARD] E-PACKAGE {exc}")
        return 1
    except FileExistsError as exc:
        print(f"[HARD] E-EXIST {exc}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
