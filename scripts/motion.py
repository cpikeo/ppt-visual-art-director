#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Animation & Transition Engine（V23）：页面切换 + 对象动画的唯一权威实现。

职责边界（一个功能只有一个实现）：
  契约枚举   plan.schema.json 的 $defs.transition / $defs.animation（本模块导入时与处理表互校，不一致即报错）
  校验       validate_page / validate_deck（plan 层，checks.validate 调用）
  归一化     normalize（plan → spec 的 motion 段；spec.py 调用；含 static / reduced 两档减少动态策略）
  编译       compile_slide（spec motion + 实际 cNvPr@id 目录 → p:transition + p:timing；post.py 与 package.py 调用）
  完整性     integrity_problems（悬空目标、cTn 重复 id、click 锚点、段落/系列越界；吸收 Anthropic
             stranded_animation_targets() 的“目标 id 必须仍在本页”思路并扩展）
  读取/编辑  summarize / remove_targets / renumber（既有稿的动画读取与定点删除；package.py 调用）

只生成 PowerPoint 动画窗格能识别、可调整、可删除的原生时间树；不用 GIF/视频/帧序列冒充。
不能可靠实现的能力（逐词/逐字构建、切换音效、音视频播放、SmartArt 构建）一律 BLOCK，不静默降级。
本模块不依赖 python-pptx。

吸收说明：stranded_animation_targets 的校验思想来自 anthropics/skills（Apache-2.0），此处为独立重写并扩展，见 NOTICE。
"""

from __future__ import annotations

import copy
from typing import Any, Dict, Iterable, List, Optional, Tuple

from lxml import etree

from model import LIMITS, SCHEMA, TEXT_TYPES

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SLIDE_W_IN, SLIDE_H_IN = 12192000 / 914400.0, 6858000 / 914400.0
_NS = {"p": P_NS, "a": A_NS, "mc": MC_NS}

CLASSES = ("entrance", "exit", "emphasis", "path")
# (class, effect) → presetID。presetID 只决定动画窗格里显示的效果名；行为由 cBhvr 子节点决定。
PRESETS: Dict[Tuple[str, str], int] = {
    ("entrance", "appear"): 1, ("entrance", "fade"): 10, ("entrance", "fly"): 2, ("entrance", "wipe"): 22, ("entrance", "zoom"): 53,
    ("exit", "appear"): 1, ("exit", "fade"): 10, ("exit", "fly"): 2, ("exit", "wipe"): 22,
    ("emphasis", "grow_shrink"): 6, ("emphasis", "spin"): 8, ("emphasis", "color"): 19,
    ("path", "motion_path"): 0,
}
PRESET_CLASS = {"entrance": "entr", "exit": "exit", "emphasis": "emph", "path": "path"}
DIR_SUB = {"top": 1, "right": 2, "bottom": 4, "left": 8}
WIPE_FILTER = {"top": "wipe(up)", "bottom": "wipe(down)", "left": "wipe(left)", "right": "wipe(right)"}
CLASS_OF_EFFECT = {"grow_shrink": "emphasis", "spin": "emphasis", "color": "emphasis", "motion_path": "path"}
DEFAULT_DUR = {"appear": 1, "fade": 500, "fly": 500, "wipe": 500, "zoom": 500, "grow_shrink": 500, "spin": 1000,
               "color": 500, "motion_path": 1000}
NODE_TYPE = {"on_click": "clickEffect", "with_previous": "withEffect", "after_previous": "afterEffect"}

# 切换：type → (元素名, 方向属性名 或 None, 允许值)。direction 取 ECMA-376 的原值（运动方向语义）；
# 未经 PowerPoint 实机核对前，不宣称与界面“效果选项”文案逐一对应。
_D4, _D8, _HV, _IO = ("l", "r", "u", "d"), ("l", "u", "r", "d", "lu", "ru", "ld", "rd"), ("horz", "vert"), ("in", "out")
TRANSITIONS: Dict[str, Tuple[str, Optional[str], Tuple[str, ...]]] = {
    "cut": ("cut", None, ()), "fade": ("fade", None, ()), "dissolve": ("dissolve", None, ()),
    "push": ("push", "dir", _D4), "wipe": ("wipe", "dir", _D4), "split": ("split", "dir", _IO),
    "cover": ("cover", "dir", _D8), "uncover": ("pull", "dir", _D8),
    "circle": ("circle", None, ()), "diamond": ("diamond", None, ()), "plus": ("plus", None, ()),
    "wedge": ("wedge", None, ()), "wheel": ("wheel", None, ()), "zoom": ("zoom", "dir", _IO),
    "checker": ("checker", "dir", _HV), "blinds": ("blinds", "dir", _HV), "comb": ("comb", "dir", _HV),
    "random_bar": ("randomBar", "dir", _HV), "newsflash": ("newsflash", None, ()), "random": ("random", None, ()),
}
REDUCED_TRANSITIONS = ("cut", "fade", "dissolve")
REDUCED_EFFECTS = {("entrance", "appear"), ("entrance", "fade"), ("entrance", "wipe"), ("exit", "appear"), ("exit", "fade"), ("exit", "wipe")}


class MotionError(Exception):
    """动画/切换无法可靠生成或目标不成立：BLOCK，不写交付件。"""


# 契约与处理表互校：枚举只在 plan.schema.json 一处声明
if (set(SCHEMA["$defs"]["animation"]["properties"]["effect"]["enum"]) != {e for (_, e) in PRESETS}
        or set(SCHEMA["$defs"]["transition"]["properties"]["type"]["enum"]) != set(TRANSITIONS)):
    raise RuntimeError("plan.schema.json 的动画/切换枚举与 motion.py 处理表不一致")


def pn(tag: str) -> str:
    return "{%s}%s" % (P_NS, tag)


def an(tag: str) -> str:
    return "{%s}%s" % (A_NS, tag)


def _el(tag: str, **attrs) -> etree._Element:
    node = etree.Element(pn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def _sub(parent, tag: str, **attrs) -> etree._Element:
    node = etree.SubElement(parent, pn(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


# ---------------------------------------------------------------- 校验（plan 层）


def effect_class(a: Dict[str, Any]) -> str:
    return a.get("class") or CLASS_OF_EFFECT.get(a["effect"], "entrance")


def _units(el: Dict[str, Any], build: str) -> int:
    """plan 层对构建步数的估计（仅用于点击数告警；编译时以实际 XML 为准）。"""
    if build == "paragraph":
        return max(1, len([ln for ln in str(el.get("content") or "").split("\n") if ln.strip()]))
    data = (el.get("chart") or {}).get("data") or {}
    if build == "series":
        return max(1, len(data.get("series") or []))
    if build == "category":
        return max(1, len(data.get("categories") or []))
    return 1


def click_stages(anims: List[Dict[str, Any]], elems: Dict[str, Dict[str, Any]]) -> int:
    n = 0
    for a in anims:
        u = _units(elems.get(a.get("target"), {}), a.get("build", "whole"))
        n += (1 if a.get("trigger", "on_click") == "on_click" else 0) + (u - 1 if a.get("step_trigger", "on_click") == "on_click" else 0)
    return n


def validate_page(page: Dict[str, Any], top: Dict[str, Dict[str, Any]], nested: Iterable[str], rep) -> None:
    """页级契约：目标引用、字段组合、顺序与克制度。结构类为 HARD，克制度为 SOFT。

    top：本页顶层元素 id → 元素；nested：group 子元素 id（不可直接作为动画目标）。
    """
    pid = page.get("id", "?")
    cfg = LIMITS["motion"]
    anims = page.get("animations") or []
    tr = page.get("transition")
    if tr:
        _, _, allowed = TRANSITIONS[tr["type"]]
        loc = f"{pid}/transition"
        if tr.get("direction") and tr["direction"] not in allowed:
            rep.hard("E-MOTION-FIELD", loc, f"切换 {tr['type']} 的 direction 只能取 {list(allowed) or '无（不接受方向）'}，收到 {tr['direction']!r}")
        if tr["type"] != "wheel" and tr.get("spokes") is not None:
            rep.hard("E-MOTION-FIELD", loc, "spokes 只属于 wheel")
        if tr.get("through_black") and tr["type"] not in ("cut", "fade"):
            rep.hard("E-MOTION-FIELD", loc, "through_black 只属于 cut / fade")
        if tr.get("duration_ms") is not None and tr["type"] == "cut":
            rep.hard("E-MOTION-FIELD", loc, "cut 没有持续时间")
        if tr.get("advance_on_click") is False and tr.get("advance_after_ms") is None:
            rep.hard("E-MOTION-ADVANCE", loc, "禁用点击前进又没有 advance_after_ms：演示会卡在这一页")
        if tr.get("advance_after_ms") == 0 and tr.get("advance_on_click") is False:
            rep.soft("W-MOTION-ADVANCE", loc, "advance_after_ms=0 且禁用点击：立即自动前进，演讲者无法停留")
    if not anims:
        return
    seen: set = set()
    entered: set = set()
    touched: Dict[str, str] = {}
    last_group: Optional[str] = None
    closed_groups: set = set()
    total_ms = 0
    nested = set(nested)
    for i, a in enumerate(anims):
        aid = a.get("id", f"#{i + 1}")
        loc = f"{pid}/{aid}"
        if aid in seen:
            rep.hard("E-MOTION-ID", loc, f"动画 id 在本页重复：{aid}")
        seen.add(aid)
        tgt = a.get("target")
        el = top.get(tgt)
        if el is None:
            where = "是 group 的子元素（PowerPoint 只能对 group 整体编排；改为指向 group）" if tgt in nested else "不是本页元素 id"
            rep.hard("E-MOTION-TARGET", loc, f"动画目标「{tgt}」{where}")
            continue
        cls, eff = effect_class(a), a["effect"]
        if (cls, eff) not in PRESETS:
            rep.hard("E-MOTION-FIELD", loc, f"class={cls} 下没有 effect={eff}；可用：{sorted(e for (c, e) in PRESETS if c == cls)}")
            continue
        own = {"direction": ("fly", "wipe"), "scale": ("grow_shrink",), "degrees": ("spin",), "color": ("color",), "path": ("motion_path",)}
        for field, owners in own.items():
            if a.get(field) is not None and eff not in owners:
                rep.hard("E-MOTION-FIELD", loc, f"字段 {field} 只属于 effect={'/'.join(owners)}，当前是 {eff}")
        if eff == "color" and a.get("color") is None:
            rep.hard("E-MOTION-FIELD", loc, "color 效果必须给出目标 color")
        if eff == "motion_path":
            path = a.get("path")
            if not path:
                rep.hard("E-MOTION-FIELD", loc, "motion_path 必须给出 path")
            else:
                pts = list(path.get("points") or [])
                if path["kind"] == "line":
                    if path.get("dx") is None or path.get("dy") is None:
                        rep.hard("E-MOTION-PATH", loc, "path.kind=line 需要 dx 与 dy（英寸，相对起点）")
                    pts = [{"x": path.get("dx") or 0, "y": path.get("dy") or 0}]
                elif len(pts) < (2 if path["kind"] == "curve" else 1):
                    rep.hard("E-MOTION-PATH", loc, f"path.kind={path['kind']} 的 points 太少")
                if any(abs(p.get("x", 0)) > SLIDE_W_IN * 1.5 or abs(p.get("y", 0)) > SLIDE_H_IN * 1.5 for p in pts):
                    rep.hard("E-MOTION-PATH", loc, "路径偏移超过画布 1.5 倍：对象会飞出很远，检查单位（英寸）")
        build = a.get("build", "whole")
        if build != "whole":
            if cls != "entrance":
                rep.hard("E-MOTION-BUILD", loc, "分步构建（paragraph / series / category）只支持进入类动画")
            if build == "paragraph" and el.get("type") not in TEXT_TYPES:
                rep.hard("E-MOTION-BUILD", loc, "build=paragraph 只能用于文本元素")
            if build in ("series", "category"):
                ch = el.get("chart") or {}
                if el.get("type") != "chart" or ch.get("kind") in (None, "table"):
                    rep.hard("E-MOTION-BUILD", loc, f"build={build} 只能用于原生图表（表格不支持）")
                elif build == "series" and ch.get("kind") in ("pie", "donut"):
                    rep.hard("E-MOTION-BUILD", loc, "饼/环只有一个系列：用 build=category")
        elif a.get("step_trigger") is not None:
            rep.hard("E-MOTION-FIELD", loc, "step_trigger 只在分步构建时有意义")
        if a.get("repeat", 1) > 1 and cls in ("entrance", "exit"):
            rep.hard("E-MOTION-FIELD", loc, "repeat > 1 只对强调类与路径类有意义")
        if a.get("autoreverse") and cls not in ("emphasis", "path"):
            rep.hard("E-MOTION-FIELD", loc, "autoreverse 只属于强调类与路径类")
        if cls == "entrance":
            if tgt in entered:
                rep.soft("W-MOTION-ORDER", loc, f"目标「{tgt}」已有进入动画，再次进入没有意义")
            if tgt in touched:
                rep.hard("E-MOTION-ORDER", loc, f"目标「{tgt}」已先有 {touched[tgt]} 动画，之后不能再做进入动画")
            entered.add(tgt)
        elif touched.get(tgt) == "exit":
            rep.hard("E-MOTION-ORDER", loc, f"目标「{tgt}」已经退出，之后不能再有{cls}动画")
        touched[tgt] = "exit" if cls == "exit" else touched.get(tgt, cls)
        if i == 0 and a.get("trigger", "on_click") != "on_click":
            rep.info("I-MOTION-AUTO", loc, "首个动画不是点击触发：页面一出现就自动播放；确认这是有意的")
        grp = a.get("group")
        if last_group is not None and grp != last_group:
            closed_groups.add(last_group)
        if grp is not None and grp in closed_groups:
            rep.hard("E-MOTION-GROUP", loc, f"group「{grp}」被其他动画打断后又出现：同一 group 必须连续")
        last_group = grp
        total_ms += a.get("delay_ms", 0) + a.get("duration_ms", DEFAULT_DUR[eff]) * a.get("repeat", 1)
        if a.get("duration_ms", 0) > cfg["slow_ms"]:
            rep.soft("W-MOTION-SLOW", loc, f"{a['duration_ms']}ms 过慢（> {cfg['slow_ms']}ms）：除非是路径叙事，否则让观众等待")
    n_click = click_stages(anims, top)
    if n_click > cfg["max_clicks_per_page"]:
        rep.soft("W-MOTION-CLICKS", pid, f"本页有 {n_click} 个单击阶段（建议 ≤ {cfg['max_clicks_per_page']}）：把相关对象并成 group、"
                                         "用 with_previous / after_previous 串起来，不要一个元素一次点击；写不出“每次点击揭示什么”就删掉动画")
    if len(anims) > cfg["max_animations_per_page"]:
        rep.soft("W-MOTION-DENSITY", pid, f"本页 {len(anims)} 个动画（建议 ≤ {cfg['max_animations_per_page']}）：动画是揭示与聚焦的手段，不是装饰")
    if total_ms > cfg["max_total_ms_per_page"]:
        rep.soft("W-MOTION-LENGTH", pid, f"本页动画累计 {total_ms / 1000:.1f}s（建议 ≤ {cfg['max_total_ms_per_page'] / 1000:.0f}s）：演讲节奏被动画牵着走")


def validate_deck(plan: Dict[str, Any], rep) -> None:
    pages = plan.get("pages") or []
    kinds = {p["transition"]["type"] for p in pages if p.get("transition")}
    if len(kinds) > LIMITS["motion"]["max_transition_kinds"]:
        rep.soft("W-MOTION-VARIETY", "整稿", f"整稿用了 {len(kinds)} 种切换 {sorted(kinds)}（建议 ≤ {LIMITS['motion']['max_transition_kinds']}）："
                                           "切换是章节语法，不是页页换花样")
    if len(pages) > 3 and all(p.get("animations") for p in pages):
        rep.soft("W-MOTION-BLANKET", "整稿", "每一页都有对象动画：动画失去强调作用；能用静态构图说清楚的页不要加")


# ---------------------------------------------------------------- 归一化（plan → spec）


def normalize(page: Dict[str, Any], name_of, mode: str = "full") -> Optional[Dict[str, Any]]:
    """页的 transition/animations → spec motion 段。mode：full 原样｜reduced 减少动态｜static 全部去除。

    name_of(element_id) → 成品中的 cNvPr@name。无切换且无动画返回 None（无动画页不产生任何节点）。
    """
    if mode == "static":
        return None
    tr = copy.deepcopy(page.get("transition"))
    anims = []
    for a in page.get("animations") or []:
        cls, eff = effect_class(a), a["effect"]
        item = {"id": a["id"], "target": name_of(a["target"]), "class": cls, "effect": eff,
                "trigger": a.get("trigger", "on_click"), "dur": int(a.get("duration_ms", DEFAULT_DUR[eff])),
                "delay": int(a.get("delay_ms", 0)), "repeat": int(a.get("repeat", 1)), "autoreverse": bool(a.get("autoreverse")),
                "direction": a.get("direction") or ("bottom" if eff == "fly" else "left" if eff == "wipe" else None),
                "build": a.get("build", "whole"), "step_trigger": a.get("step_trigger", "on_click"),
                "scale": a.get("scale", 150), "degrees": a.get("degrees", 360),
                "color": (a.get("color") or "").lstrip("#").upper() or None, "path": a.get("path"), "purpose": a.get("purpose")}
        if mode == "reduced":
            if (cls, eff) not in REDUCED_EFFECTS:
                continue   # 位移/旋转/缩放/路径类：减少动态时整体去除，对象保持可见
            item.update(dur=min(item["dur"], 400) if item["effect"] != "appear" else 1, repeat=1, autoreverse=False)
        anims.append(item)
    if mode == "reduced" and tr:
        if tr["type"] not in REDUCED_TRANSITIONS:
            tr = {"type": "fade", **{k: v for k, v in tr.items() if k in ("advance_on_click", "advance_after_ms")}}
        if tr["type"] == "cut":
            tr.pop("duration_ms", None)
        else:
            tr["duration_ms"] = min(tr.get("duration_ms") or 400, 400)
    if not tr and not anims:
        return None
    return {"transition": tr, "animations": anims}


# ---------------------------------------------------------------- 目录（名称 ↔ 实际 cNvPr@id）


def _cnvpr(node):
    hits = node.xpath("./*[1]/p:cNvPr", namespaces=_NS)
    return hits[0] if hits else None


def catalog(root, chart_info=None) -> Dict[str, Dict[str, Any]]:
    """顶层对象目录：name → {id, kind, paras, has_text, series, cats, chart_kind}；"__nested__" 为 group 内对象 名→id。

    身份以 cNvPr@name 解析到实际 cNvPr@id，不依赖渲染顺序或 shapeIndex；重名视为歧义。
    chart_info(rId) → (n_series, n_categories, plot_tag)，由调用方读取图表部件。
    """
    tree = root.find(".//" + pn("spTree"))
    out: Dict[str, Dict[str, Any]] = {"__nested__": {}}
    dupes: set = set()
    kinds = {pn("sp"): "sp", pn("pic"): "pic", pn("graphicFrame"): "frame", pn("cxnSp"): "cxn", pn("grpSp"): "group"}
    for node in tree:
        kind = kinds.get(node.tag)
        pr = _cnvpr(node) if kind else None
        if pr is None:
            continue
        name = pr.get("name", "")
        info: Dict[str, Any] = {"id": int(pr.get("id")), "kind": kind, "paras": 0, "has_text": False, "series": 0, "cats": 0, "chart_kind": None}
        if kind == "sp":
            info["paras"] = len(node.findall("./" + pn("txBody") + "/" + an("p")))
            info["has_text"] = any((t.text or "").strip() for t in node.iter(an("t")))
        elif kind == "frame":
            ch = node.find(".//{%s}chart" % C_NS)
            if ch is not None and chart_info is not None:
                info["series"], info["cats"], info["chart_kind"] = chart_info(ch.get("{%s}id" % R_NS))
        elif kind == "group":
            out["__nested__"].update({c.get("name"): int(c.get("id")) for c in node.iter(pn("cNvPr")) if c is not pr})
        if name in out:
            dupes.add(name)
        out[name] = info
    for d in dupes:
        out[d]["ambiguous"] = True
    return out


def resolve(cat: Dict[str, Dict[str, Any]], target: str) -> Dict[str, Any]:
    """target = 对象名，或 "#<id>"（既有稿按实际 id 寻址）。"""
    if target.startswith("#") and target[1:].isdigit():
        hits = [v for k, v in cat.items() if k != "__nested__" and v["id"] == int(target[1:])]
    else:
        hits = [cat[target]] if target in cat and target != "__nested__" else []
    if not hits:
        nested = target in cat.get("__nested__", {})
        raise MotionError(f"动画目标「{target}」" + ("在 group 内部：PowerPoint 只能对 group 整体编排" if nested else "在本页找不到（按 cNvPr@name 或 #id 寻址）"))
    if hits[0].get("ambiguous"):
        raise MotionError(f"动画目标「{target}」在本页重名，无法唯一定位")
    return hits[0]


# ---------------------------------------------------------------- 编译：切换


def build_transition(tr: Dict[str, Any]) -> etree._Element:
    """p:transition；有精确持续时间时写 mc:AlternateContent（p14:dur，Fallback 只有 spd），与 PowerPoint 的写法一致。"""
    tag, dattr, _ = TRANSITIONS[tr["type"]]
    dur = tr.get("duration_ms")
    spd = tr.get("speed") or ("slow" if (dur or 0) >= 1000 else "fast" if dur is not None and dur <= 400 else "med")

    def make(with_dur: bool) -> etree._Element:
        t = _el("transition", spd=spd)
        if tr.get("advance_on_click") is False:
            t.set("advClick", "0")
        if tr.get("advance_after_ms") is not None:
            t.set("advTm", str(int(tr["advance_after_ms"])))
        if with_dur and dur is not None:
            t.set("{%s}dur" % P14_NS, str(int(dur)))
        inner = _sub(t, tag)
        if dattr and tr.get("direction"):
            inner.set(dattr, tr["direction"])
        if tr["type"] == "wheel":
            inner.set("spokes", str(tr.get("spokes") or 4))
        if tr.get("through_black"):
            inner.set("thruBlk", "1")
        return t

    if dur is None:
        return make(False)
    ac = etree.Element("{%s}AlternateContent" % MC_NS, nsmap={"mc": MC_NS})
    ch = etree.SubElement(ac, "{%s}Choice" % MC_NS, nsmap={"p14": P14_NS})
    ch.set("Requires", "p14")
    ch.append(make(True))
    etree.SubElement(ac, "{%s}Fallback" % MC_NS).append(make(False))
    return ac


# ---------------------------------------------------------------- 编译：时间树


class _Tree:
    def __init__(self) -> None:
        self.n = 0

    def nid(self) -> int:
        self.n += 1
        return self.n


def _tgt(parent, spid: int, sub: Optional[Dict[str, Any]] = None):
    t = _sub(parent, "tgtEl")
    sp = _sub(t, "spTgt", spid=spid)
    if sub and "para" in sub:
        _sub(_sub(sp, "txEl"), "pRg", st=sub["para"], end=sub["para"])
    elif sub:
        c = etree.SubElement(_sub(sp, "graphicEl"), an("chart"))
        for k, v in sub["chart"].items():
            c.set(k, str(v))
    return t


def _cbhvr(parent, tree: _Tree, spid, sub, dur: int, attr: Optional[Iterable[str]] = None, additive: Optional[str] = None,
           delay: Optional[int] = None, repeat: int = 1, autorev: bool = False):
    cb = _sub(parent, "cBhvr")
    if additive:
        cb.set("additive", additive)
    ct = _sub(cb, "cTn", id=tree.nid(), dur=dur, fill="hold")
    if repeat > 1:
        ct.set("repeatCount", str(repeat * 1000))
    if autorev:
        ct.set("autoRev", "1")
    if delay is not None:
        _sub(_sub(ct, "stCondLst"), "cond", delay=delay)
    _tgt(cb, spid, sub)
    if attr:
        al = _sub(cb, "attrNameLst")
        for name in attr:
            _sub(al, "attrName").text = name
    return cb


def _set_style(parent, tree, spid, sub, name: str, value: str, delay: int = 0) -> None:
    s = _sub(parent, "set")
    _cbhvr(s, tree, spid, sub, 1, [name], delay=delay)
    _sub(_sub(s, "to"), "strVal", val=value)


def _anim_num(parent, tree, spid, sub, attr: str, dur: int, frm: str, to: str) -> None:
    node = _sub(parent, "anim", calcmode="lin", valueType="num")
    _cbhvr(node, tree, spid, sub, dur, [attr], additive="base")
    tl = _sub(node, "tavLst")
    for tm, v in ((0, frm), (100000, to)):
        _sub(_sub(_sub(tl, "tav", tm=tm), "val"), "strVal", val=v)


def _fly_vals(direction: str, enter: bool) -> Tuple[str, str, str]:
    axis = "ppt_x" if direction in ("left", "right") else "ppt_y"
    away = {"left": "0-#ppt_w/2", "right": "1+#ppt_w/2", "top": "0-#ppt_h/2", "bottom": "1+#ppt_h/2"}[direction]
    home = "#" + axis
    return (axis, away, home) if enter else (axis, home, away)


def _bez(points: List[Tuple[float, float]]) -> str:
    """Catmull-Rom → 三次贝塞尔，曲线经过每个点（分数坐标，points[0] 为起点 0,0）。"""
    segs = ["M %.5f %.5f" % points[0]]
    for i in range(len(points) - 1):
        p0, p1, p2, p3 = points[max(i - 1, 0)], points[i], points[i + 1], points[min(i + 2, len(points) - 1)]
        c1 = (p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)
        c2 = (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)
        segs.append("C %.5f %.5f %.5f %.5f %.5f %.5f" % (*c1, *c2, *p2))
    return " ".join(segs) + " E"


def path_string(path: Dict[str, Any]) -> Tuple[str, Optional[str], Tuple[float, float]]:
    """plan 路径（英寸偏移）→ OOXML 路径串（幻灯片宽高的分数）、ptsTypes、rCtr。"""
    fx, fy = 1 / SLIDE_W_IN, 1 / SLIDE_H_IN
    if path["kind"] == "line":
        pts = [(0.0, 0.0), (path["dx"] * fx, path["dy"] * fy)]
    else:
        pts = [(0.0, 0.0)] + [(p["x"] * fx, p["y"] * fy) for p in path["points"]]
    ctr = (pts[-1][0] / 2, pts[-1][1] / 2)
    if path["kind"] == "curve":
        return _bez(pts), None, ctr
    return "M 0 0 " + " ".join("L %.5f %.5f" % p for p in pts[1:]) + " E", "A" * len(pts), ctr


def _behaviors(out, tree: _Tree, a: Dict[str, Any], spid: int, sub) -> None:
    cls, eff, dur = a["class"], a["effect"], a["dur"]
    rep, rev = a["repeat"], a["autoreverse"]
    if cls == "entrance":
        _set_style(out, tree, spid, sub, "style.visibility", "visible")
        if eff == "fade":
            _cbhvr(_sub(out, "animEffect", transition="in", filter="fade"), tree, spid, sub, dur)
        elif eff == "wipe":
            _cbhvr(_sub(out, "animEffect", transition="in", filter=WIPE_FILTER[a["direction"]]), tree, spid, sub, dur)
        elif eff == "fly":
            axis, frm, to = _fly_vals(a["direction"], True)
            _anim_num(out, tree, spid, sub, axis, dur, frm, to)
        elif eff == "zoom":
            _anim_num(out, tree, spid, sub, "ppt_w", dur, "0", "#ppt_w")
            _anim_num(out, tree, spid, sub, "ppt_h", dur, "0", "#ppt_h")
            _cbhvr(_sub(out, "animEffect", transition="in", filter="fade"), tree, spid, sub, dur)
    elif cls == "exit":
        if eff == "fade":
            _cbhvr(_sub(out, "animEffect", transition="out", filter="fade"), tree, spid, sub, dur)
        elif eff == "wipe":
            _cbhvr(_sub(out, "animEffect", transition="out", filter=WIPE_FILTER[a["direction"]]), tree, spid, sub, dur)
        elif eff == "fly":
            axis, frm, to = _fly_vals(a["direction"], False)
            _anim_num(out, tree, spid, sub, axis, dur, frm, to)
        _set_style(out, tree, spid, sub, "style.visibility", "hidden", delay=max(dur - 1, 0))
    elif eff == "grow_shrink":
        s = _sub(out, "animScale")
        _cbhvr(s, tree, spid, sub, dur, repeat=rep, autorev=rev)
        _sub(s, "by", x=int(a["scale"] * 1000), y=int(a["scale"] * 1000))
    elif eff == "spin":
        _cbhvr(_sub(out, "animRot", by=int(a["degrees"] * 60000)), tree, spid, sub, dur, ["r"], repeat=rep, autorev=rev)
    elif eff == "color":
        _set_style(out, tree, spid, sub, "fill.type", "solid")
        _set_style(out, tree, spid, sub, "fill.on", "true")
        c = _sub(out, "animClr", clrSpc="rgb", dir="cw")
        _cbhvr(c, tree, spid, sub, dur, ["fillcolor"], repeat=rep, autorev=rev)
        etree.SubElement(_sub(c, "to"), an("srgbClr"), val=a["color"])
    elif eff == "motion_path":
        path, pts, _ = path_string(a["path"])
        m = _sub(out, "animMotion", origin="layout", path=path, pathEditMode="relative")
        if pts:
            m.set("ptsTypes", pts)
        _cbhvr(m, tree, spid, sub, dur, ["ppt_x", "ppt_y"], repeat=rep, autorev=rev)   # rCtr 可选，省略（ST_Percentage 取值未经实机核对）


def _expand(anims: List[Dict[str, Any]], cat: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """动画 → 效果实例：分步构建展开为每段/每系列/每类别一个实例；步数以实际 XML 为准。"""
    out: List[Dict[str, Any]] = []
    for a in anims:
        info = resolve(cat, a["target"])
        build = a["build"]
        if build == "whole":
            out.append({**a, "spid": info["id"], "sub": None, "info": info})
            continue
        if build == "paragraph":
            if info["kind"] != "sp" or info["paras"] < 2:
                raise MotionError(f"动画 {a['id']}：build=paragraph 需要至少 2 个段落的文本框，目标「{a['target']}」有 {info['paras']} 段")
            subs = [{"para": i} for i in range(info["paras"])]
        else:
            if info["kind"] != "frame" or not info["chart_kind"] or info["chart_kind"] == "table":
                raise MotionError(f"动画 {a['id']}：build={build} 目标不是原生图表")
            if build == "series":
                if info["chart_kind"] in ("pieChart", "doughnutChart"):
                    raise MotionError(f"动画 {a['id']}：饼/环只能按类别构建")
                subs = [{"chart": {"seriesIdx": i, "categoryIdx": -3, "bldStep": "series"}} for i in range(info["series"])]
            else:
                subs = [{"chart": {"seriesIdx": -3, "categoryIdx": j, "bldStep": "category"}} for j in range(info["cats"])]
            if len(subs) < 2:
                raise MotionError(f"动画 {a['id']}：图表只有 {len(subs)} 个{'系列' if build == 'series' else '类别'}，无需分步构建")
        for k, s in enumerate(subs):
            out.append({**a, "spid": info["id"], "sub": s, "info": info,
                        "trigger": a["trigger"] if k == 0 else a["step_trigger"], "delay": a["delay"] if k == 0 else 0})
    return out


def build_timing(anims: List[Dict[str, Any]], cat: Dict[str, Dict[str, Any]]) -> etree._Element:
    """click 组 → 并行段（after_previous 分段）→ 效果（with_previous 同段）；节点 id 在本时间树内顺序唯一。"""
    tree = _Tree()
    inst = _expand(anims, cat)
    timing = _el("timing")
    root_ct = _sub(_sub(_sub(_sub(timing, "tnLst"), "par"), "cTn", id=tree.nid(), dur="indefinite", restart="never", nodeType="tmRoot"), "childTnLst")
    seq = _sub(root_ct, "seq", concurrent="1", nextAc="seek")
    main_id = tree.nid()
    main_children = _sub(_sub(seq, "cTn", id=main_id, dur="indefinite", nodeType="mainSeq"), "childTnLst")
    prev_lst, next_lst = _sub(seq, "prevCondLst"), _sub(seq, "nextCondLst")
    _sub(_sub(_sub(prev_lst, "cond", evt="onPrev", delay="0"), "tgtEl"), "sldTgt")
    _sub(_sub(_sub(next_lst, "cond", evt="onNext", delay="0"), "tgtEl"), "sldTgt")

    st: Dict[str, Any] = {"outer": None, "outer_children": None, "mid_children": None, "start": 0, "end": 0}
    spids: Dict[int, Dict[str, Any]] = {}

    def new_outer(auto: bool) -> None:
        ct = _sub(_sub(main_children, "par"), "cTn", id=tree.nid(), fill="hold")
        sc = _sub(ct, "stCondLst")
        _sub(sc, "cond", delay="indefinite")
        if auto:
            _sub(_sub(sc, "cond", evt="onBegin", delay="0"), "tn", val=main_id)
        st["outer"], st["outer_children"] = ct, _sub(ct, "childTnLst")

    def new_mid(start: int) -> None:
        ct = _sub(_sub(st["outer_children"], "par"), "cTn", id=tree.nid(), fill="hold")
        _sub(_sub(ct, "stCondLst"), "cond", delay=start)
        st["mid_children"], st["start"], st["end"] = _sub(ct, "childTnLst"), start, start

    for e in inst:
        trig = e["trigger"]
        if st["outer"] is None or trig == "on_click":
            new_outer(auto=(trig != "on_click"))
            new_mid(0)
        elif trig == "after_previous":
            new_mid(st["end"])
        direction = e["direction"]
        sub_type = DIR_SUB[direction] if e["effect"] in ("fly", "wipe") else (16 if (e["class"], e["effect"]) == ("entrance", "zoom") else 0)
        ct = _sub(_sub(st["mid_children"], "par"), "cTn", id=tree.nid(), presetID=PRESETS[(e["class"], e["effect"])],
                  presetClass=PRESET_CLASS[e["class"]], presetSubtype=sub_type, fill="hold", grpId=0, nodeType=NODE_TYPE[trig])
        _sub(_sub(ct, "stCondLst"), "cond", delay=e["delay"])
        _behaviors(_sub(ct, "childTnLst"), tree, e, e["spid"], e["sub"])
        st["end"] = max(st["end"], st["start"] + e["delay"] + e["dur"] * e["repeat"] * (2 if e["autoreverse"] else 1))
        spids.setdefault(e["spid"], e)

    bld = None
    for spid, e in spids.items():     # bldLst 登记构建方式（与效果 cTn 的 grpId 对应）
        info = e["info"]
        if info["kind"] == "sp" and info["has_text"]:
            bld = bld if bld is not None else _sub(timing, "bldLst")
            node = _sub(bld, "bldP", spid=spid, grpId=0)
            node.set("build", "p") if e["build"] == "paragraph" else node.set("animBg", "1")
        elif info["kind"] == "frame" and info["chart_kind"]:
            bld = bld if bld is not None else _sub(timing, "bldLst")
            node = _sub(bld, "bldGraphic", spid=spid, grpId=0)
            if e["build"] == "whole":
                _sub(node, "bldAsOne")
            else:
                etree.SubElement(_sub(node, "bldSub"), an("bldChart"), bld=e["build"], animBg="0")
    return timing


def compile_slide(root, motion: Optional[Dict[str, Any]], cat: Dict[str, Dict[str, Any]], replace: bool = True) -> Dict[str, int]:
    """spec motion → p:sld 根；子节点顺序 cSld, clrMapOvr, transition, timing, extLst。写完立即做完整性核对。"""
    stats = {"transition": 0, "animations": 0}
    if replace:
        for old in root.findall(pn("transition")) + root.findall("{%s}AlternateContent" % MC_NS) + root.findall(pn("timing")):
            root.remove(old)
    if not motion:
        return stats
    anchor = root.find(pn("clrMapOvr"))
    anchor = anchor if anchor is not None else root.find(pn("cSld"))
    nodes: List[etree._Element] = []
    if motion.get("transition"):
        nodes.append(build_transition(motion["transition"]))
        stats["transition"] = 1
    if motion.get("animations"):
        nodes.append(build_timing(motion["animations"], cat))
        stats["animations"] = len(motion["animations"])
    for node in reversed(nodes):
        anchor.addnext(node)
    problems = integrity_problems(root, cat)
    if problems:
        raise MotionError("；".join(problems))
    return stats


# ---------------------------------------------------------------- 完整性、读取与定点编辑


def _timing(root):
    t = root.find(pn("timing"))
    return t if t is not None else root.find("{%s}AlternateContent/{%s}Choice/%s" % (MC_NS, MC_NS, pn("timing")))


def integrity_problems(root, cat: Optional[Dict[str, Dict[str, Any]]] = None) -> List[str]:
    """动画/切换引用完整性（PowerPoint 会拒绝或提示修复的结构）。返回空列表才算通过。"""
    out: List[str] = []
    order = ["cSld", "clrMapOvr", "transition", "AlternateContent", "timing", "extLst"]
    idx = [order.index(etree.QName(c).localname) for c in root if isinstance(c.tag, str) and etree.QName(c).localname in order]
    if idx != sorted(idx):
        out.append("E-MOTION-ORDER p:sld 子节点顺序非法（须 cSld, clrMapOvr, transition, timing, extLst）")
    if len(root.findall(pn("timing"))) > 1 or len(root.findall(pn("transition"))) > 1:
        out.append("E-MOTION-DUP 同一页出现多个 p:timing / p:transition")
    timing = _timing(root)
    if timing is None:
        return out
    on_slide = {c.get("id") for c in root.iter(pn("cNvPr"))}
    named = {e.get("spid") for e in timing.iter(pn("spTgt"), pn("bldP"), pn("bldGraphic"))}
    for spid in sorted(named - on_slide - {None}, key=str):
        out.append(f"E-MOTION-STRANDED p:timing 指向已不存在的对象 id={spid}（PowerPoint 会提示修复）")
    ids = [c.get("id") for c in timing.iter(pn("cTn"))]
    dup = sorted({i for i in ids if ids.count(i) > 1}, key=str)
    if dup:
        out.append(f"E-MOTION-CTN cTn id 在同一时间树内重复：{dup}")
    for outer in timing.xpath(".//p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par/p:cTn", namespaces=_NS):
        effects = outer.xpath(".//p:cTn[@presetClass]", namespaces=_NS)
        if not effects:
            out.append("E-MOTION-EMPTY 存在没有任何效果的 click 组")
            continue
        auto = bool(outer.xpath("./p:stCondLst/p:cond[@evt='onBegin']", namespaces=_NS))
        if not auto and effects[0].get("nodeType") != "clickEffect":
            out.append("E-MOTION-CLICK click 组的第一个效果不是 clickEffect（点击锚点丢失）")
    if cat:
        by_id = {v["id"]: v for k, v in cat.items() if k != "__nested__"}
        for tgt in timing.iter(pn("spTgt")):
            info = by_id.get(int(tgt.get("spid"))) if (tgt.get("spid") or "").isdigit() else None
            if info is None:
                continue
            for pr in tgt.iter(pn("pRg")):
                if int(pr.get("end", 0)) >= max(info["paras"], 1) or int(pr.get("st", 0)) > int(pr.get("end", 0)):
                    out.append(f"E-MOTION-RANGE 段落范围 {pr.get('st')}–{pr.get('end')} 超出目标的 {info['paras']} 个段落")
            for ch in tgt.iter(an("chart")):
                s, c = int(ch.get("seriesIdx", -3)), int(ch.get("categoryIdx", -3))
                if s >= info["series"] or c >= info["cats"]:
                    out.append(f"E-MOTION-RANGE 图表构建索引 series={s} category={c} 超出（{info['series']} 系列 / {info['cats']} 类别）")
    return out


def summarize(root, cat: Optional[Dict[str, Dict[str, Any]]] = None) -> Dict[str, Any]:
    """读取一页现有的切换与动画（动画窗格顺序）。供 verify（计划 ↔ 成品对账）与 package inspect 使用。"""
    names = {str(v["id"]): k for k, v in (cat or {}).items() if k != "__nested__"}
    tr = None
    node = root.find(pn("transition"))
    if node is None:
        node = root.find("{%s}AlternateContent/{%s}Choice/%s" % (MC_NS, MC_NS, pn("transition")))
    if node is not None:
        inner = next((c for c in node if isinstance(c.tag, str)), None)
        local = etree.QName(inner).localname if inner is not None else None
        kind = next((k for k, (tag, _, _) in TRANSITIONS.items() if tag == local), local)
        dur = node.get("{%s}dur" % P14_NS)
        tr = {"type": kind, "direction": inner.get("dir") if inner is not None else None, "speed": node.get("spd"),
              "duration_ms": int(dur) if dur else None, "advance_on_click": node.get("advClick", "1") != "0",
              "advance_after_ms": int(node.get("advTm")) if node.get("advTm") else None}
    effects: List[Dict[str, Any]] = []
    timing = _timing(root)
    if timing is not None:
        for ct in timing.iter(pn("cTn")):
            if ct.get("presetClass") is None:
                continue
            tgt = next(iter(ct.iter(pn("spTgt"))), None)
            spid = tgt.get("spid") if tgt is not None else None
            durs = [int(c.get("dur")) for c in ct.iter(pn("cTn")) if (c.get("dur") or "").isdigit()]
            delay = ct.find("./" + pn("stCondLst") + "/" + pn("cond"))
            sub = None
            if tgt is not None and tgt.find(pn("txEl")) is not None:
                sub = "p%s" % tgt.find(".//" + pn("pRg")).get("st")
            elif tgt is not None and tgt.find(pn("graphicEl")) is not None:
                c = tgt.find(".//" + an("chart"))
                sub = c.get("bldStep") + ":" + (c.get("seriesIdx") if c.get("bldStep") == "series" else c.get("categoryIdx"))
            effects.append({"spid": spid, "name": names.get(spid), "class": ct.get("presetClass"), "preset": int(ct.get("presetID", -1)),
                            "trigger": ct.get("nodeType"), "dur_ms": max(durs or [0]), "sub": sub,
                            "delay_ms": int(delay.get("delay")) if delay is not None and (delay.get("delay") or "").isdigit() else 0})
    return {"transition": tr, "effects": effects}


def renumber(root) -> None:
    """时间树 cTn id 按文档序重排为 1..n（同步 onBegin 的 tn 引用）。删除/复制效果后保持唯一。"""
    timing = _timing(root)
    if timing is None:
        return
    mapping: Dict[str, str] = {}
    for i, ct in enumerate(timing.iter(pn("cTn")), 1):
        mapping[ct.get("id")] = str(i)
        ct.set("id", str(i))
    for tn in timing.iter(pn("tn")):
        tn.set("val", mapping.get(tn.get("val"), tn.get("val")))


def remap_ids(root, mapping: Dict[int, int]) -> int:
    """对象 id 被改动时，同步改写 p:timing 中的 spid（先整体映射再写回，避免链式覆盖）。"""
    timing = _timing(root)
    n = 0
    if timing is None:
        return n
    for node in timing.iter(pn("spTgt"), pn("bldP"), pn("bldGraphic")):
        old = node.get("spid")
        if old and old.isdigit() and int(old) in mapping:
            node.set("spid", str(mapping[int(old)]))
            n += 1
    return n


def remove_targets(root, spids: Iterable[int]) -> int:
    """删除对象时清理其动画：移除以这些 spid 为目标的效果节点与 bld 登记；click 锚点必要时顺延；空时间树整体移除。

    返回移除的效果数。效果节点指向多个对象等无法安全拆解的情形抛 MotionError，不静默丢失。
    """
    gone = {str(s) for s in spids}
    timing = _timing(root)
    if timing is None:
        return 0
    removed = 0
    for ct in list(timing.iter(pn("cTn"))):
        if ct.get("presetClass") is None or ct.getparent() is None:
            continue
        tg = {t.get("spid") for t in ct.iter(pn("spTgt"))}
        if not tg & gone:
            continue
        if len(tg) > 1:
            raise MotionError("一个效果节点指向多个对象，无法安全拆解；请先在 PowerPoint 动画窗格中拆开")
        par = ct.getparent()
        par.getparent().remove(par)
        removed += 1
    for b in list(timing.iter(pn("bldP"), pn("bldGraphic"))):
        if b.get("spid") in gone:
            b.getparent().remove(b)
    for ct in list(timing.xpath(".//p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par/p:cTn", namespaces=_NS)):
        if not ct.xpath(".//p:cTn[@presetClass]", namespaces=_NS):
            ct.getparent().getparent().remove(ct.getparent())
            continue
        for mid in list(ct.xpath("./p:childTnLst/p:par", namespaces=_NS)):
            if not mid.xpath(".//p:cTn[@presetClass]", namespaces=_NS):
                mid.getparent().remove(mid)
        first = ct.xpath(".//p:cTn[@presetClass]", namespaces=_NS)[0]
        if not ct.xpath("./p:stCondLst/p:cond[@evt='onBegin']", namespaces=_NS) and first.get("nodeType") != "clickEffect":
            first.set("nodeType", "clickEffect")
    if not timing.xpath(".//p:cTn[@presetClass]", namespaces=_NS):
        timing.getparent().remove(timing)
    else:
        for bl in timing.findall(pn("bldLst")):
            if len(bl) == 0:
                timing.remove(bl)
        renumber(root)
    return removed
