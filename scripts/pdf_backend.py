#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 解析与 PNG 渲染适配器。

优先使用 PyMuPDF（pymupdf）；未安装时自动回退到 Poppler 工具链
（pdftoppm、pdftotext、pdffonts、pdfinfo），保证无 PyMuPDF 的 Linux/macOS 环境也能执行渲染核验。
"""

from __future__ import annotations

import re
import subprocess
from pathlib import Path
from shutil import which
from typing import Any, Dict, List, Optional, Set, Tuple

from lxml import etree


def detect_backend() -> Tuple[Optional[str], str]:
    try:
        import pymupdf  # noqa: F401
        return "pymupdf", ""
    except ImportError:
        pass
    missing = [tool for tool in ("pdftoppm", "pdftotext", "pdffonts", "pdfinfo") if which(tool) is None]
    if not missing:
        return "poppler", ""
    return None, f"缺少 PyMuPDF（pip install pymupdf）且 Poppler 工具不全（缺少 {', '.join(missing)}）；渲染核对跳过"


def font_family_base(name: str) -> str:
    base = name.split("+")[-1]
    for suffix in ("-BoldItalic", "-BoldOblique", "-Bold", "-Italic", "-Oblique", "-Regular",
                   "-Light", "-Medium", "-Semibold", "-SemiBold", "-Black", "-Heavy", "-Thin",
                   "-ExtraBold", "-ExtraLight"):
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    return base.lower().replace(" ", "")


class PdfPageAdapter:
    def save_png(self, png_path: Path, dpi: int) -> None:
        raise NotImplementedError

    @property
    def width_in(self) -> float:
        raise NotImplementedError

    @property
    def height_in(self) -> float:
        raise NotImplementedError

    def extract_text(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        raise NotImplementedError

    def extract_words(self) -> List[Tuple[float, float, float, float, str]]:
        raise NotImplementedError

    def fonts_used(self) -> Set[str]:
        raise NotImplementedError


class _PyMuPDFPage(PdfPageAdapter):
    def __init__(self, page) -> None:
        self._page = page

    def save_png(self, png_path: Path, dpi: int) -> None:
        self._page.get_pixmap(dpi=dpi).save(str(png_path))

    @property
    def width_in(self) -> float:
        return round(self._page.rect.width / 72.0, 3)

    @property
    def height_in(self) -> float:
        return round(self._page.rect.height / 72.0, 3)

    def extract_text(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        lines: List[Dict[str, Any]] = []
        spans: List[Dict[str, Any]] = []
        for block in self._page.get_text("dict")["blocks"]:
            blines = block.get("lines", [])
            for i, line in enumerate(blines):
                text = "".join(span.get("text", "") for span in line.get("spans", []))
                direction = tuple(line.get("dir", (1.0, 0.0)))
                x0, y0, x1, y1 = line["bbox"]
                if text.strip():
                    lines.append({
                        "text": text,
                        "dir": direction,
                        "bbox": (x0 / 72.0, y0 / 72.0, x1 / 72.0, y1 / 72.0),
                        "last": i == len(blines) - 1,
                        "n_lines": len(blines),
                    })
                for span in line.get("spans", []):
                    sx0, sy0, sx1, sy1 = span.get("bbox") or line["bbox"]
                    spans.append({
                        "text": span.get("text", ""),
                        "font": span.get("font", ""),
                        "bbox": (sx0 / 72.0, sy0 / 72.0, sx1 / 72.0, sy1 / 72.0),
                    })
        return lines, spans

    def extract_words(self) -> List[Tuple[float, float, float, float, str]]:
        return [(w[0], w[1], w[2], w[3], w[4]) for w in self._page.get_text("words")]

    def fonts_used(self) -> Set[str]:
        _, spans = self.extract_text()
        return {font_family_base(s["font"]) for s in spans if s["text"].strip() and s.get("font")}


class _PopplerPage(PdfPageAdapter):
    def __init__(self, pdf_path: Path, page_num: int, width_pt: float, height_pt: float,
                 page_el: Optional[etree._Element]) -> None:
        self._pdf = pdf_path
        self._num = page_num
        self._w_pt = width_pt
        self._h_pt = height_pt
        self._el = page_el
        self._fonts_cache: Optional[Set[str]] = None
        self._text_cache: Optional[Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Tuple[float, float, float, float, str]]]] = None

    def save_png(self, png_path: Path, dpi: int) -> None:
        prefix = png_path.with_suffix("")
        subprocess.run(
            ["pdftoppm", "-png", "-r", str(dpi), "-f", str(self._num), "-l", str(self._num),
             "-singlefile", str(self._pdf), str(prefix)],
            capture_output=True, check=True, timeout=60,
        )

    @property
    def width_in(self) -> float:
        return round(self._w_pt / 72.0, 3)

    @property
    def height_in(self) -> float:
        return round(self._h_pt / 72.0, 3)

    def _parse(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Tuple[float, float, float, float, str]]]:
        if self._text_cache is not None:
            return self._text_cache
        lines: List[Dict[str, Any]] = []
        spans: List[Dict[str, Any]] = []
        words: List[Tuple[float, float, float, float, str]] = []
        if self._el is not None:
            ns = "{http://www.w3.org/1999/xhtml}" if self._el.tag.startswith("{") else ""
            for block in self._el.findall(f".//{ns}block"):
                blines = block.findall(f"{ns}line")
                for i, ln in enumerate(blines):
                    lx0 = float(ln.get("xMin", 0))
                    ly0 = float(ln.get("yMin", 0))
                    lx1 = float(ln.get("xMax", 0))
                    ly1 = float(ln.get("yMax", 0))
                    w_els = ln.findall(f"{ns}word")
                    pieces: List[str] = []
                    for w_el in w_els:
                        txt = "".join(w_el.itertext())
                        if not txt:
                            continue
                        wx0 = float(w_el.get("xMin", lx0))
                        wy0 = float(w_el.get("yMin", ly0))
                        wx1 = float(w_el.get("xMax", lx1))
                        wy1 = float(w_el.get("yMax", ly1))
                        pieces.append(txt)
                        words.append((wx0, wy0, wx1, wy1, txt))
                        spans.append({
                            "text": txt,
                            "font": "",
                            "bbox": (wx0 / 72.0, wy0 / 72.0, wx1 / 72.0, wy1 / 72.0),
                        })
                    line_text = " ".join(pieces)
                    if line_text.strip():
                        # Poppler bbox-layout 不直接输出 dir；若宽明显小于高且多字，近似竖排
                        bw, bh = lx1 - lx0, ly1 - ly0
                        direction = (0.0, -1.0) if (bh > bw * 2.2 and len(line_text.strip()) >= 3) else (1.0, 0.0)
                        lines.append({
                            "text": line_text,
                            "dir": direction,
                            "bbox": (lx0 / 72.0, ly0 / 72.0, lx1 / 72.0, ly1 / 72.0),
                            "last": i == len(blines) - 1,
                            "n_lines": len(blines),
                        })
        self._text_cache = (lines, spans, words)
        return self._text_cache

    def extract_text(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        lines, spans, _ = self._parse()
        return lines, spans

    def extract_words(self) -> List[Tuple[float, float, float, float, str]]:
        return self._parse()[2]

    def fonts_used(self) -> Set[str]:
        if self._fonts_cache is not None:
            return self._fonts_cache
        proc = subprocess.run(
            ["pdffonts", "-f", str(self._num), "-l", str(self._num), str(self._pdf)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        fonts: Set[str] = set()
        for idx, line in enumerate(proc.stdout.splitlines()):
            if idx < 2 or not line.strip():
                continue
            name = line.split()[0]
            if name:
                fonts.add(font_family_base(name))
        self._fonts_cache = fonts
        return fonts


class PdfDocAdapter:
    def __init__(self, backend: str, page_count: int, pages: List[PdfPageAdapter]) -> None:
        self.backend = backend
        self.page_count = page_count
        self._pages = pages

    def __iter__(self):
        return iter(self._pages)


def _poppler_page_sizes(pdf_path: Path, count: int) -> List[Tuple[float, float]]:
    proc = subprocess.run(
        ["pdfinfo", "-f", "1", "-l", str(count), str(pdf_path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
    )
    sizes: Dict[int, Tuple[float, float]] = {}
    default_size = (960.0, 540.0)
    for line in proc.stdout.splitlines():
        m = re.match(r"Page\s+(\d+)\s+size:\s+([0-9.]+)\s+x\s+([0-9.]+)\s+pts", line)
        if m:
            sizes[int(m.group(1))] = (float(m.group(2)), float(m.group(3)))
        elif line.startswith("Page size:"):
            m2 = re.search(r"([0-9.]+)\s+x\s+([0-9.]+)\s+pts", line)
            if m2:
                default_size = (float(m2.group(1)), float(m2.group(2)))
    return [sizes.get(i + 1, default_size) for i in range(count)]


def open_pdf(pdf_path: Path) -> PdfDocAdapter:
    backend, reason = detect_backend()
    if backend == "pymupdf":
        import pymupdf
        doc = pymupdf.open(str(pdf_path))
        return PdfDocAdapter("pymupdf", doc.page_count, [_PyMuPDFPage(p) for p in doc])
    if backend == "poppler":
        proc = subprocess.run(
            ["pdftotext", "-bbox-layout", str(pdf_path), "-"],
            capture_output=True, timeout=60, check=True,
        )
        parser = etree.XMLParser(recover=True)
        root = etree.fromstring(proc.stdout, parser=parser)
        ns = "{http://www.w3.org/1999/xhtml}"
        page_els = (root.findall(f".//{ns}page") or root.findall(".//page")) if root is not None else []
        info = subprocess.run(
            ["pdfinfo", str(pdf_path)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        m = re.search(r"Pages:\s+(\d+)", info.stdout)
        count = int(m.group(1)) if m else len(page_els)
        sizes = _poppler_page_sizes(pdf_path, count)
        pages = [
            _PopplerPage(
                pdf_path,
                i + 1,
                sizes[i][0],
                sizes[i][1],
                page_els[i] if i < len(page_els) else None,
            )
            for i in range(count)
        ]
        return PdfDocAdapter("poppler", count, pages)
    raise RuntimeError(reason)
