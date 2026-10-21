#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""assets.py · 素材的确定性度量（只测量，不判断好坏）。

边界：
  * 这里只回答「这张图多少像素、这块区域的亮度/纹理是多少、对比度多少」。
  * 「够不够、能不能压字、该不该保留」由 guard.py 按 policy 判定 —— 度量和规则分开，
    否则改一条阈值就要动度量代码。
  * 不做美学评价：均值亮度与标准差不是「好看」，只是「文字能不能读」的可测前提。
"""
from __future__ import annotations

import hashlib
from pathlib import Path

MAX_PIXELS = 60_000_000        # 兜底常量：仅当调用方没有策略对象时使用（生产路径读 assets.pixels.max_decode）


class AssetError(RuntimeError):
    pass


def resolve_source(source, base_dir) -> Path | None:
    """素材路径：绝对路径直接用；相对路径相对 brief 所在目录解析。"""
    if not source:
        return None
    path = Path(str(source)).expanduser()
    if not path.is_absolute() and base_dir:
        path = Path(base_dir) / path
    return path if path.is_file() else None


def digest(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _decode_budget(policy) -> int:
    """单张图解码预算来自策略（assets.pixels.max_decode）；没有策略对象才用兜底常量。"""
    if policy is None:
        return MAX_PIXELS
    return int(policy.req("assets.pixels.max_decode"))


def inspect(path: Path, policy=None) -> dict:
    """读图的客观事实：像素尺寸、格式、字节摘要。不解码整幅到内存之外。

    `policy` 给了就从 `assets.pixels.max_decode` 取解码预算——预算是产品决定，不是代码常量。
    """
    from PIL import Image
    path = Path(path)
    try:
        with Image.open(path) as image:
            width, height = image.size
            image_format = (image.format or path.suffix.lstrip(".")).upper()
            budget = _decode_budget(policy)
            if width * height > budget:
                raise AssetError(f"素材像素数 {width}×{height} 超过上限 {budget}")
    except AssetError:
        raise
    except (OSError, ValueError) as exc:
        raise AssetError(f"素材无法解码：{path.name} — {exc}") from None
    return {"path": str(path), "width": int(width), "height": int(height),
            "long_side": max(width, height), "format": image_format,
            "bytes": path.stat().st_size, "sha256": digest(path)}


def load(path: Path, max_side: int | None = None):
    """解码为 RGB，必要时按长边采样（区域统计不需要原始分辨率）。"""
    from PIL import Image
    with Image.open(path) as image:
        image = image.convert("RGB")
        if max_side and max(image.size) > max_side:
            ratio = max_side / max(image.size)
            image = image.resize((max(1, int(image.width * ratio)), max(1, int(image.height * ratio))),
                                 Image.BILINEAR)
        return image.copy()


def luma(rgb) -> float:
    """相对亮度（0..1），按 sRGB 系数；用于对比度与亮度差。"""
    r, g, b = (channel / 255.0 for channel in rgb[:3])

    def linear(value: float) -> float:
        return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4

    return 0.2126 * linear(r) + 0.7152 * linear(g) + 0.0722 * linear(b)


def contrast_ratio(luma_a: float, luma_b: float) -> float:
    light, dark = max(luma_a, luma_b), min(luma_a, luma_b)
    return (light + 0.05) / (dark + 0.05)


def region_stats(image, box_ratio, policy) -> dict:
    """测量「文字安全区」对应的图像区域：均值亮度、纹理标准差、平坦区占比。

    box_ratio: (x, y, w, h) 以图像为基准的比例坐标。
    """
    if image is None:
        return {}
    width, height = image.size
    x = max(0, min(width - 1, int(box_ratio[0] * width)))
    y = max(0, min(height - 1, int(box_ratio[1] * height)))
    w = max(1, int(box_ratio[2] * width))
    h = max(1, int(box_ratio[3] * height))
    crop = image.crop((x, y, min(width, x + w), min(height, y + h)))
    if max(crop.size) > int(policy.req("assets.text_over_image.sample_max_side")):
        ratio = int(policy.req("assets.text_over_image.sample_max_side")) / max(crop.size)
        crop = crop.resize((max(1, int(crop.width * ratio)), max(1, int(crop.height * ratio))))
    grey = crop.convert("L")
    pixels = list(grey.tobytes())          # 8-bit 灰度：字节即像素，避免弃用 API
    total = len(pixels) or 1
    mean = sum(pixels) / total / 255.0
    variance = sum((value / 255.0 - mean) ** 2 for value in pixels) / total
    std = variance ** 0.5
    flat_std = float(policy.req("assets.qc.flat_std"))
    # 平坦区占比：按 16×16 块统计，块内标准差低于阈值即视为平坦（可用作留白）
    blocks, flat_blocks = 0, 0
    block = 16
    for by in range(0, grey.height, block):
        for bx in range(0, grey.width, block):
            tile = grey.crop((bx, by, min(grey.width, bx + block), min(grey.height, by + block)))
            values = list(tile.tobytes())
            if len(values) < 4:
                continue
            blocks += 1
            tile_mean = sum(values) / len(values) / 255.0
            tile_std = (sum((v / 255.0 - tile_mean) ** 2 for v in values) / len(values)) ** 0.5
            if tile_std < flat_std:
                flat_blocks += 1
    return {"mean_luma": round(mean, 4), "std": round(std, 4),
            "flat_ratio": round(flat_blocks / blocks, 4) if blocks else 0.0}


def cover_crop(source_px, target_px) -> dict:
    """cover 时被裁掉多少：返回每边裁掉的像素比例与「损失倍数」。

    损失倍数 = 需要的可视面积 / 实际保留的可视面积；> 1 表示有内容被裁掉。
    """
    sw, sh = source_px
    tw, th = target_px
    if not sw or not sh or not tw or not th:
        return {"scale": 1.0, "loss": 1.0, "crop": [0.0, 0.0, 0.0, 0.0]}
    source_ratio, target_ratio = sw / sh, tw / th
    if abs(source_ratio - target_ratio) < 1e-6:
        return {"scale": 1.0, "loss": 1.0, "crop": [0.0, 0.0, 0.0, 0.0]}
    if source_ratio > target_ratio:       # 源更宽：左右被裁
        keep = target_ratio / source_ratio
        side = (1 - keep) / 2
        return {"scale": keep, "loss": round(1 / keep, 3), "crop": [side, 0.0, side, 0.0]}
    keep = source_ratio / target_ratio    # 源更高：上下被裁
    side = (1 - keep) / 2
    return {"scale": keep, "loss": round(1 / keep, 3), "crop": [0.0, side, 0.0, side]}


def required_source_px(box_px, policy) -> int:
    """放置尺寸需要的源像素长边：放置长边 × oversample，且不低于策略下限。"""
    long_side = max(float(box_px[0]), float(box_px[1]))
    oversample = float(policy.req("assets.oversample"))
    floor = int(policy.req("assets.min_long_side_px"))
    return int(max(floor, round(long_side * oversample)))
