#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Explicit, tested migrations for persisted PVAD diagnostic documents."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parent.parent


class MigrationError(ValueError):
    pass


def migrate_inspection(document: dict) -> dict:
    version = document.get("schema")
    if version == "pvad-inspection-v2":
        return document
    if version != "pvad-inspection-v1":
        raise MigrationError(f"不支持的 inspection schema：{version!r}")
    slides = document.get("slides") or []
    images = missing_alt = 0
    hyperlinks, fonts = [], set()
    inherited = 0
    for slide in slides:
        slide.setdefault("layout", "")
        slide.setdefault("notes_present", False)
        for obj in slide.get("objects") or []:
            kind = obj.get("type")
            if kind == "image":
                images += 1
                if not obj.get("alt_text"):
                    missing_alt += 1
            if kind == "text" and obj.get("text"):
                obj_fonts = obj.get("fonts") or []
                fonts.update(obj_fonts)
                if not obj_fonts:
                    inherited += 1
                hyperlinks.extend({"slide": slide.get("index"), "object": obj.get("name", ""), **link}
                                  for link in obj.get("hyperlinks", []))
    result = dict(document)
    result["schema"] = "pvad-inspection-v2"
    result["slides"] = slides
    result.setdefault("accessibility", {
        "images": images,
        "images_missing_alt": [{"migration": "count-only", "count": missing_alt}] if missing_alt else [],
        "hyperlinks": hyperlinks,
        "slides_with_notes": sum(bool(slide.get("notes_present")) for slide in slides),
        "reading_order": "legacy document: z_order may be unavailable; human review required",
    })
    result.setdefault("font_diagnostics", {
        "explicit_fonts": sorted(fonts),
        "inherited_font_objects": inherited,
        "target_environment_match": "未测量",
    })
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Migrate persisted PVAD documents")
    parser.add_argument("kind", choices=["inspection"])
    parser.add_argument("source")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    try:
        source = Path(args.source)
        document = json.loads(source.read_text(encoding="utf-8"))
        result = migrate_inspection(document)
        schema = json.loads((ROOT / "schema" / "inspection.schema.json").read_text(encoding="utf-8"))
        Draft202012Validator(schema).validate(result)
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except (OSError, json.JSONDecodeError, MigrationError, Exception) as exc:
        print(f"迁移失败：{exc}", file=sys.stderr)
        return 2
    print(json.dumps({"status": "PASS", "schema": result["schema"], "out": str(out)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
