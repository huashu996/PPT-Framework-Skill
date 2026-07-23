#!/usr/bin/env python3
"""Validate PPTX package integrity and reject invalid negative drawing extents."""

from __future__ import annotations

import argparse
import json
import posixpath
import sys
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def relationship_owner(rels_name: str) -> tuple[str, str]:
    path = PurePosixPath(rels_name)
    if rels_name == "_rels/.rels":
        return "", ""
    parts = list(path.parts)
    rels_index = parts.index("_rels")
    owner_name = parts[rels_index + 1]
    if not owner_name.endswith(".rels"):
        raise ValueError(f"Unexpected relationship part: {rels_name}")
    owner_name = owner_name[: -len(".rels")]
    owner_dir = "/".join(parts[:rels_index])
    return owner_dir, posixpath.join(owner_dir, owner_name)


def validate(path: Path) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    xml_count = 0
    extent_count = 0
    relationship_count = 0
    slide_count = 0

    try:
        archive = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        return {
            "ok": False,
            "path": str(path),
            "errors": [{"kind": "invalid_zip", "detail": str(exc)}],
            "warnings": [],
        }

    with archive:
        names = set(archive.namelist())
        bad_crc = archive.testzip()
        if bad_crc:
            errors.append({"kind": "crc_failure", "part": bad_crc})

        parsed: dict[str, ET.Element] = {}
        for name in sorted(names):
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                parsed[name] = ET.fromstring(archive.read(name))
                xml_count += 1
            except ET.ParseError as exc:
                errors.append(
                    {"kind": "malformed_xml", "part": name, "detail": str(exc)}
                )

        for name, root in parsed.items():
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                slide_count += 1
                ids: list[str] = []
                for node in root.findall(".//p:cNvPr", NS):
                    shape_id = node.get("id")
                    # PowerPoint may use repeated id="0" records inside embedded
                    # OLE preview markup; only positive drawing IDs identify slide shapes.
                    if shape_id not in (None, "0"):
                        ids.append(shape_id)
                duplicates = sorted({shape_id for shape_id in ids if ids.count(shape_id) > 1})
                if duplicates:
                    errors.append(
                        {
                            "kind": "duplicate_shape_id",
                            "part": name,
                            "ids": duplicates,
                        }
                    )

            for node in root.findall(".//a:ext", NS) + root.findall(".//a:chExt", NS):
                # a:ext is also used for extension-list entries with a URI and no
                # geometry attributes. Only extent nodes carrying cx/cy are geometry.
                if node.get("cx") is None and node.get("cy") is None:
                    continue
                extent_count += 1
                for attribute in ("cx", "cy"):
                    raw = node.get(attribute)
                    try:
                        value = int(raw) if raw is not None else None
                    except ValueError:
                        value = None
                    if value is None:
                        errors.append(
                            {
                                "kind": "invalid_extent",
                                "part": name,
                                "attribute": attribute,
                                "value": raw,
                            }
                        )
                    elif value < 0:
                        errors.append(
                            {
                                "kind": "negative_extent",
                                "part": name,
                                "attribute": attribute,
                                "value": value,
                            }
                        )

            if not name.endswith(".rels"):
                continue
            try:
                owner_dir, _ = relationship_owner(name)
            except ValueError as exc:
                warnings.append(
                    {"kind": "unrecognized_relationship_part", "part": name, "detail": str(exc)}
                )
                continue
            for rel in root.findall("r:Relationship", NS):
                relationship_count += 1
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target")
                if not target:
                    errors.append(
                        {"kind": "missing_relationship_target", "part": name}
                    )
                    continue
                if target.startswith("/"):
                    resolved = posixpath.normpath(target.lstrip("/"))
                else:
                    resolved = posixpath.normpath(posixpath.join(owner_dir, target))
                if resolved not in names:
                    errors.append(
                        {
                            "kind": "broken_relationship",
                            "part": name,
                            "target": target,
                            "resolved": resolved,
                        }
                    )

    return {
        "ok": not errors,
        "path": str(path),
        "slide_count": slide_count,
        "xml_part_count": xml_count,
        "extent_count": extent_count,
        "relationship_count": relationship_count,
        "errors": errors,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    args = parser.parse_args()
    result = validate(args.input.resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
