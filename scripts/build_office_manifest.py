#!/usr/bin/env python3
"""Build a minimal batch manifest for editable Office equations."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from latex_to_office_linear import convert_latex_to_office


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--target", default="powerpoint", choices=("powerpoint",))
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8-sig"))
    source_items = payload if isinstance(payload, list) else payload.get("equations", [])
    if not source_items:
        raise SystemExit("Input contains no formulas.")

    equations = []
    for index, item in enumerate(source_items, 1):
        formula_id = str(item.get("id") or f"formula_{index:03d}")
        latex = str(item.get("latex") or "").strip()
        if not latex:
            raise SystemExit(f"Missing latex for {formula_id}")
        linear, warnings, detected_number = convert_latex_to_office(latex)
        target = item.get("target")
        if not target:
            page = max(1, int(item.get("page", 1)))
            target = {"type": "powerpoint_shape", "slide_index": page, "shape_name": f"{formula_id}_target"}
        if str(target.get("type")) not in {"powerpoint_shape", "powerpoint_rect"}:
            raise SystemExit(f"PowerPoint target required for {formula_id}")
        equations.append({
            "id": formula_id,
            "native_linear": linear,
            "warnings": warnings,
            "equation_number": item.get("equation_number") or detected_number,
            "target": target,
        })

    args.output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = (args.output_dir / "manifest.json").resolve()
    manifest_path.write_text(json.dumps({"mode": "Office", "equations": equations}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"manifest": str(manifest_path), "formula_count": len(equations)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
