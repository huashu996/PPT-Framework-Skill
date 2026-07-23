#!/usr/bin/env python3
"""Convert LaTeX to well-formed, MathType-compatible Presentation MathML."""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path

MATH_NS = "http://www.w3.org/1998/Math/MathML"
ET.register_namespace("", MATH_NS)


def normalize_latex(latex: str) -> str:
    value = latex.strip()
    if re.search(r"\\begin\{(?:align\*?|aligned)\}", value):
        value = re.sub(r"\\begin\{(?:align\*?|aligned)\}", r"\\begin{gathered}", value)
        value = re.sub(r"\\end\{(?:align\*?|aligned)\}", r"\\end{gathered}", value)
        value = value.replace("&", "")
    return value


def variant_from_name(name: str) -> str | None:
    if "SCRIPT" in name:
        return "bold-script" if "BOLD" in name else "script"
    if "DOUBLE-STRUCK" in name:
        return "double-struck"
    if "FRAKTUR" in name:
        return "bold-fraktur" if "BOLD" in name else "fraktur"
    if "MONOSPACE" in name:
        return "monospace"
    if "SANS-SERIF" in name:
        return "sans-serif"
    if "BOLD ITALIC" in name:
        return "bold-italic"
    if "BOLD" in name:
        return "bold"
    if "ITALIC" in name:
        return "italic"
    return None


def normalize_math_alphanumerics(mathml: str) -> str:
    pattern = re.compile(r"<mi>&#x([0-9A-Fa-f]+);</mi>")

    def replace(match: re.Match[str]) -> str:
        character = chr(int(match.group(1), 16))
        name = unicodedata.name(character, "")
        base = unicodedata.normalize("NFKC", character)
        variant = variant_from_name(name)
        if name.startswith("MATHEMATICAL ") and variant and len(base) == 1 and base.isalnum():
            return f'<mi mathvariant="{variant}">{base}</mi>'
        return match.group(0)

    return pattern.sub(replace, mathml)


def normalize_linebreaks(root: ET.Element) -> None:
    linebreak_tag = f"{{{MATH_NS}}}mspace"
    for parent in list(root.iter()):
        children = list(parent)
        if not any(child.tag == linebreak_tag and child.get("linebreak") == "newline" for child in children):
            continue
        rows: list[list[ET.Element]] = [[]]
        for child in children:
            if child.tag == linebreak_tag and child.get("linebreak") == "newline":
                rows.append([])
            else:
                rows[-1].append(child)
        for child in children:
            parent.remove(child)
        table = ET.SubElement(parent, f"{{{MATH_NS}}}mtable")
        for items in rows:
            row = ET.SubElement(table, f"{{{MATH_NS}}}mtr")
            cell = ET.SubElement(row, f"{{{MATH_NS}}}mtd")
            group = ET.SubElement(cell, f"{{{MATH_NS}}}mrow")
            for item in items:
                group.append(item)


def remove_empty_formula_tokens(root: ET.Element) -> None:
    """Remove empty Presentation MathML token nodes before MathType import."""
    token_names = {"mi", "mo", "mn", "mtext"}
    for parent in list(root.iter()):
        for child in list(parent):
            local_name = child.tag.rsplit("}", 1)[-1]
            if local_name not in token_names:
                continue
            if list(child):
                continue
            if "".join(child.itertext()).strip():
                continue
            parent.remove(child)


def convert_mathtype_mathml(latex: str) -> str:
    from latex2mathml.converter import convert

    raw = normalize_math_alphanumerics(convert(normalize_latex(latex)))
    root = ET.fromstring(raw)
    normalize_linebreaks(root)
    remove_empty_formula_tokens(root)
    result = ET.tostring(root, encoding="unicode")
    ET.fromstring(result)
    if 'linebreak="newline"' in result or "<mi>&</mi>" in result:
        raise ValueError("MathML contains a MathType-incompatible multiline marker.")
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--latex", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    mathml = convert_mathtype_mathml(args.latex)
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(mathml, encoding="utf-8")
    print(json.dumps({"latex": args.latex, "mathml_file": str(output), "needs_review": False, "characters": len(mathml)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
