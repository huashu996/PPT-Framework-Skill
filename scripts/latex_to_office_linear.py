#!/usr/bin/env python3
"""Convert common LaTeX into Office UnicodeMath linear notation."""

from __future__ import annotations

import re

GREEK = {
    "alpha": "\u03b1", "beta": "\u03b2", "gamma": "\u03b3", "delta": "\u03b4",
    "epsilon": "\u03f5", "varepsilon": "\u03b5", "zeta": "\u03b6", "eta": "\u03b7",
    "theta": "\u03b8", "vartheta": "\u03d1", "iota": "\u03b9", "kappa": "\u03ba",
    "lambda": "\u03bb", "mu": "\u03bc", "nu": "\u03bd", "xi": "\u03be", "pi": "\u03c0",
    "rho": "\u03c1", "sigma": "\u03c3", "tau": "\u03c4", "upsilon": "\u03c5",
    "phi": "\u03c6", "varphi": "\u03d5", "chi": "\u03c7", "psi": "\u03c8", "omega": "\u03c9",
    "Gamma": "\u0393", "Delta": "\u0394", "Theta": "\u0398", "Lambda": "\u039b",
    "Xi": "\u039e", "Pi": "\u03a0", "Sigma": "\u03a3", "Phi": "\u03a6",
    "Psi": "\u03a8", "Omega": "\u03a9",
}

SYMBOLS = {
    "sum": "\u2211", "prod": "\u220f", "int": "\u222b", "iint": "\u222c", "iiint": "\u222d",
    "oint": "\u222e", "infty": "\u221e", "partial": "\u2202", "nabla": "\u2207",
    "cdot": "\u22c5", "times": "\u00d7", "pm": "\u00b1", "mp": "\u2213",
    "le": "\u2264", "leq": "\u2264", "ge": "\u2265", "geq": "\u2265",
    "ne": "\u2260", "neq": "\u2260", "approx": "\u2248", "sim": "\u223c",
    "to": "\u2192", "rightarrow": "\u2192", "leftarrow": "\u2190", "leftrightarrow": "\u2194",
    "in": "\u2208", "notin": "\u2209", "subset": "\u2282", "subseteq": "\u2286",
    "cup": "\u222a", "cap": "\u2229", "forall": "\u2200", "exists": "\u2203",
    "ell": "\u2113", "hbar": "\u210f", "langle": "\u27e8", "rangle": "\u27e9",
    "vert": "|", "mid": "|", "lVert": "\u2016", "rVert": "\u2016",
}

ACCENTS = {
    "hat": "\u0302", "widehat": "\u0302", "tilde": "\u0303", "widetilde": "\u0303",
    "bar": "\u0305", "overline": "\u0305", "dot": "\u0307", "ddot": "\u0308", "vec": "\u20d7",
}

FUNCTIONS = {"sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "lim", "min", "max", "det", "arg"}


def styled_math(text: str, style: str) -> str:
    if style == "mathbb":
        upper, lower, digits = 0x1D538, 0x1D552, 0x1D7D8
        exceptions = {"C": "\u2102", "H": "\u210d", "N": "\u2115", "P": "\u2119", "Q": "\u211a", "R": "\u211d", "Z": "\u2124"}
    elif style == "mathcal":
        upper, lower, digits = 0x1D49C, 0x1D4B6, None
        exceptions = {"B": "\u212c", "E": "\u2130", "F": "\u2131", "H": "\u210b", "I": "\u2110", "L": "\u2112", "M": "\u2133", "R": "\u211b", "e": "\u212f", "g": "\u210a", "o": "\u2134"}
    else:
        upper, lower, digits = 0x1D400, 0x1D41A, 0x1D7CE
        exceptions = {}
    result = []
    for character in text:
        if character in exceptions:
            result.append(exceptions[character])
        elif "A" <= character <= "Z":
            result.append(chr(upper + ord(character) - ord("A")))
        elif "a" <= character <= "z":
            result.append(chr(lower + ord(character) - ord("a")))
        elif digits is not None and "0" <= character <= "9":
            result.append(chr(digits + ord(character) - ord("0")))
        else:
            result.append(character)
    return "".join(result)


def group(text: str, index: int, opener: str = "{", closer: str = "}") -> tuple[str, int]:
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != opener:
        raise ValueError(f"Expected {opener} at offset {index}")
    depth = 1
    cursor = index + 1
    while cursor < len(text):
        if text[cursor] == opener:
            depth += 1
        elif text[cursor] == closer:
            depth -= 1
            if depth == 0:
                return text[index + 1:cursor], cursor + 1
        cursor += 1
    raise ValueError(f"Unclosed {opener} group")


def replace_environments(text: str, warnings: list[str]) -> str:
    pattern = re.compile(r"\\begin\{(matrix|pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix|cases|aligned|align\*?|gathered)\}(.*?)\\end\{\1\}", re.S)
    while True:
        match = pattern.search(text)
        if not match:
            return text
        kind, body = match.group(1), match.group(2)
        rows = []
        for row in re.split(r"\\\\", body):
            cells = [convert_latex_to_office(cell.strip())[0] for cell in row.split("&")]
            rows.append("&".join(cells))
        marker = "\u2588" if kind.startswith("align") or kind == "gathered" else "\u25a0"
        matrix = marker + "(" + "@".join(rows) + ")"
        wrappers = {"pmatrix": ("(", ")"), "bmatrix": ("[", "]"), "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("\u2016", "\u2016")}
        if kind == "cases":
            matrix = "{" + matrix
        elif kind in wrappers:
            left, right = wrappers[kind]
            matrix = left + matrix + right
        text = text[:match.start()] + matrix + text[match.end():]


def convert_latex_to_office(latex: str) -> tuple[str, list[str], str | None]:
    warnings: list[str] = []
    text = latex.strip()
    text = re.sub(r"^\s*\$+|\$+\s*$", "", text)
    text = text.replace("\\[", "").replace("\\]", "").replace("\\(", "").replace("\\)", "")
    equation_number = None
    tag = re.search(r"\\tag\{([^{}]+)\}", text)
    if tag:
        equation_number = tag.group(1)
        text = text[:tag.start()] + text[tag.end():]
    text = replace_environments(text, warnings)

    output: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\":
            match = re.match(r"\\([A-Za-z]+|.)", text[index:])
            if not match:
                output.append(char)
                index += 1
                continue
            command = match.group(1)
            index += len(match.group(0))
            while index < len(text) and text[index].isspace() and command.isalpha():
                index += 1
            if command in {"left", "right", "!", ",", ";", ":"}:
                continue
            if command in {"quad", "qquad"}:
                output.append(" ")
                continue
            if command in {"frac", "dfrac", "tfrac"}:
                numerator, index = group(text, index)
                denominator, index = group(text, index)
                n, nw, _ = convert_latex_to_office(numerator)
                d, dw, _ = convert_latex_to_office(denominator)
                warnings.extend(nw + dw)
                output.append(f"({n})/({d})")
                continue
            if command == "sqrt":
                root_index = None
                if index < len(text) and text[index] == "[":
                    root_index, index = group(text, index, "[", "]")
                body, index = group(text, index)
                converted, child_warnings, _ = convert_latex_to_office(body)
                warnings.extend(child_warnings)
                output.append("\u221a(" + ((root_index + "&") if root_index else "") + converted + ")")
                continue
            if command in ACCENTS or command in {"mathbf", "boldsymbol", "bm", "mathbb", "mathcal", "mathit", "mathrm", "operatorname", "text"}:
                body, index = group(text, index)
                converted, child_warnings, _ = convert_latex_to_office(body)
                warnings.extend(child_warnings)
                if command in ACCENTS:
                    converted = "".join(character + ACCENTS[command] for character in converted)
                elif command in {"mathbf", "boldsymbol", "bm"}:
                    converted = styled_math(converted, "mathbf")
                elif command in {"mathbb", "mathcal"}:
                    converted = styled_math(converted, command)
                output.append(converted)
                continue
            if command in GREEK:
                output.append(GREEK[command])
                continue
            if command in SYMBOLS:
                output.append(SYMBOLS[command])
                continue
            if command in FUNCTIONS:
                output.append(command)
                continue
            warnings.append(f"unsupported LaTeX command: \\{command}")
            output.append(command)
            continue
        if char in "_^":
            output.append(char)
            index += 1
            while index < len(text) and text[index].isspace():
                index += 1
            if index < len(text) and text[index] == "{":
                body, index = group(text, index)
                converted, child_warnings, _ = convert_latex_to_office(body)
                warnings.extend(child_warnings)
                output.append(f"({converted})")
            elif index < len(text):
                output.append(text[index])
                index += 1
            continue
        if char == "{":
            body, index = group(text, index)
            converted, child_warnings, _ = convert_latex_to_office(body)
            warnings.extend(child_warnings)
            output.append(f"({converted})")
            continue
        if char == "}":
            index += 1
            continue
        output.append(char)
        index += 1

    linear = re.sub(r"\s+", " ", "".join(output)).strip()
    return linear, sorted(set(warnings)), equation_number
