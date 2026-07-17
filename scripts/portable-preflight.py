#!/usr/bin/env python3
"""Non-blocking portability and formula-capability check for paper-fig-skill."""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


OPTIONAL_FILES = (
    "agents/openai.yaml",
    "references/quality-check.md",
    "references/mathtype-tool.md",
    "scripts/mathtype-ppt.ps1",
    "scripts/portable-preflight.py",
)


def probe_windows_ole() -> dict:
    result = {
        "powershell": None,
        "desktop_powerpoint": False,
        "mathtype_ribbon_addin": False,
        "mathtype_addins": [],
        "mathtype_prog_id": False,
        "mathtype_executable": None,
        "reason": None,
    }
    if os.name != "nt":
        result["reason"] = "PowerPoint COM and MathType OLE require Windows."
        return result

    powershell = shutil.which("pwsh") or shutil.which("powershell")
    result["powershell"] = powershell
    if not powershell:
        result["reason"] = "PowerShell was not found."
        return result

    probe = r"""
$ErrorActionPreference = 'Stop'
$r = [ordered]@{
  desktop_powerpoint = $false
  mathtype_ribbon_addin = $false
  mathtype_addins = @()
  mathtype_prog_id = $false
  mathtype_executable = $null
  reason = $null
}
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $r.desktop_powerpoint = $true
  foreach ($addin in $ppt.COMAddIns) {
    $description = [string]$addin.Description
    $progId = [string]$addin.ProgId
    if ("$description $progId" -match '(?i)MathType|Design Science|WIRIS') {
      $r.mathtype_ribbon_addin = $true
      $r.mathtype_addins += [ordered]@{
        kind = 'COMAddIn'
        name = $description
        prog_id = $progId
        connected = [bool]$addin.Connect
      }
    }
  }
  foreach ($addin in $ppt.AddIns) {
    $name = [string]$addin.Name
    $fullName = [string]$addin.FullName
    if ("$name $fullName" -match '(?i)MathType|Design Science|WIRIS') {
      $r.mathtype_ribbon_addin = $true
      $r.mathtype_addins += [ordered]@{
        kind = 'PowerPointAddIn'
        name = $name
        path = $fullName
        loaded = [bool]$addin.Loaded
      }
    }
  }
  $ppt.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
} catch {
  $r.reason = "PowerPoint COM unavailable: $($_.Exception.Message)"
}
try {
  $prog = 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4'
  $clsid = (Get-ItemProperty -LiteralPath "$prog\CLSID" -ErrorAction Stop).'(default)'
  if ($clsid) {
    $r.mathtype_prog_id = $true
    $server = (Get-ItemProperty -LiteralPath "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer32" -ErrorAction SilentlyContinue).'(default)'
    if (-not $server) {
      $server = (Get-ItemProperty -LiteralPath "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer" -ErrorAction SilentlyContinue).'(default)'
    }
    if ($server) {
      if ($server -match '^"([^"]+\.exe)"') { $r.mathtype_executable = $Matches[1] }
      elseif ($server -match '^(.+?\.exe)(?:\s|$)') { $r.mathtype_executable = $Matches[1].Trim() }
    }
  }
} catch {
  if (-not $r.reason) { $r.reason = 'MathType OLE ProgID Equation.DSMT4 is not registered.' }
}
$r | ConvertTo-Json -Compress
"""
    try:
        completed = subprocess.run(
            [
                powershell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                probe,
            ],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
        if completed.returncode:
            result["reason"] = (
                completed.stderr.strip()
                or f"PowerShell capability probe exited {completed.returncode}."
            )
        else:
            result.update(json.loads(completed.stdout.strip()))
    except (OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as exc:
        result["reason"] = f"PowerShell capability probe failed: {exc}"
    return result


def build_report(skill_root: Path) -> dict:
    skill_root = skill_root.expanduser().resolve()
    missing = [
        relative for relative in OPTIONAL_FILES if not (skill_root / relative).is_file()
    ]
    ole = probe_windows_ole()
    mathtype_available = bool(
        ole["desktop_powerpoint"]
        and (ole["mathtype_ribbon_addin"] or ole["mathtype_prog_id"])
    )
    strategies = []
    if ole["desktop_powerpoint"] and ole["mathtype_ribbon_addin"]:
        strategies.append("mathtype_powerpoint_ribbon")
    if ole["desktop_powerpoint"] and ole["mathtype_prog_id"]:
        strategies.append("mathtype_ole")
    if ole["desktop_powerpoint"]:
        strategies.append("office_native_equation")
    strategies.append("grouped_editable_text")

    warnings = []
    if missing:
        warnings.append(
            "Optional bundle files are missing. Continue with the self-contained "
            "SKILL.md workflow; do not infer that MathType is absent."
        )
    if not mathtype_available:
        warnings.append(
            "MathType OLE is not currently usable. Use an authorized fallback or "
            "repair the reported Windows registration/PowerPoint capability."
        )

    return {
        "skill_root": str(skill_root),
        "skill_md_present": (skill_root / "SKILL.md").is_file(),
        "bundle": {
            "complete": not missing,
            "optional_files_missing": missing,
        },
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "python": sys.version.split()[0],
        },
        "capabilities": {
            **ole,
            "mathtype_available": mathtype_available,
            "mathtype_ole": bool(
                ole["desktop_powerpoint"] and ole["mathtype_prog_id"]
            ),
            "office_native_equation": bool(ole["desktop_powerpoint"]),
            "grouped_editable_text": True,
        },
        "formula_strategy_order": strategies,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--skill-root",
        type=Path,
        default=Path(__file__).resolve().parent.parent,
        help="Directory containing SKILL.md.",
    )
    parser.add_argument(
        "--require",
        choices=("none", "bundle", "powerpoint", "mathtype"),
        default="none",
        help="Exit 2 only when this explicitly required capability is absent.",
    )
    args = parser.parse_args()
    report = build_report(args.skill_root)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    satisfied = {
        "none": True,
        "bundle": report["bundle"]["complete"],
        "powerpoint": report["capabilities"]["desktop_powerpoint"],
        "mathtype": report["capabilities"]["mathtype_available"],
    }[args.require]
    return 0 if satisfied else 2


if __name__ == "__main__":
    raise SystemExit(main())
