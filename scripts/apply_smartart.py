#!/usr/bin/env python3
"""Replace marked native diagram fallbacks with PowerPoint SmartArt on Windows."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess


def _windows_path(path: Path) -> str:
    if os.name == "nt":
        return str(path)
    result = subprocess.run(
        ["wslpath", "-w", str(path)],
        check=True,
        text=True,
        capture_output=True,
    )
    return result.stdout.strip()


def apply_smartart(pptx_path: Path, spec_path: Path, *, slide_offset: int = 0) -> dict:
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if not powershell:
        return {
            "available": False,
            "converted": 0,
            "skipped": 0,
            "failures": ["PowerPoint SmartArt post-processing requires Windows PowerShell and desktop PowerPoint"],
        }
    script = Path(__file__).with_name("apply_smartart.ps1").resolve()
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        _windows_path(script),
        "-PptxPath",
        _windows_path(pptx_path.resolve()),
        "-SpecPath",
        _windows_path(spec_path.resolve()),
        "-SlideOffset",
        str(slide_offset),
    ]
    result = subprocess.run(command, check=False, text=True, capture_output=True)
    if result.returncode != 0:
        return {
            "available": True,
            "converted": 0,
            "skipped": 0,
            "failures": [(result.stderr or result.stdout or "PowerPoint SmartArt conversion failed").strip()],
        }
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError:
        payload = {
            "converted": 0,
            "skipped": 0,
            "failures": [f"Unexpected SmartArt converter output: {result.stdout.strip()}"],
        }
    return {"available": True, **payload}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", help="PPTX to update in place")
    parser.add_argument("spec", help="deck_spec.json used to build the PPTX")
    parser.add_argument("--slide-offset", type=int, default=0, help="Existing slides before generated content")
    args = parser.parse_args()
    report = apply_smartart(
        Path(args.pptx).expanduser().resolve(),
        Path(args.spec).expanduser().resolve(),
        slide_offset=args.slide_offset,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not report.get("failures") else 1


if __name__ == "__main__":
    raise SystemExit(main())
