#!/usr/bin/env python3
"""Inspect a deck specification for summary quality and semantic-diagram risks."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx_common import content_quality_report, load_json


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", help="Path to deck_spec.json")
    parser.add_argument("--json-out", help="Optional JSON report path")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit nonzero when the quality score is below 80 or the spec is invalid",
    )
    args = parser.parse_args()

    spec_path = Path(args.spec).expanduser().resolve()
    try:
        spec = load_json(spec_path)
        report = content_quality_report(spec)
    except ValueError as exc:
        print(str(exc))
        return 2

    body = json.dumps(report, ensure_ascii=False, indent=2)
    print(body)
    if args.json_out:
        out = Path(args.json_out).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(body + "\n", encoding="utf-8")
    return 0 if report["passed"] or not args.strict else 1


if __name__ == "__main__":
    raise SystemExit(main())
