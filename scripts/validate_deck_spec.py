#!/usr/bin/env python3
"""Validate a codex-ppt-simple deck specification."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx_common import load_json, validate_spec


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", help="Path to deck_spec.json")
    parser.add_argument("--check-assets", action="store_true", help="Require local image assets to exist")
    parser.add_argument("--json", action="store_true", help="Print machine-readable output")
    args = parser.parse_args()

    spec_path = Path(args.spec).expanduser().resolve()
    try:
        spec = load_json(spec_path)
    except ValueError as exc:
        print(str(exc))
        return 2
    errors, warnings = validate_spec(spec, spec_path=spec_path, check_assets=args.check_assets)
    result = {
        "valid": not errors,
        "slides": len(spec.get("slides", [])) if isinstance(spec.get("slides"), list) else 0,
        "errors": errors,
        "warnings": warnings,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"valid={str(result['valid']).lower()} slides={result['slides']}")
        for message in errors:
            print(f"ERROR: {message}")
        for message in warnings:
            print(f"WARNING: {message}")
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
