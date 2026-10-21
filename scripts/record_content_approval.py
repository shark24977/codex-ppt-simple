#!/usr/bin/env python3
"""Record explicit user approval for a validated Stage 1 content review."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path

from validate_content_review import load_json, validate


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("review", help="Approved content_review.json")
    parser.add_argument("--instruction", required=True, help="User message that explicitly approved generation")
    parser.add_argument("--out", required=True, help="Output content_approval.json")
    args = parser.parse_args()

    review_path = Path(args.review).expanduser().resolve()
    output_path = Path(args.out).expanduser().resolve()
    instruction = str(args.instruction).strip()
    if not instruction:
        print("Approval instruction must not be empty")
        return 2
    try:
        review = load_json(review_path)
    except ValueError as exc:
        print(str(exc))
        return 2
    report = validate(review)
    if not report["passed"]:
        print("Content review is not valid; approval was not recorded")
        for error in report["errors"]:
            print(f"ERROR: {error}")
        return 2

    output_path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "version": 1,
        "status": "approved",
        "approved_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "content_review": os.path.relpath(review_path, output_path.parent),
        "content_sha256": sha256_file(review_path),
        "user_instruction": instruction,
    }
    output_path.write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(record, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
