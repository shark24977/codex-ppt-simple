#!/usr/bin/env python3
"""Compare a generated PPTX page family with representative reference slides."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any

from extract_reference_style import extract_style, parse_slide_numbers


def _distance(left: str, right: str) -> float:
    left_rgb = [int(left[index : index + 2], 16) for index in (0, 2, 4)]
    right_rgb = [int(right[index : index + 2], 16) for index in (0, 2, 4)]
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(left_rgb, right_rgb)))


def _palette(profile: dict[str, Any]) -> list[str]:
    evidence = profile.get("evidence", {})
    candidates = evidence.get("local_palette_candidates") or []
    if not candidates:
        for key in ("common_fill_colors", "common_line_colors", "common_font_colors"):
            candidates.extend(evidence.get(key, []))
    colors = []
    for color in candidates:
        color = str(color).upper()
        if len(color) == 6 and color not in colors:
            colors.append(color)
    return colors[:8]


def compare_style(
    generated: Path,
    reference: Path,
    *,
    generated_slides: list[int] | None = None,
    reference_slides: list[int] | None = None,
) -> dict[str, Any]:
    generated_profile = extract_style(generated, generated_slides)
    reference_profile = extract_style(reference, reference_slides)
    generated_palette = _palette(generated_profile)
    reference_palette = _palette(reference_profile)
    color_matches = {
        color: min((_distance(color, candidate) for candidate in generated_palette), default=999.0)
        for color in reference_palette
    }
    palette_coverage = (
        sum(distance <= 28 for distance in color_matches.values()) / len(color_matches)
        if color_matches
        else 1.0
    )

    generated_fonts = generated_profile.get("fonts", {})
    reference_fonts = reference_profile.get("fonts", {})
    heading_font_match = generated_fonts.get("heading") == reference_fonts.get("heading")
    body_font_match = generated_fonts.get("body") == reference_fonts.get("body")
    generated_title = float(generated_profile.get("sizes", {}).get("title", 0) or 0)
    reference_title = float(reference_profile.get("sizes", {}).get("title", 0) or 0)
    title_size_delta = abs(generated_title - reference_title)
    title_size_score = max(0.0, 1.0 - title_size_delta / 10.0)
    background_match = (
        _distance(
            generated_profile["colors"]["background"],
            reference_profile["colors"]["background"],
        )
        <= 28
    )

    score = round(
        palette_coverage * 50
        + (15 if heading_font_match else 0)
        + (15 if body_font_match else 0)
        + title_size_score * 15
        + (5 if background_match else 0),
        1,
    )
    warnings: list[str] = []
    if palette_coverage < 0.6:
        warnings.append(
            f"palette coverage is {palette_coverage:.0%}; generated slides do not sufficiently follow the reference family"
        )
    if not heading_font_match:
        warnings.append(
            f"heading font differs: generated={generated_fonts.get('heading')!r}, "
            f"reference={reference_fonts.get('heading')!r}"
        )
    if not body_font_match:
        warnings.append(
            f"body font differs: generated={generated_fonts.get('body')!r}, reference={reference_fonts.get('body')!r}"
        )
    if title_size_delta > 6:
        warnings.append(
            f"median heading size differs by {title_size_delta:.1f} pt "
            f"(generated={generated_title:.1f}, reference={reference_title:.1f})"
        )
    if not background_match:
        warnings.append("background treatment differs from the selected reference family")

    return {
        "passed": score >= 70 and palette_coverage >= 0.5,
        "score": score,
        "generated": {
            "pptx": str(generated.resolve()),
            "slide_numbers": generated_profile["slide_numbers"],
            "palette": generated_palette,
            "fonts": generated_fonts,
            "title_size": generated_title,
        },
        "reference": {
            "pptx": str(reference.resolve()),
            "slide_numbers": reference_profile["slide_numbers"],
            "palette": reference_palette,
            "fonts": reference_fonts,
            "title_size": reference_title,
        },
        "metrics": {
            "palette_coverage": round(palette_coverage, 3),
            "heading_font_match": heading_font_match,
            "body_font_match": body_font_match,
            "title_size_delta": round(title_size_delta, 1),
            "background_match": background_match,
            "reference_color_distances": {key: round(value, 1) for key, value in color_matches.items()},
        },
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("generated_pptx", help="Generated PPTX")
    parser.add_argument("reference_pptx", help="Reference PPTX")
    parser.add_argument("--generated-slides", help="Generated page family, for example 2-4")
    parser.add_argument("--reference-slides", help="Representative reference family, for example 10-13,16")
    parser.add_argument("--json-out", help="Optional JSON report path")
    parser.add_argument("--strict", action="store_true", help="Exit nonzero when style fidelity is below threshold")
    args = parser.parse_args()

    generated = Path(args.generated_pptx).expanduser().resolve()
    reference = Path(args.reference_pptx).expanduser().resolve()
    for path in (generated, reference):
        if not path.is_file():
            print(f"PPTX not found: {path}")
            return 2
    try:
        report = compare_style(
            generated,
            reference,
            generated_slides=parse_slide_numbers(args.generated_slides),
            reference_slides=parse_slide_numbers(args.reference_slides),
        )
    except Exception as exc:
        print(f"Could not compare style: {exc}")
        return 1

    body = json.dumps(report, ensure_ascii=False, indent=2)
    print(body)
    if args.json_out:
        out = Path(args.json_out).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(body + "\n", encoding="utf-8")
    return 0 if report["passed"] or not args.strict else 1


if __name__ == "__main__":
    raise SystemExit(main())
