#!/usr/bin/env python3
"""Validate Stage 1 slide content and optionally render a human-readable proof."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


CONTENT_EXEMPT_ROLES = {"封面", "章节", "结束", "致谢", "cover", "section", "closing"}
PRIORITIES = {"critical", "high", "supporting"}


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot read JSON from {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object in {path}")
    return value


def nonempty_list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def normalized(value: Any) -> str:
    return re.sub(r"[\s\W_]+", "", str(value or "").lower(), flags=re.UNICODE)


def strings_from(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return [str(value)]
    if isinstance(value, list):
        result: list[str] = []
        for item in value:
            result.extend(strings_from(item))
        return result
    if isinstance(value, dict):
        result = []
        for item in value.values():
            result.extend(strings_from(item))
        return result
    return []


def slide_visible_strings(slide: dict[str, Any]) -> list[str]:
    values: list[str] = []
    for key in (
        "title",
        "takeaway",
        "subtitle",
        "kicker",
        "footer",
        "callouts",
        "metrics",
        "deliverables",
        "diagram_copy",
        "visible_content",
    ):
        values.extend(strings_from(slide.get(key)))
    for block in nonempty_list(slide.get("content_blocks")):
        if isinstance(block, dict):
            values.extend(strings_from(block.get("heading")))
            values.extend(strings_from(block.get("body")))
            values.extend(strings_from(block.get("bullets")))
            values.extend(strings_from(block.get("outcome")))
    return values


def is_content_slide(slide: dict[str, Any]) -> bool:
    if "content_slide" in slide:
        return bool(slide.get("content_slide"))
    return str(slide.get("page_role") or "").strip().lower() not in CONTENT_EXEMPT_ROLES


def validate(review: dict[str, Any], source_digest: dict[str, Any] | None = None) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []

    if review.get("version") != 1:
        errors.append("version must be 1")
    if review.get("status") not in {"review", "revised"}:
        errors.append("status must be 'review' or 'revised' during Stage 1")
    if review.get("approval_required") is not True:
        errors.append("approval_required must be true")
    for key in ("deck_title", "audience", "purpose"):
        if not str(review.get(key) or "").strip():
            errors.append(f"{key} is required")

    slides = review.get("slides")
    if not isinstance(slides, list) or not slides:
        errors.append("slides must be a non-empty array")
        slides = []

    slide_numbers: set[int] = set()
    slide_fact_ids: dict[int, set[str]] = {}
    required_term_count = 0
    matched_term_count = 0
    slide_reports: list[dict[str, Any]] = []

    for index, raw_slide in enumerate(slides, start=1):
        prefix = f"slide {index}"
        if not isinstance(raw_slide, dict):
            errors.append(f"{prefix} must be an object")
            continue
        slide = raw_slide
        number = slide.get("slide", index)
        if not isinstance(number, int) or isinstance(number, bool) or number < 1:
            errors.append(f"{prefix}.slide must be a positive integer")
            number = index
        elif number in slide_numbers:
            errors.append(f"{prefix}.slide duplicates page number {number}")
        slide_numbers.add(number)

        for key in ("page_role", "question", "answer", "title", "takeaway", "visual_relationship"):
            if not str(slide.get(key) or "").strip():
                errors.append(f"{prefix}.{key} is required")

        content = is_content_slide(slide)
        blocks = slide.get("content_blocks")
        if content and (not isinstance(blocks, list) or not blocks):
            errors.append(f"{prefix}.content_blocks must be a non-empty array")
            blocks = []
        elif not isinstance(blocks, list):
            blocks = []
        for block_index, block in enumerate(blocks, start=1):
            block_prefix = f"{prefix}.content_blocks[{block_index}]"
            if not isinstance(block, dict):
                errors.append(f"{block_prefix} must be an object")
                continue
            if not str(block.get("heading") or "").strip():
                errors.append(f"{block_prefix}.heading is required")
            body = strings_from(block.get("body")) + strings_from(block.get("bullets"))
            if not body and not str(block.get("outcome") or "").strip():
                errors.append(f"{block_prefix} needs body, bullets, or outcome text")
            elif content and len(normalized(" ".join(body))) < 8:
                warnings.append(f"{block_prefix} is very terse; verify that it states a mechanism or result")

        fact_ids = [str(item).strip() for item in nonempty_list(slide.get("fact_ids")) if str(item).strip()]
        must_terms = [
            str(item).strip() for item in nonempty_list(slide.get("must_show_terms")) if str(item).strip()
        ]
        source_refs = [
            str(item).strip() for item in nonempty_list(slide.get("source_refs")) if str(item).strip()
        ]
        answerability = [
            str(item).strip()
            for item in nonempty_list(slide.get("audience_can_explain"))
            if str(item).strip()
        ]
        if content and not fact_ids:
            errors.append(f"{prefix}.fact_ids must identify supporting source facts")
        if content and not must_terms:
            errors.append(f"{prefix}.must_show_terms must protect source-specific visible terminology")
        if content and not source_refs:
            errors.append(f"{prefix}.source_refs must identify source locations")
        if content and len(answerability) < 2:
            errors.append(f"{prefix}.audience_can_explain must contain at least 2 review questions")
        slide_fact_ids[number] = set(fact_ids)

        visible = normalized(" ".join(slide_visible_strings(slide)))
        missing_terms = [term for term in must_terms if normalized(term) not in visible]
        required_term_count += len(must_terms)
        matched_term_count += len(must_terms) - len(missing_terms)
        if missing_terms:
            errors.append(f"{prefix} is missing must_show_terms in visible copy: {', '.join(missing_terms)}")
        if content and len(visible) < 40:
            warnings.append(f"{prefix} has little visible copy; verify that this is more than a topic outline")

        slide_reports.append(
            {
                "slide": number,
                "title": str(slide.get("title") or ""),
                "fact_ids": fact_ids,
                "must_show_terms": len(must_terms),
                "missing_terms": missing_terms,
                "visible_characters": len(visible),
            }
        )

    coverage = review.get("evidence_coverage")
    if not isinstance(coverage, list) or not coverage:
        errors.append("evidence_coverage must be a non-empty array")
        coverage = []

    coverage_ids: set[str] = set()
    priority_totals = {priority: 0 for priority in PRIORITIES}
    priority_visible = {priority: 0 for priority in PRIORITIES}
    valid_slide_numbers = set(slide_fact_ids)
    for index, raw_item in enumerate(coverage, start=1):
        prefix = f"evidence_coverage[{index}]"
        if not isinstance(raw_item, dict):
            errors.append(f"{prefix} must be an object")
            continue
        fact_id = str(raw_item.get("fact_id") or "").strip()
        priority = str(raw_item.get("priority") or "").strip()
        visible = raw_item.get("visible")
        pages = raw_item.get("slides")
        if not fact_id:
            errors.append(f"{prefix}.fact_id is required")
        elif fact_id in coverage_ids:
            errors.append(f"{prefix}.fact_id is duplicated: {fact_id}")
        coverage_ids.add(fact_id)
        if priority not in PRIORITIES:
            errors.append(f"{prefix}.priority must be critical, high, or supporting")
            continue
        priority_totals[priority] += 1
        if not isinstance(visible, bool):
            errors.append(f"{prefix}.visible must be true or false")
            visible = False
        if visible:
            priority_visible[priority] += 1
        if not isinstance(pages, list) or not all(isinstance(page, int) and page > 0 for page in pages):
            errors.append(f"{prefix}.slides must be an array of positive page numbers")
            pages = []
        unknown_pages = sorted(set(pages) - valid_slide_numbers)
        if unknown_pages:
            errors.append(f"{prefix}.slides references unknown pages: {unknown_pages}")
        if visible and not pages:
            errors.append(f"{prefix} is visible but has no slide mapping")
        if visible and not str(raw_item.get("rendered_as") or "").strip():
            errors.append(f"{prefix}.rendered_as is required for visible facts")
        if not str(raw_item.get("source_location") or "").strip():
            errors.append(f"{prefix}.source_location is required")
        if priority == "critical" and visible is not True:
            errors.append(f"{prefix} is critical and must remain visible")
        for page in pages:
            if fact_id and fact_id not in slide_fact_ids.get(page, set()):
                errors.append(f"{prefix} maps {fact_id} to slide {page}, but that slide omits it from fact_ids")

    high_total = priority_totals["high"]
    high_ratio = priority_visible["high"] / high_total if high_total else 1.0
    critical_total = priority_totals["critical"]
    critical_ratio = priority_visible["critical"] / critical_total if critical_total else 1.0
    if high_ratio < 0.8:
        errors.append(f"high-priority visible coverage is {high_ratio:.0%}; at least 80% is required")

    all_slide_facts = set().union(*slide_fact_ids.values()) if slide_fact_ids else set()
    missing_coverage = sorted(all_slide_facts - coverage_ids)
    if missing_coverage:
        errors.append(f"slide fact_ids missing from evidence_coverage: {', '.join(missing_coverage)}")

    if source_digest is not None:
        raw_facts = source_digest.get("facts")
        digest_ids = {
            str(item.get("id") or "").strip()
            for item in raw_facts
            if isinstance(item, dict) and str(item.get("id") or "").strip()
        } if isinstance(raw_facts, list) else set()
        if not digest_ids:
            warnings.append("source_digest has no fact IDs to cross-check")
        else:
            unknown = sorted(coverage_ids - digest_ids)
            omitted = sorted(digest_ids - coverage_ids)
            if unknown:
                errors.append(f"evidence_coverage references unknown source facts: {', '.join(unknown)}")
            if omitted:
                errors.append(f"source facts missing from evidence_coverage: {', '.join(omitted)}")

    term_ratio = matched_term_count / required_term_count if required_term_count else 1.0
    score = max(0, 100 - min(100, len(errors) * 12 + len(warnings) * 2))
    return {
        "valid": not errors,
        "passed": not errors,
        "score": score,
        "slides": len(slides),
        "coverage": {
            "critical_visible": critical_ratio,
            "high_visible": high_ratio,
            "required_terms_visible": term_ratio,
            "facts_total": len(coverage_ids),
        },
        "errors": errors,
        "warnings": warnings,
        "slide_reports": slide_reports,
    }


def markdown(review: dict[str, Any], report: dict[str, Any]) -> str:
    lines = [
        f"# {review.get('deck_title', 'PPT内容审稿')}",
        "",
        f"- 受众：{review.get('audience', '')}",
        f"- 目的：{review.get('purpose', '')}",
        f"- 建议页数：{report.get('slides', 0)}",
        f"- 内容QA：{report.get('score', 0)}/100",
        f"- 关键事实可见率：{report.get('coverage', {}).get('critical_visible', 0):.0%}",
        f"- 高优先级事实可见率：{report.get('coverage', {}).get('high_visible', 0):.0%}",
        "",
    ]
    for slide in nonempty_list(review.get("slides")):
        if not isinstance(slide, dict):
            continue
        number = slide.get("slide", "?")
        lines.extend(
            [
                f"## 第{number}页｜{slide.get('page_role', '')}",
                "",
                f"**回答问题：** {slide.get('question', '')}",
                "",
                f"**核心回答：** {slide.get('answer', '')}",
                "",
                f"**页面标题：** {slide.get('title', '')}",
                "",
                f"**结论句：** {slide.get('takeaway', '')}",
                "",
                "### 页面可见内容",
                "",
            ]
        )
        for block in nonempty_list(slide.get("content_blocks")):
            if not isinstance(block, dict):
                continue
            lines.append(f"- **{block.get('heading', '')}**")
            body = strings_from(block.get("body")) + strings_from(block.get("bullets"))
            for item in body:
                lines.append(f"  - {item}")
            if str(block.get("outcome") or "").strip():
                lines.append(f"  - 结果：{block.get('outcome')}")
        for key, label in (("callouts", "强调信息"), ("metrics", "指标"), ("deliverables", "交付物")):
            values = strings_from(slide.get(key))
            if values:
                lines.append(f"- **{label}**：{'；'.join(values)}")
        lines.extend(
            [
                "",
                f"**视觉关系：** {slide.get('visual_relationship', '')}",
                "",
                f"**必须保留术语：** {'、'.join(strings_from(slide.get('must_show_terms')))}",
                "",
                f"**来源事实：** {'、'.join(strings_from(slide.get('fact_ids')))}",
                "",
                f"**来源位置：** {'；'.join(strings_from(slide.get('source_refs')))}",
                "",
            ]
        )
    if report.get("errors") or report.get("warnings"):
        lines.extend(["## 内容QA问题", ""])
        for item in report.get("errors", []):
            lines.append(f"- 错误：{item}")
        for item in report.get("warnings", []):
            lines.append(f"- 提醒：{item}")
        lines.append("")
    lines.extend(["---", "", "当前仅为内容审稿，尚未生成PPTX。", ""])
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("review", help="Path to content_review.json")
    parser.add_argument("--source-digest", help="Optional source_digest.json for fact-ID cross-checking")
    parser.add_argument("--json-out", help="Optional JSON QA output")
    parser.add_argument("--markdown-out", help="Optional human-readable Markdown proof")
    parser.add_argument("--strict", action="store_true", help="Exit nonzero when content review fails")
    args = parser.parse_args()

    review_path = Path(args.review).expanduser().resolve()
    try:
        review = load_json(review_path)
        source_digest = load_json(Path(args.source_digest).expanduser().resolve()) if args.source_digest else None
    except ValueError as exc:
        print(str(exc))
        return 2

    report = validate(review, source_digest)
    body = json.dumps(report, ensure_ascii=False, indent=2)
    print(body)
    if args.json_out:
        output = Path(args.json_out).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(body + "\n", encoding="utf-8")
    if args.markdown_out:
        output = Path(args.markdown_out).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(markdown(review, report), encoding="utf-8")
    return 0 if report["passed"] or not args.strict else 1


if __name__ == "__main__":
    raise SystemExit(main())
