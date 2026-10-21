#!/usr/bin/env python3
"""Shared specification and style helpers for codex-ppt-simple."""

from __future__ import annotations

import copy
from collections import Counter
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple


SUPPORTED_SLIDE_TYPES = {
    "cover",
    "section",
    "title_body",
    "two_column",
    "comparison",
    "cards",
    "decomposition",
    "process",
    "diagram",
    "composite_diagram",
    "timeline",
    "chart",
    "table",
    "image_text",
    "closing",
}

SUPPORTED_DIAGRAM_KINDS = {
    "flow",
    "decision",
    "hierarchy",
    "cycle",
    "swimlane",
    "architecture",
}

SUPPORTED_DIAGRAM_DIRECTIONS = {"left_to_right", "top_to_bottom"}
SUPPORTED_DIAGRAM_RENDERS = {"auto", "native", "smartart"}
SUPPORTED_DIAGRAM_ROLES = {
    "default",
    "start",
    "end",
    "process",
    "decision",
    "input",
    "output",
    "system",
    "group",
    "constraint",
}

SUPPORTED_EDGE_TYPES = {
    "flow",
    "data",
    "control",
    "constraint",
    "feedback",
    "output",
    "reference",
}

SUPPORTED_SEMANTIC_INTENTS = {
    "decomposition",
    "sequence",
    "decision",
    "hierarchy",
    "cycle",
    "architecture",
    "responsibility",
    "composite",
}

SUMMARY_LEVEL_LIMITS = {
    "executive": {"max_bullets": 3, "max_bullet_chars": 34},
    "standard": {"max_bullets": 4, "max_bullet_chars": 52},
    "technical": {"max_bullets": 5, "max_bullet_chars": 76},
}

DEFAULT_READABILITY_POLICY: Dict[str, Any] = {
    "min_title_font": 26.0,
    "min_takeaway_font": 17.0,
    "min_module_font": 15.0,
    "min_body_font": 13.5,
    "min_diagram_font": 12.0,
    "min_label_font": 11.0,
    "overflow_action": "revise_or_split",
}

DEFAULT_EMPHASIS_POLICY: Dict[str, Any] = {
    "enabled": True,
    "min_bold_ratio": 0.10,
    "max_bold_ratio": 0.30,
    "max_highlights_per_block": 3,
    "auto_emphasize_leading_labels": True,
}

DEFAULT_DENSITY_POLICY: Dict[str, Any] = {
    "max_visible_characters": 420,
    "max_primary_blocks": 5,
    "max_visual_focal_points": 2,
}

DEFAULT_STYLE: Dict[str, Any] = {
    "colors": {
        "background": "F7F9FC",
        "surface": "FFFFFF",
        "text": "17212B",
        "muted_text": "667085",
        "primary": "173F5F",
        "accent": "00A6A6",
        "accent_2": "F6AE2D",
        "line": "D9E2EC",
    },
    "fonts": {
        "heading": "Aptos Display",
        "body": "Aptos",
    },
    "sizes": {
        "cover_title": 34,
        "title": 28,
        "subtitle": 18,
        "body": 15,
        "small": 11,
        "card_value": 24,
    },
    "geometry": {
        "margin_x": 0.72,
        "margin_top": 0.46,
        "margin_bottom": 0.42,
        "gutter": 0.28,
        "title_height": 0.55,
    },
}

_HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def load_json(path: Path) -> Dict[str, Any]:
    path = Path(path)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ValueError(f"JSON file not found: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object in {path}")
    return value


def deep_merge(base: Dict[str, Any], override: Dict[str, Any] | None) -> Dict[str, Any]:
    result = copy.deepcopy(base)
    if not override:
        return result
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def normalize_hex(value: Any, fallback: str | None = None) -> str:
    text = str(value or "").strip().lstrip("#")
    if _HEX_RE.match(text):
        return text.upper()
    if fallback is not None:
        return normalize_hex(fallback)
    raise ValueError(f"Expected a six-digit RGB color, got: {value!r}")


def merged_style(profile: Dict[str, Any] | None, overrides: Dict[str, Any] | None) -> Dict[str, Any]:
    style = deep_merge(DEFAULT_STYLE, profile or {})
    style = deep_merge(style, overrides or {})
    colors = style.setdefault("colors", {})
    for key, default in DEFAULT_STYLE["colors"].items():
        colors[key] = normalize_hex(colors.get(key), default)
    return style


def merged_policy(spec: Dict[str, Any], key: str, defaults: Dict[str, Any]) -> Dict[str, Any]:
    raw = spec.get(key)
    return deep_merge(defaults, raw if isinstance(raw, dict) else {})


def _as_list(value: Any) -> List[Any]:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def _image_entries(slide: Dict[str, Any]) -> Iterable[Dict[str, Any]]:
    image = slide.get("image")
    if isinstance(image, str):
        yield {"path": image}
    elif isinstance(image, dict):
        yield image
    for image in _as_list(slide.get("images")):
        if isinstance(image, str):
            yield {"path": image}
        elif isinstance(image, dict):
            yield image


def resolve_asset_path(spec_path: Path, value: str) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = spec_path.resolve().parent / path
    return path.resolve()


def _normalized_text(value: Any) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", str(value or "").lower())


def _text_similarity(left: Any, right: Any) -> float:
    """Return a small, dependency-free character-bigram similarity score."""

    def bigrams(value: Any) -> set[str]:
        text = _normalized_text(value)
        if len(text) < 2:
            return {text} if text else set()
        return {text[index : index + 2] for index in range(len(text) - 1)}

    left_parts = bigrams(left)
    right_parts = bigrams(right)
    if not left_parts or not right_parts:
        return 0.0
    return len(left_parts & right_parts) / len(left_parts | right_parts)


def _content_profile(spec: Dict[str, Any]) -> Dict[str, Any]:
    raw = spec.get("content_profile")
    return raw if isinstance(raw, dict) else {}


def _slide_bullets(slide: Dict[str, Any]) -> List[str]:
    bullets: List[str] = []
    for value in _as_list(slide.get("bullets")):
        if str(value or "").strip():
            bullets.append(str(value).strip())
    for side in ("left", "right"):
        section = slide.get(side)
        if isinstance(section, dict):
            for value in _as_list(section.get("bullets")):
                if str(value or "").strip():
                    bullets.append(str(value).strip())
    for key in ("cards", "steps", "items", "events"):
        for item in _as_list(slide.get(key)):
            if not isinstance(item, dict):
                continue
            for value in _as_list(item.get("bullets")):
                if str(value or "").strip():
                    bullets.append(str(value).strip())
    return bullets


def _visible_string_values(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return [str(value)]
    if isinstance(value, list):
        result: List[str] = []
        for item in value:
            result.extend(_visible_string_values(item))
        return result
    return []


def _diagram_visible_strings(diagram: Any) -> List[str]:
    if not isinstance(diagram, dict):
        return []
    values: List[str] = []
    for node in _as_list(diagram.get("nodes")):
        if isinstance(node, dict):
            values.extend(_visible_string_values(node.get("label")))
            values.extend(_visible_string_values(node.get("body")))
    for edge in _as_list(diagram.get("edges")):
        if isinstance(edge, dict):
            values.extend(_visible_string_values(edge.get("label")))
    for lane in _as_list(diagram.get("lanes")):
        if isinstance(lane, dict):
            values.extend(_visible_string_values(lane.get("label")))
        else:
            values.extend(_visible_string_values(lane))
    return values


def _slide_visible_strings(slide: Dict[str, Any]) -> List[str]:
    """Collect text rendered on a slide, excluding notes, evidence, and review metadata."""

    values: List[str] = []
    for key in ("title", "takeaway", "subtitle", "kicker", "footer", "bullets"):
        values.extend(_visible_string_values(slide.get(key)))
    for side in ("left", "right"):
        section = slide.get(side)
        if isinstance(section, dict):
            for key in ("heading", "body", "bullets", "outcome", "value"):
                values.extend(_visible_string_values(section.get(key)))
    for key in ("cards", "steps", "items", "events"):
        for item in _as_list(slide.get(key)):
            if not isinstance(item, dict):
                values.extend(_visible_string_values(item))
                continue
            for field in ("heading", "label", "body", "bullets", "outcome", "value"):
                values.extend(_visible_string_values(item.get(field)))
    values.extend(_diagram_visible_strings(slide.get("diagram")))
    for panel in _as_list(slide.get("panels")):
        if not isinstance(panel, dict):
            continue
        values.extend(_visible_string_values(panel.get("heading")))
        values.extend(_visible_string_values(panel.get("takeaway")))
        values.extend(_diagram_visible_strings(panel.get("diagram")))
    values.extend(_visible_string_values(slide.get("columns")))
    values.extend(_visible_string_values(slide.get("rows")))
    chart = slide.get("chart")
    if isinstance(chart, dict):
        values.extend(_visible_string_values(chart.get("categories")))
        for series in _as_list(chart.get("series")):
            if isinstance(series, dict):
                values.extend(_visible_string_values(series.get("name")))
    return values


def _visible_character_count(slide: Dict[str, Any]) -> int:
    text = "".join(_slide_visible_strings(slide)).replace("**", "")
    return len(re.sub(r"\s+", "", text))


def _primary_block_count(slide: Dict[str, Any]) -> int:
    slide_type = str(slide.get("type", "title_body"))
    if slide_type in {"cards", "process", "decomposition", "timeline"}:
        key = {"cards": "cards", "process": "steps", "decomposition": "items", "timeline": "events"}[slide_type]
        return len(_as_list(slide.get(key)))
    if slide_type in {"two_column", "comparison"}:
        return 2
    if slide_type == "diagram" and isinstance(slide.get("diagram"), dict):
        return len(_as_list(slide["diagram"].get("nodes")))
    if slide_type == "composite_diagram":
        return len(_as_list(slide.get("panels")))
    return max(1, len(_slide_bullets(slide)))


def _validate_style(style: Any, *, prefix: str, errors: List[str]) -> None:
    if style is not None and not isinstance(style, dict):
        errors.append(f"{prefix} must be an object")
        return
    if not isinstance(style, dict):
        return
    colors = style.get("colors", {})
    if colors is not None and not isinstance(colors, dict):
        errors.append(f"{prefix}.colors must be an object")
        return
    for key, color in colors.items():
        try:
            normalize_hex(color)
        except ValueError:
            errors.append(f"{prefix}.colors.{key} must be a six-digit RGB hex color")


def _validate_generation_policies(spec: Dict[str, Any], errors: List[str]) -> None:
    readability = spec.get("readability_policy")
    if readability is not None and not isinstance(readability, dict):
        errors.append("readability_policy must be an object")
    elif isinstance(readability, dict):
        for key in (
            "min_title_font",
            "min_takeaway_font",
            "min_module_font",
            "min_body_font",
            "min_diagram_font",
            "min_label_font",
        ):
            if key in readability:
                value = readability[key]
                if not isinstance(value, (int, float)) or isinstance(value, bool) or not 8 <= value <= 48:
                    errors.append(f"readability_policy.{key} must be a number from 8 to 48")
        if readability.get("overflow_action", "revise_or_split") not in {"revise_or_split", "split", "fail"}:
            errors.append("readability_policy.overflow_action must be revise_or_split, split, or fail")

    emphasis = spec.get("emphasis_policy")
    if emphasis is not None and not isinstance(emphasis, dict):
        errors.append("emphasis_policy must be an object")
    elif isinstance(emphasis, dict):
        for key in ("min_bold_ratio", "max_bold_ratio"):
            if key in emphasis:
                value = emphasis[key]
                if not isinstance(value, (int, float)) or isinstance(value, bool) or not 0 <= value <= 1:
                    errors.append(f"emphasis_policy.{key} must be a number from 0 to 1")
        if (
            isinstance(emphasis.get("min_bold_ratio"), (int, float))
            and isinstance(emphasis.get("max_bold_ratio"), (int, float))
            and emphasis["min_bold_ratio"] > emphasis["max_bold_ratio"]
        ):
            errors.append("emphasis_policy.min_bold_ratio cannot exceed max_bold_ratio")
        if "max_highlights_per_block" in emphasis:
            value = emphasis["max_highlights_per_block"]
            if not isinstance(value, int) or isinstance(value, bool) or not 1 <= value <= 8:
                errors.append("emphasis_policy.max_highlights_per_block must be an integer from 1 to 8")
        for key in ("enabled", "auto_emphasize_leading_labels"):
            if key in emphasis and not isinstance(emphasis[key], bool):
                errors.append(f"emphasis_policy.{key} must be true or false")

    density = spec.get("density_policy")
    if density is not None and not isinstance(density, dict):
        errors.append("density_policy must be an object")
    elif isinstance(density, dict):
        ranges = {
            "max_visible_characters": (80, 1200),
            "max_primary_blocks": (2, 10),
            "max_visual_focal_points": (1, 4),
        }
        for key, (low, high) in ranges.items():
            if key in density:
                value = density[key]
                if not isinstance(value, int) or isinstance(value, bool) or not low <= value <= high:
                    errors.append(f"density_policy.{key} must be an integer from {low} to {high}")


def _structured_items(slide: Dict[str, Any]) -> set[str]:
    """Collect comparable semantic labels without treating prose as exact structure."""

    values: List[Any] = []
    values.extend(_slide_bullets(slide))
    for key in ("cards", "steps", "items", "events"):
        for item in _as_list(slide.get(key)):
            if isinstance(item, dict):
                values.extend([item.get("heading"), item.get("label"), item.get("outcome")])
            else:
                values.append(item)
    diagrams: List[Any] = []
    if isinstance(slide.get("diagram"), dict):
        diagrams.append(slide["diagram"])
    for panel in _as_list(slide.get("panels")):
        if isinstance(panel, dict) and isinstance(panel.get("diagram"), dict):
            diagrams.append(panel["diagram"])
    for diagram in diagrams:
        for node in _as_list(diagram.get("nodes")):
            if isinstance(node, dict):
                values.append(node.get("label"))
    return {normalized for value in values if len(normalized := _normalized_text(value)) >= 2}


_NUMBER_MAP = {"一": "1", "二": "2", "三": "3", "四": "4", "1": "1", "2": "2", "3": "3", "4": "4"}


def _direction_number(value: Any) -> str | None:
    match = re.search(r"(?:任务|方向)\s*([一二三四1234])", str(value or ""))
    return _NUMBER_MAP.get(match.group(1)) if match else None


def _validate_diagram(
    diagram: Any,
    *,
    prefix: str,
    max_nodes: int,
    errors: List[str],
    warnings: List[str],
) -> None:
    if not isinstance(diagram, dict):
        errors.append(f"{prefix}.diagram must be an object")
        return

    kind = str(diagram.get("kind", "flow"))
    if kind not in SUPPORTED_DIAGRAM_KINDS:
        errors.append(f"{prefix}.diagram.kind is unsupported: {kind!r}")
    direction = str(
        diagram.get(
            "direction",
            "top_to_bottom" if kind in {"hierarchy", "architecture", "swimlane"} else "left_to_right",
        )
    )
    if direction not in SUPPORTED_DIAGRAM_DIRECTIONS:
        errors.append(f"{prefix}.diagram.direction is unsupported: {direction!r}")
    render = str(diagram.get("render", "auto"))
    if render not in SUPPORTED_DIAGRAM_RENDERS:
        errors.append(f"{prefix}.diagram.render is unsupported: {render!r}")
    elif render == "smartart" and kind not in {"flow", "hierarchy", "cycle"}:
        warnings.append(
            f"{prefix}.diagram.render=smartart is not native-supported for {kind}; native editable shapes will be kept"
        )

    nodes = diagram.get("nodes")
    if not isinstance(nodes, list) or len(nodes) < 2:
        errors.append(f"{prefix}.diagram.nodes must contain at least 2 items")
        return
    if len(nodes) > 12:
        errors.append(f"{prefix}.diagram.nodes must contain at most 12 items; split the diagram")
    elif len(nodes) > max_nodes:
        warnings.append(
            f"{prefix}.diagram has {len(nodes)} nodes, above max_nodes_per_slide={max_nodes}; consider splitting it"
        )

    node_ids: set[str] = set()
    node_roles: Dict[str, str] = {}
    node_lanes: Dict[str, str] = {}
    for node_index, node in enumerate(nodes, start=1):
        node_prefix = f"{prefix}.diagram.nodes[{node_index}]"
        if not isinstance(node, dict):
            errors.append(f"{node_prefix} must be an object")
            continue
        node_id = str(node.get("id") or "").strip()
        if not node_id:
            errors.append(f"{node_prefix}.id is required")
        elif node_id in node_ids:
            errors.append(f"{node_prefix}.id is duplicated: {node_id!r}")
        else:
            node_ids.add(node_id)
        if not str(node.get("label") or "").strip():
            errors.append(f"{node_prefix}.label is required")
        role = str(node.get("role", "default"))
        if role not in SUPPORTED_DIAGRAM_ROLES:
            errors.append(f"{node_prefix}.role is unsupported: {role!r}")
        if node_id:
            node_roles[node_id] = role
            if node.get("lane") is not None:
                node_lanes[node_id] = str(node.get("lane"))

    lanes = diagram.get("lanes", [])
    lane_ids: set[str] = set()
    if kind == "swimlane":
        if not isinstance(lanes, list) or not 2 <= len(lanes) <= 5:
            errors.append(f"{prefix}.diagram.lanes must contain 2 to 5 items for swimlane diagrams")
        else:
            for lane_index, lane in enumerate(lanes, start=1):
                if isinstance(lane, str):
                    lane_id = lane.strip()
                elif isinstance(lane, dict):
                    lane_id = str(lane.get("id") or lane.get("label") or "").strip()
                else:
                    lane_id = ""
                if not lane_id:
                    errors.append(f"{prefix}.diagram.lanes[{lane_index}] needs an id or label")
                elif lane_id in lane_ids:
                    errors.append(f"{prefix}.diagram.lanes[{lane_index}] is duplicated: {lane_id!r}")
                else:
                    lane_ids.add(lane_id)
            for node_id, lane_id in node_lanes.items():
                if lane_id not in lane_ids:
                    errors.append(f"{prefix}.diagram node {node_id!r} uses unknown lane {lane_id!r}")
            for node_id in node_ids - node_lanes.keys():
                warnings.append(f"{prefix}.diagram node {node_id!r} has no lane")

    edges = diagram.get("edges", [])
    if not isinstance(edges, list):
        errors.append(f"{prefix}.diagram.edges must be an array")
        return
    if not edges and kind not in {"flow", "cycle"}:
        errors.append(f"{prefix}.diagram.edges must not be empty for {kind} diagrams")
        return

    incident: Counter[str] = Counter()
    outgoing: Counter[str] = Counter()
    constraint_nodes = {node_id for node_id, role in node_roles.items() if role == "constraint"}
    constraint_edges: set[str] = set()
    for edge_index, edge in enumerate(edges, start=1):
        edge_prefix = f"{prefix}.diagram.edges[{edge_index}]"
        if not isinstance(edge, dict):
            errors.append(f"{edge_prefix} must be an object")
            continue
        source = str(edge.get("from") or "").strip()
        target = str(edge.get("to") or "").strip()
        edge_type = str(edge.get("type", "flow"))
        if edge_type not in SUPPORTED_EDGE_TYPES:
            errors.append(f"{edge_prefix}.type is unsupported: {edge_type!r}")
        if edge_type == "constraint":
            constraint_edges.update({source, target})
        if source not in node_ids:
            errors.append(f"{edge_prefix}.from references unknown node: {source!r}")
        if target not in node_ids:
            errors.append(f"{edge_prefix}.to references unknown node: {target!r}")
        if source and source == target:
            errors.append(f"{edge_prefix} cannot connect a node to itself")
        if source in node_ids and target in node_ids and source != target:
            incident[source] += 1
            incident[target] += 1
            outgoing[source] += 1

    if edges:
        for node_id in sorted(node_ids):
            if not incident[node_id]:
                warnings.append(f"{prefix}.diagram node {node_id!r} is isolated")
        for node_id, role in node_roles.items():
            if role == "decision" and outgoing[node_id] < 2:
                warnings.append(f"{prefix}.diagram decision node {node_id!r} should have at least 2 outgoing edges")
        for node_id in sorted(constraint_nodes - constraint_edges):
            warnings.append(
                f"{prefix}.diagram constraint node {node_id!r} has no edge with type='constraint'"
            )


def validate_spec(
    spec: Dict[str, Any],
    *,
    spec_path: Path | None = None,
    check_assets: bool = False,
) -> Tuple[List[str], List[str]]:
    errors: List[str] = []
    warnings: List[str] = []

    if spec.get("version", 1) != 1:
        errors.append("version must be 1")
    if spec.get("mode", "hybrid") not in {"native", "hybrid"}:
        errors.append("mode must be native or hybrid")
    if spec.get("aspect_ratio", "16:9") not in {"16:9", "4:3"}:
        errors.append("aspect_ratio must be 16:9 or 4:3")
    _validate_generation_policies(spec, errors)

    workflow = spec.get("workflow")
    two_stage_workflow = False
    if workflow is not None and not isinstance(workflow, dict):
        errors.append("workflow must be an object")
    elif isinstance(workflow, dict):
        workflow_mode = str(workflow.get("mode") or "").strip()
        if workflow_mode and workflow_mode != "two_stage":
            errors.append("workflow.mode must be 'two_stage'")
        two_stage_workflow = workflow_mode == "two_stage"
        if two_stage_workflow:
            if workflow.get("stage") != "generation":
                errors.append("workflow.stage must be 'generation' for a buildable two-stage spec")
            review_value = str(workflow.get("content_review") or "").strip()
            approval_value = str(workflow.get("approval_record") or "").strip()
            if not review_value:
                errors.append("workflow.content_review is required for a two-stage build")
            if not approval_value:
                errors.append("workflow.approval_record is required for a two-stage build")
            if spec_path is not None and review_value and approval_value:
                review_path = resolve_asset_path(spec_path, review_value)
                approval_path = resolve_asset_path(spec_path, approval_value)
                if not review_path.is_file():
                    errors.append(f"approved content review not found: {review_path}")
                if not approval_path.is_file():
                    errors.append(f"content approval record not found: {approval_path}")
                if review_path.is_file() and approval_path.is_file():
                    try:
                        approval = load_json(approval_path)
                    except ValueError as exc:
                        errors.append(str(exc))
                    else:
                        if approval.get("status") != "approved":
                            errors.append("content approval record status must be 'approved'")
                        if not str(approval.get("user_instruction") or "").strip():
                            errors.append("content approval record must preserve the user's approval instruction")
                        approved_review_value = str(approval.get("content_review") or "").strip()
                        if not approved_review_value:
                            errors.append("content approval record is missing content_review")
                        else:
                            approved_review_path = resolve_asset_path(approval_path, approved_review_value)
                            if approved_review_path != review_path:
                                errors.append("content approval record points to a different content review")
                        expected_digest = str(approval.get("content_sha256") or "").strip().lower()
                        if not re.fullmatch(r"[0-9a-f]{64}", expected_digest):
                            errors.append("content approval record has an invalid content_sha256")
                        else:
                            actual_digest = hashlib.sha256(review_path.read_bytes()).hexdigest()
                            if actual_digest != expected_digest:
                                errors.append(
                                    "approved content review changed after approval; return to Stage 1 and obtain approval again"
                                )

    content_profile = spec.get("content_profile", {})
    if content_profile is not None and not isinstance(content_profile, dict):
        errors.append("content_profile must be an object")
        content_profile = {}
    elif isinstance(content_profile, dict):
        summary_level = content_profile.get("summary_level", "standard")
        if summary_level not in SUMMARY_LEVEL_LIMITS:
            errors.append("content_profile.summary_level must be executive, standard, or technical")
        title_style = content_profile.get("title_style", "conclusion_first")
        if title_style not in {"conclusion_first", "topic"}:
            errors.append("content_profile.title_style must be conclusion_first or topic")
        for key, low, high in (("max_bullets", 2, 8), ("max_nodes_per_slide", 3, 12)):
            if key in content_profile:
                value = content_profile[key]
                if not isinstance(value, int) or isinstance(value, bool) or not low <= value <= high:
                    errors.append(f"content_profile.{key} must be an integer from {low} to {high}")
        if "source_traceability" in content_profile and not isinstance(
            content_profile["source_traceability"], bool
        ):
            errors.append("content_profile.source_traceability must be true or false")

    _validate_style(spec.get("style", {}), prefix="style", errors=errors)
    style_families = spec.get("style_families", {})
    if style_families is not None and not isinstance(style_families, dict):
        errors.append("style_families must be an object")
        style_families = {}
    elif isinstance(style_families, dict):
        for family_name, family_style in style_families.items():
            if not str(family_name).strip():
                errors.append("style_families cannot contain an empty family name")
            _validate_style(family_style, prefix=f"style_families.{family_name}", errors=errors)

    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        errors.append("slides must be a non-empty array")
        return errors, warnings

    seen_titles: Dict[str, int] = {}
    content_types = SUPPORTED_SLIDE_TYPES - {"cover", "section", "closing"}
    summary_level = str(content_profile.get("summary_level", "standard"))
    limits = SUMMARY_LEVEL_LIMITS.get(summary_level, SUMMARY_LEVEL_LIMITS["standard"])
    max_bullets = int(content_profile.get("max_bullets", limits["max_bullets"]))
    max_nodes = int(content_profile.get("max_nodes_per_slide", 7))
    source_traceability = bool(content_profile.get("source_traceability", False))
    density_policy = merged_policy(spec, "density_policy", DEFAULT_DENSITY_POLICY)
    max_visible_characters = int(density_policy["max_visible_characters"])
    max_primary_blocks = int(density_policy["max_primary_blocks"])

    for index, slide in enumerate(slides, start=1):
        prefix = f"slide {index}"
        if not isinstance(slide, dict):
            errors.append(f"{prefix} must be an object")
            continue
        slide_type = slide.get("type", "title_body")
        if slide_type not in SUPPORTED_SLIDE_TYPES:
            errors.append(f"{prefix}.type is unsupported: {slide_type!r}")
            continue
        style_family = slide.get("style_family")
        if style_family is not None and str(style_family) not in style_families:
            errors.append(f"{prefix}.style_family references unknown family: {style_family!r}")
        _validate_style(slide.get("style"), prefix=f"{prefix}.style", errors=errors)
        semantic_intent = slide.get("semantic_intent")
        if semantic_intent is not None and str(semantic_intent) not in SUPPORTED_SEMANTIC_INTENTS:
            errors.append(f"{prefix}.semantic_intent is unsupported: {semantic_intent!r}")
        if slide_type == "process" and semantic_intent == "decomposition":
            warnings.append(f"{prefix} uses process for decomposition; use type='decomposition' to avoid false sequence")
        if slide_type == "decomposition" and semantic_intent in {"sequence", "cycle"}:
            warnings.append(f"{prefix} uses decomposition for {semantic_intent}; use process or diagram")
        title = str(slide.get("title") or "").strip()
        if not title:
            errors.append(f"{prefix}.title is required")
        elif title in seen_titles and slide_type not in {"section"}:
            warnings.append(f"{prefix} repeats the title from slide {seen_titles[title]}: {title}")
        else:
            seen_titles[title] = index

        if slide_type in content_types and not str(slide.get("takeaway") or "").strip():
            warnings.append(f"{prefix} has no takeaway; cross-slide logic may be weaker")

        source_refs = slide.get("source_refs", [])
        if source_refs is not None and not isinstance(source_refs, list):
            errors.append(f"{prefix}.source_refs must be an array")
        elif isinstance(source_refs, list) and not all(str(item or "").strip() for item in source_refs):
            errors.append(f"{prefix}.source_refs cannot contain empty values")
        evidence = slide.get("evidence", [])
        if evidence is not None and not isinstance(evidence, list):
            errors.append(f"{prefix}.evidence must be an array")
        elif isinstance(evidence, list) and not all(str(item or "").strip() for item in evidence):
            errors.append(f"{prefix}.evidence cannot contain empty values")
        if source_traceability and slide_type in content_types:
            if not source_refs and not evidence:
                warnings.append(f"{prefix} has no source_refs or evidence while source_traceability is enabled")

        fact_ids = slide.get("fact_ids", [])
        if fact_ids is not None and not isinstance(fact_ids, list):
            errors.append(f"{prefix}.fact_ids must be an array")
            fact_ids = []
        elif isinstance(fact_ids, list) and not all(str(item or "").strip() for item in fact_ids):
            errors.append(f"{prefix}.fact_ids cannot contain empty values")
        must_show_terms = slide.get("must_show_terms", [])
        if must_show_terms is not None and not isinstance(must_show_terms, list):
            errors.append(f"{prefix}.must_show_terms must be an array")
            must_show_terms = []
        elif isinstance(must_show_terms, list) and not all(str(item or "").strip() for item in must_show_terms):
            errors.append(f"{prefix}.must_show_terms cannot contain empty values")
        if two_stage_workflow and slide_type in content_types:
            if not fact_ids:
                errors.append(f"{prefix}.fact_ids are required for content slides in a two-stage build")
            if not must_show_terms:
                errors.append(f"{prefix}.must_show_terms are required for content slides in a two-stage build")
        visible_text = _normalized_text(" ".join(_slide_visible_strings(slide)))
        missing_terms = [
            str(term)
            for term in must_show_terms
            if _normalized_text(term) and _normalized_text(term) not in visible_text
        ]
        if missing_terms:
            errors.append(
                f"{prefix} omits must_show_terms from visible slide text: {', '.join(missing_terms)}"
            )

        if slide_type in content_types:
            visible_characters = _visible_character_count(slide)
            if visible_characters > max_visible_characters:
                warnings.append(
                    f"{prefix} has {visible_characters} visible characters, above "
                    f"density_policy.max_visible_characters={max_visible_characters}; "
                    "compress syntax, restructure, or split instead of shrinking fonts"
                )
            primary_blocks = _primary_block_count(slide)
            if primary_blocks > max_primary_blocks:
                warnings.append(
                    f"{prefix} has {primary_blocks} primary blocks, above "
                    f"density_policy.max_primary_blocks={max_primary_blocks}; reduce competing focal points"
                )

        bullets = _slide_bullets(slide)
        if len(bullets) > max_bullets:
            warnings.append(f"{prefix} has {len(bullets)} bullets, above content_profile.max_bullets={max_bullets}")
        for bullet_index, bullet in enumerate(bullets, start=1):
            compact_length = len(re.sub(r"\s+", "", bullet))
            if compact_length > limits["max_bullet_chars"]:
                warnings.append(
                    f"{prefix}.bullets[{bullet_index}] has {compact_length} characters; "
                    f"{summary_level} mode recommends at most {limits['max_bullet_chars']}"
                )

        if slide_type in {"two_column", "comparison"}:
            for side in ("left", "right"):
                if not isinstance(slide.get(side), dict):
                    errors.append(f"{prefix}.{side} must be an object")
        elif slide_type == "cards":
            cards = slide.get("cards")
            if not isinstance(cards, list) or not 2 <= len(cards) <= 6:
                errors.append(f"{prefix}.cards must contain 2 to 6 items")
        elif slide_type == "decomposition":
            items = slide.get("items")
            if not isinstance(items, list) or not 2 <= len(items) <= 6:
                errors.append(f"{prefix}.items must contain 2 to 6 items")
            elif not all(isinstance(item, dict) and str(item.get("heading") or "").strip() for item in items):
                errors.append(f"{prefix}.items must be objects with non-empty headings")
        elif slide_type == "process":
            steps = slide.get("steps")
            if not isinstance(steps, list) or not 2 <= len(steps) <= 6:
                errors.append(f"{prefix}.steps must contain 2 to 6 items")
        elif slide_type == "diagram":
            _validate_diagram(
                slide.get("diagram"),
                prefix=prefix,
                max_nodes=max_nodes,
                errors=errors,
                warnings=warnings,
            )
        elif slide_type == "composite_diagram":
            panels = slide.get("panels")
            if not isinstance(panels, list) or not 2 <= len(panels) <= 3:
                errors.append(f"{prefix}.panels must contain 2 to 3 items")
            else:
                for panel_index, panel in enumerate(panels, start=1):
                    panel_prefix = f"{prefix}.panels[{panel_index}]"
                    if not isinstance(panel, dict):
                        errors.append(f"{panel_prefix} must be an object")
                        continue
                    if not str(panel.get("heading") or "").strip():
                        errors.append(f"{panel_prefix}.heading is required")
                    ratio = panel.get("width_ratio", 1)
                    if not isinstance(ratio, (int, float)) or isinstance(ratio, bool) or ratio <= 0:
                        errors.append(f"{panel_prefix}.width_ratio must be a positive number")
                    _validate_diagram(
                        panel.get("diagram"),
                        prefix=panel_prefix,
                        max_nodes=max_nodes,
                        errors=errors,
                        warnings=warnings,
                    )
        elif slide_type == "timeline":
            events = slide.get("events")
            if not isinstance(events, list) or not 2 <= len(events) <= 6:
                errors.append(f"{prefix}.events must contain 2 to 6 items")
        elif slide_type == "chart":
            chart = slide.get("chart")
            if not isinstance(chart, dict):
                errors.append(f"{prefix}.chart must be an object")
            else:
                chart_type = chart.get("type", "column")
                if chart_type not in {"column", "bar", "line", "pie", "doughnut"}:
                    errors.append(f"{prefix}.chart.type is unsupported: {chart_type!r}")
                categories = chart.get("categories")
                series = chart.get("series")
                if not isinstance(categories, list) or not categories:
                    errors.append(f"{prefix}.chart.categories must be a non-empty array")
                if not isinstance(series, list) or not series:
                    errors.append(f"{prefix}.chart.series must be a non-empty array")
                elif isinstance(categories, list):
                    for series_index, item in enumerate(series, start=1):
                        values = item.get("values") if isinstance(item, dict) else None
                        if not isinstance(values, list) or len(values) != len(categories):
                            errors.append(
                                f"{prefix}.chart.series[{series_index}] values must match category count"
                            )
                        elif not all(isinstance(value, (int, float)) for value in values):
                            errors.append(f"{prefix}.chart.series[{series_index}] values must be numeric")
                if chart_type in {"pie", "doughnut"} and isinstance(series, list) and len(series) != 1:
                    errors.append(f"{prefix}.{chart_type} chart must contain exactly one series")
        elif slide_type == "table":
            columns = slide.get("columns")
            rows = slide.get("rows")
            if not isinstance(columns, list) or not columns:
                errors.append(f"{prefix}.columns must be a non-empty array")
            if not isinstance(rows, list) or not rows:
                errors.append(f"{prefix}.rows must be a non-empty array")
            elif isinstance(columns, list):
                for row_index, row in enumerate(rows, start=1):
                    if not isinstance(row, list) or len(row) != len(columns):
                        errors.append(f"{prefix}.rows[{row_index}] must match column count")
        elif slide_type == "image_text" and not slide.get("image"):
            errors.append(f"{prefix}.image is required for image_text")

        if check_assets and spec_path is not None:
            for image in _image_entries(slide):
                raw_path = image.get("path")
                if not isinstance(raw_path, str) or not raw_path.strip():
                    errors.append(f"{prefix}.image.path is required")
                    continue
                if raw_path.startswith(("http://", "https://")):
                    warnings.append(f"{prefix} uses a remote image; download it before deterministic builds")
                    continue
                resolved = resolve_asset_path(spec_path, raw_path)
                if not resolved.is_file():
                    errors.append(f"{prefix} image not found: {resolved}")

    return errors, warnings


def content_quality_report(spec: Dict[str, Any]) -> Dict[str, Any]:
    """Inspect narrative density, traceability, repetition, and diagram semantics."""

    errors, validation_warnings = validate_spec(spec)
    profile = _content_profile(spec)
    summary_level = str(profile.get("summary_level", "standard"))
    limits = SUMMARY_LEVEL_LIMITS.get(summary_level, SUMMARY_LEVEL_LIMITS["standard"])
    max_bullets = int(profile.get("max_bullets", limits["max_bullets"]))
    title_style = str(profile.get("title_style", "conclusion_first"))
    source_traceability = bool(profile.get("source_traceability", False))
    density_policy = merged_policy(spec, "density_policy", DEFAULT_DENSITY_POLICY)
    max_visible_characters = int(density_policy["max_visible_characters"])
    max_primary_blocks = int(density_policy["max_primary_blocks"])
    generic_titles = {
        "背景",
        "概述",
        "介绍",
        "方案",
        "技术方案",
        "流程",
        "架构",
        "总结",
        "结论",
        "overview",
        "introduction",
        "solution",
        "summary",
    }
    report_warnings: List[str] = list(validation_warnings)
    slides = spec.get("slides") if isinstance(spec.get("slides"), list) else []
    slide_reports: List[Dict[str, Any]] = []
    previous_content: List[Tuple[int, str, str]] = []
    previous_structures: List[Tuple[int, str, set[str]]] = []
    content_types = SUPPORTED_SLIDE_TYPES - {"cover", "section", "closing"}

    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        slide_type = str(slide.get("type", "title_body"))
        title = str(slide.get("title") or "").strip()
        takeaway = str(slide.get("takeaway") or "").strip()
        bullets = _slide_bullets(slide)
        structured_items = _structured_items(slide)
        issues: List[str] = []

        normalized_title = _normalized_text(title)
        if slide_type in content_types and title_style == "conclusion_first":
            if normalized_title in generic_titles or len(normalized_title) <= 4:
                issues.append("title is topic-like; use a conclusion-first title")
        if slide_type in content_types and not takeaway:
            issues.append("missing takeaway")
        if len(bullets) > max_bullets:
            issues.append(f"{len(bullets)} bullets exceed the preferred maximum of {max_bullets}")
        long_bullets = [
            bullet for bullet in bullets if len(re.sub(r"\s+", "", bullet)) > limits["max_bullet_chars"]
        ]
        if long_bullets:
            issues.append(f"{len(long_bullets)} bullet(s) are too long for {summary_level} mode")
        visible_characters = _visible_character_count(slide)
        primary_blocks = _primary_block_count(slide)
        if slide_type in content_types and visible_characters > max_visible_characters:
            issues.append(
                f"{visible_characters} visible characters exceed the preferred maximum of "
                f"{max_visible_characters}; revise or split instead of shrinking fonts"
            )
        if slide_type in content_types and primary_blocks > max_primary_blocks:
            issues.append(
                f"{primary_blocks} primary blocks exceed the preferred maximum of {max_primary_blocks}"
            )
        if source_traceability and slide_type in content_types:
            if not slide.get("source_refs") and not slide.get("evidence"):
                issues.append("missing source_refs/evidence")
        must_show_terms = [
            str(item).strip()
            for item in _as_list(slide.get("must_show_terms"))
            if str(item or "").strip()
        ]
        visible_text = _normalized_text(" ".join(_slide_visible_strings(slide)))
        missing_required_terms = [
            term for term in must_show_terms if _normalized_text(term) not in visible_text
        ]
        if missing_required_terms:
            issues.append(
                "approved visible terms are missing: " + ", ".join(missing_required_terms)
            )

        comparison_text = takeaway or title
        for prior_index, prior_title, prior_text in previous_content:
            similarity = _text_similarity(comparison_text, prior_text)
            if similarity >= 0.78:
                issues.append(
                    f"message is very similar to slide {prior_index} ({prior_title!r}, similarity={similarity:.2f})"
                )
                break
        if slide_type in content_types and comparison_text:
            previous_content.append((index, title, comparison_text))

        for prior_index, prior_title, prior_items in previous_structures:
            overlap = structured_items & prior_items
            denominator = min(len(structured_items), len(prior_items))
            if len(overlap) >= 3 and denominator and len(overlap) / denominator >= 0.6:
                issues.append(
                    f"structured content duplicates slide {prior_index} ({prior_title!r}); "
                    f"{len(overlap)} repeated item(s)"
                )
                break
        if structured_items:
            previous_structures.append((index, title, structured_items))

        semantic_intent = str(slide.get("semantic_intent") or "")
        if slide_type in {"process", "decomposition", "diagram", "composite_diagram"} and not semantic_intent:
            issues.append("semantic_intent is missing; declare the source relationship before choosing the visual")
        if slide_type == "process" and semantic_intent == "decomposition":
            issues.append("false sequence risk: decomposition is rendered as a process")
        if slide_type == "decomposition" and semantic_intent in {"sequence", "cycle"}:
            issues.append(f"relation mismatch: {semantic_intent} is rendered as decomposition")

        title_direction = _direction_number(title)
        notes_direction = _direction_number(slide.get("notes"))
        if title_direction and notes_direction and title_direction != notes_direction:
            issues.append(
                f"speaker notes reference direction/task {notes_direction}, but title references {title_direction}"
            )

        if slide_type == "cards":
            numbered_headings = 0
            for card in _as_list(slide.get("cards")):
                heading = str(card.get("heading") or "") if isinstance(card, dict) else ""
                if re.match(r"^\s*(?:0?\d+|[一二三四五六])[\s.、:：-]+", heading):
                    numbered_headings += 1
            if numbered_headings >= 2 and semantic_intent != "sequence":
                issues.append(
                    "numbered cards may imply a false sequence; use decomposition or declare semantic_intent='sequence'"
                )

        diagrams: List[Dict[str, Any]] = []
        if slide_type == "diagram" and isinstance(slide.get("diagram"), dict):
            diagrams.append(slide["diagram"])
        if slide_type == "composite_diagram":
            for panel in _as_list(slide.get("panels")):
                if isinstance(panel, dict) and isinstance(panel.get("diagram"), dict):
                    diagrams.append(panel["diagram"])
        for diagram in diagrams:
            node_count = len(diagram.get("nodes", [])) if isinstance(diagram.get("nodes"), list) else 0
            edge_count = len(diagram.get("edges", [])) if isinstance(diagram.get("edges"), list) else 0
            max_nodes = int(profile.get("max_nodes_per_slide", 7))
            if node_count > max_nodes:
                issues.append(f"diagram has {node_count} nodes; split or simplify above {max_nodes}")
            if edge_count == 0 and diagram.get("kind") not in {"flow", "cycle"}:
                issues.append("diagram has no explicit semantic edges")
            edges = diagram.get("edges") if isinstance(diagram.get("edges"), list) else []
            if diagram.get("kind") == "cycle" and edges and not any(
                isinstance(edge, dict) and edge.get("type") == "feedback" for edge in edges
            ):
                issues.append("cycle has explicit edges but no type='feedback' edge")
            nodes = diagram.get("nodes") if isinstance(diagram.get("nodes"), list) else []
            labels = {
                str(node.get("id")): str(node.get("label") or "")
                for node in nodes
                if isinstance(node, dict) and node.get("id") is not None
            }
            mapping_ids = {
                node_id for node_id, label in labels.items() if re.search(r"映射|调度|mapping|schedul", label, re.I)
            }
            hardware_ids = {
                node_id
                for node_id, label in labels.items()
                if re.search(r"硬件|芯片|阵列|存储|互连|资源|hardware|chip|array", label, re.I)
            }
            has_constraint_input = any(
                isinstance(edge, dict)
                and edge.get("type") == "constraint"
                and str(edge.get("from")) in hardware_ids
                and str(edge.get("to")) in mapping_ids
                for edge in edges
            )
            if mapping_ids and hardware_ids and not has_constraint_input:
                issues.append(
                    "mapping and hardware appear together without a typed hardware-to-mapping constraint edge; "
                    "verify that architecture constraints enter the optimizer"
                )

            feedback_edges = [
                edge
                for edge in edges
                if isinstance(edge, dict)
                and (
                    edge.get("type") == "feedback"
                    or re.search(r"反馈|反哺|feedback", str(edge.get("label") or ""), re.I)
                )
            ]
            context = " ".join(
                [
                    str(slide.get("takeaway") or ""),
                    str(slide.get("notes") or ""),
                    " ".join(str(item) for item in _as_list(slide.get("evidence"))),
                ]
            )
            feedback_terms = {
                term
                for term in ("平台参数", "算子实现", "映射策略", "调度配置", "编译映射", "运行时")
                if term in context
            }
            feedback_targets = {str(edge.get("to")) for edge in feedback_edges}
            aggregate_target = any(
                re.search(r"多层|协同|综合|联合", labels.get(target, "")) for target in feedback_targets
            )
            if len(feedback_terms) >= 2 and feedback_edges and len(feedback_targets) < 2 and not aggregate_target:
                issues.append(
                    "evidence names multiple feedback targets, but the diagram feeds back to only one narrow node"
                )

        for issue in issues:
            report_warnings.append(f"slide {index}: {issue}")
        slide_reports.append(
            {
                "slide": index,
                "type": slide_type,
                "title": title,
                "takeaway": bool(takeaway),
                "bullet_count": len(bullets),
                "visible_characters": visible_characters,
                "primary_blocks": primary_blocks,
                "fact_ids": len(slide.get("fact_ids", []))
                if isinstance(slide.get("fact_ids"), list)
                else 0,
                "must_show_terms": len(must_show_terms),
                "missing_required_terms": missing_required_terms,
                "source_refs": len(slide.get("source_refs", []))
                if isinstance(slide.get("source_refs"), list)
                else 0,
                "issues": issues,
            }
        )

    report_warnings = list(dict.fromkeys(report_warnings))
    deduction = min(60, len(errors) * 15 + len(report_warnings) * 3)
    score = max(0, 100 - deduction)
    return {
        "valid": not errors,
        "passed": not errors and score >= 80,
        "score": score,
        "summary_level": summary_level,
        "slides": len(slides),
        "errors": errors,
        "warnings": report_warnings,
        "slide_reports": slide_reports,
    }
