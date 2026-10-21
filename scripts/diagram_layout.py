#!/usr/bin/env python3
"""Dependency-free semantic diagram layout for editable PowerPoint shapes."""

from __future__ import annotations

from collections import defaultdict, deque
import math
from typing import Any, Dict, List


def _normalized_edges(diagram: Dict[str, Any], node_ids: List[str]) -> List[Dict[str, Any]]:
    edges = [dict(edge) for edge in diagram.get("edges", []) if isinstance(edge, dict)]
    kind = str(diagram.get("kind", "flow"))
    if not edges and kind in {"flow", "cycle"}:
        for index in range(len(node_ids) - 1):
            edges.append({"from": node_ids[index], "to": node_ids[index + 1]})
        if kind == "cycle" and len(node_ids) > 2:
            edges.append({"from": node_ids[-1], "to": node_ids[0]})
    return edges


def _node_levels(nodes: List[Dict[str, Any]], edges: List[Dict[str, Any]]) -> Dict[str, int]:
    node_ids = [str(node["id"]) for node in nodes]
    incoming: Dict[str, int] = {node_id: 0 for node_id in node_ids}
    outgoing: Dict[str, List[str]] = defaultdict(list)
    for edge in edges:
        source = str(edge.get("from") or "")
        target = str(edge.get("to") or "")
        if source in incoming and target in incoming and source != target:
            outgoing[source].append(target)
            incoming[target] += 1

    explicit = {
        str(node["id"]): max(0, int(node["level"]))
        for node in nodes
        if isinstance(node.get("level"), int) and not isinstance(node.get("level"), bool)
    }
    levels: Dict[str, int] = dict(explicit)
    roots = [node_id for node_id in node_ids if incoming[node_id] == 0]
    if not roots and node_ids:
        roots = [node_ids[0]]
    queue = deque(roots)
    for root in roots:
        levels.setdefault(root, 0)

    visits: Dict[str, int] = defaultdict(int)
    while queue:
        source = queue.popleft()
        visits[source] += 1
        if visits[source] > len(node_ids):
            continue
        for target in outgoing[source]:
            proposed = levels.get(source, 0) + 1
            if target not in explicit and proposed > levels.get(target, -1):
                levels[target] = proposed
            queue.append(target)

    next_level = max(levels.values(), default=-1) + 1
    for node_id in node_ids:
        if node_id not in levels:
            levels[node_id] = next_level
            next_level += 1
    return levels


def _linear_layout(
    nodes: List[Dict[str, Any]],
    edges: List[Dict[str, Any]],
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    direction: str,
) -> Dict[str, Dict[str, float]]:
    levels = _node_levels(nodes, edges)
    grouped: Dict[int, List[Dict[str, Any]]] = defaultdict(list)
    for node in nodes:
        grouped[levels[str(node["id"])]].append(node)
    ordered_levels = sorted(grouped)
    if not ordered_levels:
        return {}

    positions: Dict[str, Dict[str, float]] = {}
    level_count = len(ordered_levels)
    max_group = max(len(grouped[level]) for level in ordered_levels)
    major_gap = 0.28
    minor_gap = 0.24

    if direction == "top_to_bottom":
        node_h = min(0.86, max(0.58, (h - major_gap * (level_count - 1)) / level_count))
        level_step = (h - node_h) / max(1, level_count - 1)
        for level_index, level in enumerate(ordered_levels):
            group = grouped[level]
            node_w = min(2.25, max(1.35, (w - minor_gap * (len(group) - 1)) / len(group)))
            used_w = len(group) * node_w + (len(group) - 1) * minor_gap
            start_x = x + (w - used_w) / 2
            for item_index, node in enumerate(group):
                positions[str(node["id"])] = {
                    "x": start_x + item_index * (node_w + minor_gap),
                    "y": y + level_index * level_step,
                    "w": node_w,
                    "h": node_h,
                }
    else:
        node_w = min(2.08, max(1.28, (w - major_gap * (level_count - 1)) / level_count))
        level_step = (w - node_w) / max(1, level_count - 1)
        for level_index, level in enumerate(ordered_levels):
            group = grouped[level]
            node_h = min(0.9, max(0.58, (h - minor_gap * (len(group) - 1)) / len(group)))
            used_h = len(group) * node_h + (len(group) - 1) * minor_gap
            start_y = y + (h - used_h) / 2
            for item_index, node in enumerate(group):
                positions[str(node["id"])] = {
                    "x": x + level_index * level_step,
                    "y": start_y + item_index * (node_h + minor_gap),
                    "w": node_w,
                    "h": node_h,
                }
    return positions


def _cycle_layout(
    nodes: List[Dict[str, Any]], *, x: float, y: float, w: float, h: float
) -> Dict[str, Dict[str, float]]:
    count = len(nodes)
    node_w = min(1.82, max(1.25, w / 4.7))
    node_h = min(0.82, max(0.58, h / 5.2))
    center_x = x + w / 2
    center_y = y + h / 2
    radius_x = max(0.1, (w - node_w) * 0.42)
    radius_y = max(0.1, (h - node_h) * 0.39)
    positions: Dict[str, Dict[str, float]] = {}
    for index, node in enumerate(nodes):
        angle = -math.pi / 2 + 2 * math.pi * index / count
        positions[str(node["id"])] = {
            "x": center_x + radius_x * math.cos(angle) - node_w / 2,
            "y": center_y + radius_y * math.sin(angle) - node_h / 2,
            "w": node_w,
            "h": node_h,
        }
    return positions


def _lane_values(diagram: Dict[str, Any]) -> List[Dict[str, str]]:
    values: List[Dict[str, str]] = []
    for lane in diagram.get("lanes", []):
        if isinstance(lane, str):
            values.append({"id": lane, "label": lane})
        elif isinstance(lane, dict):
            lane_id = str(lane.get("id") or lane.get("label") or "")
            values.append({"id": lane_id, "label": str(lane.get("label") or lane_id)})
    return values


def _node_order(node: Dict[str, Any]) -> tuple[int, str]:
    try:
        order = int(node.get("order", 9999))
    except (TypeError, ValueError):
        order = 9999
    return order, str(node.get("id", ""))


def _swimlane_layout(
    diagram: Dict[str, Any],
    nodes: List[Dict[str, Any]],
    *,
    x: float,
    y: float,
    w: float,
    h: float,
) -> tuple[Dict[str, Dict[str, float]], List[Dict[str, float | str]]]:
    lanes = _lane_values(diagram)
    lane_h = h / max(1, len(lanes))
    label_w = min(1.18, w * 0.14)
    content_x = x + label_w
    content_w = w - label_w
    positions: Dict[str, Dict[str, float]] = {}
    lane_boxes: List[Dict[str, float | str]] = []
    for lane_index, lane in enumerate(lanes):
        lane_y = y + lane_index * lane_h
        lane_boxes.append(
            {
                "id": lane["id"],
                "label": lane["label"],
                "x": x,
                "y": lane_y,
                "w": w,
                "h": lane_h,
                "label_w": label_w,
            }
        )
        lane_nodes = [node for node in nodes if str(node.get("lane") or "") == lane["id"]]
        lane_nodes.sort(key=_node_order)
        if not lane_nodes:
            continue
        gap = 0.25
        node_w = min(1.86, max(1.15, (content_w - gap * (len(lane_nodes) + 1)) / len(lane_nodes)))
        node_h = min(0.68, max(0.5, lane_h - 0.28))
        step = content_w / len(lane_nodes)
        for node_index, node in enumerate(lane_nodes):
            center_x = content_x + step * (node_index + 0.5)
            positions[str(node["id"])] = {
                "x": center_x - node_w / 2,
                "y": lane_y + (lane_h - node_h) / 2,
                "w": node_w,
                "h": node_h,
            }
    return positions, lane_boxes


def layout_diagram(
    diagram: Dict[str, Any],
    *,
    x: float,
    y: float,
    w: float,
    h: float,
) -> Dict[str, Any]:
    """Return native-shape positions for a validated semantic diagram."""

    nodes = [dict(node) for node in diagram.get("nodes", []) if isinstance(node, dict)]
    node_ids = [str(node["id"]) for node in nodes]
    edges = _normalized_edges(diagram, node_ids)
    kind = str(diagram.get("kind", "flow"))
    default_direction = "top_to_bottom" if kind in {"hierarchy", "architecture", "swimlane"} else "left_to_right"
    direction = str(diagram.get("direction", default_direction))
    lane_boxes: List[Dict[str, float | str]] = []

    if kind == "cycle":
        positions = _cycle_layout(nodes, x=x, y=y, w=w, h=h)
    elif kind == "swimlane":
        positions, lane_boxes = _swimlane_layout(diagram, nodes, x=x, y=y, w=w, h=h)
    else:
        positions = _linear_layout(nodes, edges, x=x, y=y, w=w, h=h, direction=direction)

    return {
        "kind": kind,
        "direction": direction,
        "nodes": positions,
        "edges": edges,
        "lanes": lane_boxes,
    }
