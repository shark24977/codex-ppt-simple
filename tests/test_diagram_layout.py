from __future__ import annotations

from pathlib import Path
import sys
import unittest


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from diagram_layout import layout_diagram  # noqa: E402


class DiagramLayoutTest(unittest.TestCase):
    def test_flow_without_edges_is_chained(self):
        diagram = {
            "kind": "flow",
            "nodes": [
                {"id": "a", "label": "A"},
                {"id": "b", "label": "B"},
                {"id": "c", "label": "C"},
            ],
            "edges": [],
        }
        layout = layout_diagram(diagram, x=0, y=0, w=10, h=4)
        self.assertEqual([(edge["from"], edge["to"]) for edge in layout["edges"]], [("a", "b"), ("b", "c")])
        self.assertLess(layout["nodes"]["a"]["x"], layout["nodes"]["b"]["x"])

    def test_cycle_closes_last_node_to_first(self):
        diagram = {
            "kind": "cycle",
            "nodes": [
                {"id": "a", "label": "A"},
                {"id": "b", "label": "B"},
                {"id": "c", "label": "C"},
            ],
            "edges": [],
        }
        layout = layout_diagram(diagram, x=0, y=0, w=10, h=4)
        self.assertEqual(layout["edges"][-1], {"from": "c", "to": "a"})
        self.assertEqual(len(layout["nodes"]), 3)

    def test_hierarchy_uses_distinct_vertical_levels(self):
        diagram = {
            "kind": "hierarchy",
            "direction": "top_to_bottom",
            "nodes": [
                {"id": "root", "label": "Root"},
                {"id": "left", "label": "Left"},
                {"id": "right", "label": "Right"},
            ],
            "edges": [
                {"from": "root", "to": "left"},
                {"from": "root", "to": "right"},
            ],
        }
        layout = layout_diagram(diagram, x=0, y=0, w=10, h=5)
        self.assertLess(layout["nodes"]["root"]["y"], layout["nodes"]["left"]["y"])
        self.assertEqual(layout["nodes"]["left"]["y"], layout["nodes"]["right"]["y"])

    def test_swimlane_keeps_nodes_in_their_lane(self):
        diagram = {
            "kind": "swimlane",
            "lanes": [
                {"id": "team-a", "label": "Team A"},
                {"id": "team-b", "label": "Team B"},
            ],
            "nodes": [
                {"id": "a", "label": "A", "lane": "team-a"},
                {"id": "b", "label": "B", "lane": "team-b"},
            ],
            "edges": [{"from": "a", "to": "b"}],
        }
        layout = layout_diagram(diagram, x=0, y=0, w=10, h=4)
        self.assertEqual(len(layout["lanes"]), 2)
        self.assertLess(layout["nodes"]["a"]["y"], layout["nodes"]["b"]["y"])


if __name__ == "__main__":
    unittest.main()
