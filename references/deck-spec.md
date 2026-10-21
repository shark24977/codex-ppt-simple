# Editable deck specification

Read this when creating or modifying `deck_spec.json`.

## Top-level fields

- `version`: currently `1`.
- `title`: deck title.
- `language`: presentation language; defaults to `zh-CN`.
- `mode`: `native` or `hybrid`. Both keep information editable; hybrid may add replaceable raster illustrations.
- `aspect_ratio`: `16:9` or `4:3`. A reference PPTX keeps its own size unless `force_aspect_ratio` is true.
- `content_profile`: optional summary, traceability, and diagram-density rules. See `content-and-diagrams.md`.
- `readability_policy`: role-specific minimum font sizes and the required response to overflow.
- `emphasis_policy`: controls selective bold emphasis; visible strings may use `**bold**` markers, which the builder removes while creating bold native runs.
- `density_policy`: limits visible characters, primary blocks, and competing focal points before font reduction is considered.
- `workflow`: required for new two-stage builds after content approval. It records the approved review and approval record.
- `style`: optional overrides merged onto the extracted or default profile.
- `style_families`: optional named style overrides for page families such as `research_task` or `technical_architecture`.
- `slides`: ordered slide objects.

Use paths relative to the spec file. Do not put important facts, labels, or numerical data inside image assets.

For a two-stage build, include:

```json
{
  "workflow": {
    "mode": "two_stage",
    "stage": "generation",
    "content_review": "content_review.json",
    "approval_record": "content_approval.json"
  }
}
```

For ordinary 16:9 working and technical decks, use or inherit:

```json
{
  "readability_policy": {
    "min_title_font": 26,
    "min_takeaway_font": 17,
    "min_module_font": 15,
    "min_body_font": 13.5,
    "min_diagram_font": 12,
    "min_label_font": 11,
    "overflow_action": "revise_or_split"
  },
  "emphasis_policy": {
    "enabled": true,
    "min_bold_ratio": 0.10,
    "max_bold_ratio": 0.30,
    "max_highlights_per_block": 3,
    "auto_emphasize_leading_labels": true
  },
  "density_policy": {
    "max_visible_characters": 420,
    "max_primary_blocks": 5,
    "max_visual_focal_points": 2
  }
}
```

The builder treats these font sizes as floors even when a reference style contains smaller values. If content cannot fit at the floors, shorten redundant syntax, restructure the page, or return to Stage 1 and split it. Never resolve overflow by silently shrinking below the policy.

Use `**...**` only for information hierarchy, not decoration. Good emphasis targets are conclusions, named objects, mechanisms, metrics, deliverables, innovations, and decisive constraints. Each block should normally contain one to three emphasized spans; do not bold whole paragraphs. A leading label followed by `:` or `：` is automatically emphasized when `auto_emphasize_leading_labels` is enabled.

The validator loads the approval record and verifies that its SHA-256 digest matches `content_review.json`. Do not create or refresh the approval record until the user explicitly approves the current review.

## Shared slide fields

- `type`: one of the supported layouts below.
- `title`: visible slide title.
- `takeaway`: the one sentence the audience should retain; it can be visible as a callout or used to guide the layout.
- `notes`: optional speaker notes.
- `kicker`: optional small overline.
- `footer`: optional footer text.
- `show_slide_number`: defaults to the deck-level setting.
- `image`: optional illustration object with `path`, `alt`, and optional `fit: contain`.
- `evidence`: optional facts supporting the page conclusion; detailed evidence is appended to speaker notes by default.
- `source_refs`: optional source-page or section references appended to speaker notes.
- `fact_ids`: source fact IDs carried forward from the approved content review. Required on content slides in a two-stage build.
- `must_show_terms`: exact terms that must remain in visible slide text. Required on content slides in a two-stage build when source-specific terminology is material.
- `semantic_intent`: optional relation declaration: `decomposition`, `sequence`, `decision`, `hierarchy`, `cycle`, `architecture`, `responsibility`, or `composite`.
- `style_family`: optional name from top-level `style_families`.
- `style`: optional slide-specific override applied after the family style.

## Layouts

### `cover`

Fields: `title`, `subtitle`, `kicker`, optional `image`.

### `section`

Fields: `title`, optional `subtitle`, optional `section_number`.

### `title_body`

Fields: `title`, `bullets`, optional `takeaway`, optional `image`.

### `two_column` and `comparison`

Fields: `title`, `left`, `right`. Each column is an object containing `heading`, `bullets`, and optional `accent`.

### `cards`

Fields: `title`, `cards` containing two to six objects with `heading`, `body`, optional `value`, and optional `accent`.

### `decomposition`

Fields: `title`, `items` containing two to six parallel objects with `heading`, optional `body`, optional `bullets`, optional `outcome`, and optional `accent`.

Use this for research tasks, work packages, capability groups, or other unordered partitions. It intentionally renders no step numbers or arrows. Set `semantic_intent: decomposition` when the distinction from a process is important.

### `process`

Fields: `title`, `steps` containing two to six objects with `heading` and optional `body`.

Use `process` only for a genuine linear sequence where order changes meaning. Use `decomposition` for parallel items and `diagram` when the source includes branching, hierarchy, cycles, responsibilities, or component relationships.

### `diagram`

Fields: `title`, `takeaway`, and `diagram`:

```json
{
  "kind": "decision",
  "direction": "left_to_right",
  "render": "native",
  "nodes": [
    {"id": "input", "label": "输入", "role": "start"},
    {"id": "check", "label": "满足条件？", "role": "decision"},
    {"id": "yes", "label": "主路径", "role": "output"},
    {"id": "no", "label": "回退路径", "role": "output"}
  ],
  "edges": [
    {"from": "input", "to": "check", "type": "flow"},
    {"from": "check", "to": "yes", "label": "是", "type": "output"},
    {"from": "check", "to": "no", "label": "否", "type": "output"}
  ]
}
```

Supported kinds: `flow`, `decision`, `hierarchy`, `cycle`, `swimlane`, and `architecture`.

Supported directions: `left_to_right` and `top_to_bottom`.

Render modes:

- `native`: native PowerPoint shapes and connectors.
- `auto`: currently uses native objects and leaves the renderer free to choose future compatible backends.
- `smartart`: native-shape fallback; `--powerpoint-smartart` converts supported `flow`, `hierarchy`, and `cycle` diagrams to real SmartArt when desktop PowerPoint is available.

Node roles: `start`, `end`, `process`, `decision`, `input`, `output`, `system`, `group`, `constraint`, and `default`. Nodes may also include `body`, `level`, `order`, `lane`, and `accent`.

Edge types: `flow`, `data`, `control`, `constraint`, `feedback`, `output`, and `reference`. They control connector color/dash treatment and preserve the intended meaning for review. Edge `color` remains available as an explicit override.

For `swimlane`, provide two to five `lanes`, for example `[{"id": "compiler", "label": "编译器"}]`, and set every node's `lane`.

### `composite_diagram`

Fields: `title`, `takeaway`, and `panels` containing two or three objects. Each panel requires `heading` and `diagram`; it may also include `width_ratio`, `accent`, and a panel-level `takeaway`.

Use this when one slide must show two coordinated but internally different systems, for example “architecture-aware mapping” beside “validation and tuning loop”. Keep each panel to one clear relation type and connect systems conceptually through the slide takeaway instead of forcing every node into one unreadable graph.

### `timeline`

Fields: `title`, `events` containing two to six objects with `label`, `heading`, and optional `body`.

### `chart`

Fields: `title`, optional `takeaway`, and `chart`:

```json
{
  "type": "column",
  "categories": ["A", "B", "C"],
  "series": [
    {"name": "2025", "values": [10, 15, 21]},
    {"name": "2026", "values": [13, 19, 28]}
  ]
}
```

Supported chart types: `column`, `bar`, `line`, `pie`, and `doughnut`. Values remain editable in PowerPoint.

### `table`

Fields: `title`, `columns`, `rows`, optional `takeaway`.

### `image_text`

Fields: `title`, `bullets`, `image`, and optional `image_side: left|right`. Text remains native; the illustration is a replaceable image object.

### `closing`

Fields: `title`, optional `subtitle`, optional `contact`.

## Style overrides

The style object may override:

```json
{
  "colors": {
    "background": "F7F8FA",
    "surface": "FFFFFF",
    "text": "17212B",
    "muted_text": "667085",
    "primary": "173F5F",
    "accent": "00A6A6",
    "accent_2": "F6AE2D",
    "line": "D9E2EC"
  },
  "fonts": {"heading": "Aptos Display", "body": "Aptos"},
  "sizes": {"title": 28, "body": 15, "small": 11},
  "geometry": {"margin_x": 0.7, "margin_top": 0.45, "gutter": 0.28}
}
```

Colors use six-digit RGB hex without `#`. Dimensions are inches and font sizes are points.

For multiple page families:

```json
{
  "style_families": {
    "research_task": {"colors": {"primary": "173F5F", "accent": "B21F2D"}},
    "technical_architecture": {"colors": {"background": "F4F6F8", "line": "8D99A6"}}
  },
  "slides": [
    {"type": "decomposition", "style_family": "research_task", "title": "…", "items": []},
    {"type": "composite_diagram", "style_family": "technical_architecture", "title": "…", "panels": []}
  ]
}
```

The merge order is default/extracted profile -> top-level `style` -> selected `style_family` -> slide `style`.

## Narrative checks

Before building, confirm:

- a two-stage spec points to the current approved content review and matching approval record;
- every content slide preserves the approved `fact_ids` and `must_show_terms`;
- every non-divider slide has a distinct takeaway;
- adjacent slides do not repeat the same argument;
- charts and tables use supplied data rather than invented values;
- terminology is consistent;
- the conclusion follows from the evidence;
- dense content has been split instead of reduced to unreadable type.
- body, diagram, and auxiliary text can remain at or above the configured font floors;
- emphasis is selective: key conclusions, mechanisms, metrics, and deliverables are bold, while full paragraphs are not;
- the slide stays within its density and primary-block budgets, or returns to content review for restructuring;
- unordered content is not mislabeled as a process;
- hardware/resource constraints point into mapping or scheduling;
- feedback edges identify what is tuned;
- speaker notes reference the same task/direction as the slide title.
