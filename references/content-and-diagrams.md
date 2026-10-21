# Content synthesis and semantic diagrams

Read this for reports, papers, technical documents, and any deck that needs strong summarization or process diagrams.

## Content pipeline

Do not go directly from source text to slide copy for a substantial document. Use four layers:

1. `source_digest.json`: traceable facts, metrics, terminology, actors, sequence, conditions, comparisons, and source locations.
2. `storyline.json`: audience, purpose, core conclusion, sections, slide questions, and the evidence supporting each proposed conclusion.
3. `content_review.json`: exact proposed visible copy for every slide, fact-to-slide coverage, required terminology, visual relationship, and source references. Validate it and obtain explicit user approval.
4. `deck_spec.json`: the approved copy translated into editable layouts, semantic diagrams, detailed speaker notes, and source references.

Do not create `deck_spec.json` before Stage 1 approval. Use conclusion-first titles for working and technical decks. A content slide should normally contain one takeaway and no more than the configured bullet limit, but do not obtain brevity by deleting the mechanism, constraint, metric, or deliverable that makes a technical claim meaningful. Move secondary explanation and caveats into speaker notes; critical facts must remain visible.

Treat readability as a content constraint, not a rendering preference. Draft copy to fit the role-specific font floors in `readability_policy`. When a page is too dense, compress grammar into object-action-result phrasing, remove repetition, restructure, or split the slide. Do not keep all prose and ask the renderer to reduce the font.

## Content-review standard

Treat a slide as a set of technical assertions, not a set of layout slots. For technical and research material, each major content block should normally identify:

- the object or problem being addressed;
- the action or method;
- the mechanism, constraint, or decision criterion;
- the result, metric, deliverable, or technical value.

The slide does not need to repeat all four labels, but the audience should not have to invent the missing causal link. Avoid standalone generic claims such as “高效适配”, “协同优化”, “形成体系”, or “持续提升” unless the same block names how the result is achieved or measured.

Use a fact coverage matrix. Mark source facts as `critical`, `high`, or `supporting`; record whether they are visible and the exact copy that carries them. All critical facts and at least 80% of high-priority facts must be visible before user review. Preserve named systems, modules, algorithms, quantities, comparison dimensions, constraints, and assessment targets through `must_show_terms`.

For each slide, apply a slide-alone answerability test: after seeing only that page, can the intended audience explain the object, method, important constraint or difficulty, and expected result? If the page only names categories, it is still an outline and is not ready for approval.

## Emphasis hierarchy

Every content slide needs visible hierarchy below the title. Prefer one sentence-level takeaway, then three to five primary blocks, then supporting detail. Mark selected native-text runs with `**...**` in approved copy or the deck spec; the builder renders the enclosed text bold and removes the markers.

Prioritize conclusions, named objects, critical actions, mechanisms, metrics, deliverables, innovations, and decisive constraints. Use one to three emphasized spans per block and keep the overall emphasis selective. Do not bold full paragraphs or combine bold, accent color, and underline on the same phrase without a specific semantic reason.

For compact technical writing, prefer:

`**对象或动作：**机制、约束与结果。`

This preserves factual depth while making the page scannable. Emphasis is formatting, so adding or changing bold markers does not change the approved claim; changing words, facts, or scope still requires Stage 1 review.

## Content profile

Use `content_profile` to make summary behavior explicit:

```json
{
  "audience": "technical",
  "summary_level": "standard",
  "title_style": "conclusion_first",
  "max_bullets": 4,
  "max_nodes_per_slide": 7,
  "source_traceability": true,
  "notes_include_evidence": true
}
```

Summary levels:

- `executive`: strongest compression; value, decision, risk, and recommendation first.
- `standard`: balance conclusion, mechanism, and evidence.
- `technical`: preserve terminology, constraints, and implementation detail.

## Evidence and traceability

Slides may include:

```json
{
  "evidence": ["AtomicCASOp maps to linalg_ext"],
  "source_refs": ["Triton-Linalg.pdf p.3"]
}
```

The builder appends evidence and source references to speaker notes. Never invent a source reference. Preserve exact values and domain terms from the source digest.

In the approved `deck_spec.json`, also preserve:

```json
{
  "fact_ids": ["F2", "F4"],
  "must_show_terms": ["Evoformer", "PE阵列映射", "配置生成"]
}
```

`fact_ids` connect the rendered slide to the approved review. `must_show_terms` are visible-copy invariants, not speaker-note keywords.

## When to generate a diagram

Generate a semantic diagram only when the source expresses a relationship that is easier to understand spatially:

- unordered work packages, research contents, or capability groups -> `decomposition` slide;
- sequence or transformation -> `flow`;
- if/else, eligibility, or routing -> `decision`;
- parent/child or composition -> `hierarchy`;
- iteration or feedback -> `cycle`;
- responsibility across teams/modules -> `swimlane`;
- layers, components, or interfaces -> `architecture`.
- two or three coordinated subsystems with different internal logic -> `composite_diagram` slide.

Classify the relation before choosing the visual. Ask whether swapping two items would change the meaning. If not, it is probably decomposition rather than sequence. Do not turn an unordered list into a fake process, number parallel tasks, or connect them with arrows. When an edge is inferred rather than explicit, record the inference in notes or keep the content as a structured list.

For architecture pages, model constraints as inputs to mapping or scheduling rather than as a downstream output. For validation loops, name what feedback changes: platform parameters, operator implementation, compiler mapping, or scheduling strategy. Avoid a generic arrow back to the first box when the actual target is more specific.

## Diagram specification

```json
{
  "type": "diagram",
  "title": "结构化实现优先，缺失时回退到扩展方言",
  "takeaway": "转换策略先争取结构化优化机会，再提供兼容回退。",
  "diagram": {
    "kind": "decision",
    "direction": "left_to_right",
    "render": "native",
    "nodes": [
      {"id": "input", "label": "Triton Op", "role": "start"},
      {"id": "check", "label": "存在结构化实现？", "role": "decision"},
      {"id": "structured", "label": "Structured Op", "role": "output"},
      {"id": "fallback", "label": "Generic / 扩展方言", "role": "output"}
    ],
    "edges": [
      {"from": "input", "to": "check", "type": "flow"},
      {"from": "check", "to": "structured", "label": "是", "type": "output"},
      {"from": "check", "to": "fallback", "label": "否", "type": "output"}
    ]
  }
}
```

Node roles: `start`, `end`, `process`, `decision`, `input`, `output`, `system`, `group`, `constraint`, and `default`.

Edge types make meaning visible and editable: `flow`, `data`, `control`, `constraint`, `feedback`, `output`, and `reference`. Use `constraint` for hardware/resource limits and `feedback` only for an actual tuning or iteration loop.

For swimlanes, add `lanes` and a `lane` on every node. Use `level` to guide hierarchy or architecture placement, and `order` to control ordering within a swimlane.

## SmartArt policy

`render: smartart` is supported for simple `flow`, `hierarchy`, and `cycle` diagrams when desktop PowerPoint is available. The builder first creates native editable shapes. With `--powerpoint-smartart`, PowerPoint replaces them only after native SmartArt creation succeeds. On unsupported platforms or diagram types, the editable shape version remains.

Use native shapes for decisions, architecture, and swimlanes because they provide more predictable routing and semantic control.

## Quality gates

Run `qa_content.py` before building. Review:

- missing approved fact IDs or required visible terms;
- generic or topic-only titles;
- missing takeaways, evidence, or source references;
- repeated slide messages;
- repeated structured items across slides;
- false process diagrams used for decomposition;
- direction/task numbers in notes that conflict with the title;
- bullet count and length versus the selected summary level, including bullets nested inside cards, tasks, and decomposition items;
- isolated nodes and unknown edge targets;
- decision nodes with fewer than two outgoing branches;
- diagrams above the configured node limit.
- cycles without an explicit feedback edge and constraint nodes without constraint edges.

After building, run editability QA and inspect a rendered preview. Diagram labels, branch names, reading order, connector crossings, page focus, and pages with medium/high editing complexity require visual review.

Also run `qa_readability.py` against the actual PPTX and deck spec. Treat undersized text, excessive visible-character density, and missing non-title emphasis as build failures. An over-high bold ratio is a warning that too much content has equal visual weight. The final preview must still be checked for overflow and line wrapping because OOXML alone cannot guarantee PowerPoint's rendered fit.
