---
name: codex-ppt-simple
description: "Create native editable PowerPoint decks from source content while preserving the visual system of a reference PPTX, PDF, or slide image. Uses a two-stage workflow: first produce per-slide content for user approval, then build the PPTX. Use image-to-editable-ppt instead when reconstructing an existing image-only deck."
---

# Codex PPT Simple

Create `.pptx` files whose text, shapes, tables, diagrams, and charts remain editable. Default to a hybrid deck: native PowerPoint objects for information and optional generated images only for decorative or illustrative assets.

For non-trivial reports, papers, or technical documents, optimize for content quality before layout: extract traceable facts, design a conclusion-first storyline, and use semantic diagrams only when the source contains real sequence, condition, hierarchy, responsibility, cycle, or system relationships.

## Mandatory two-stage approval gate

For every new deck, use the two-stage workflow in [references/two-stage-workflow.md](references/two-stage-workflow.md). Do not collapse the stages merely because the initial request says “生成PPT”. Only skip the pause when the same request explicitly says to skip content review or directly generate without confirmation.

### Stage 1 — content review

1. Read and synthesize the source before designing slides. For substantial material, read [references/content-and-diagrams.md](references/content-and-diagrams.md).
2. Create `source_digest.json`, `storyline.json`, and `content_review.json`. The review must contain the exact proposed visible copy for every slide, not just titles or a short abstract. Map critical facts to visible slide text and retain source-specific objects, actions, mechanisms, constraints, metrics, and deliverables.
3. Run `scripts/validate_content_review.py`, produce `content_review_qa.json` and `slide_content_review.md`, and resolve errors. Treat missing critical facts, missing required terms, weak source mapping, and notes-only critical evidence as content failures.
4. Present the per-slide outline and copy to the user, report the review artifacts, ask whether to continue, and stop the turn. Do not create `deck_spec.json`, generate images, build a PPTX, or run slide-layout/style QA before explicit approval.

If the user requests revisions, update the Stage 1 artifacts, rerun content-review QA, present the revised content, and stop again. A revision invalidates any earlier approval.

### Stage 2 — PPT generation

Start only after an explicit user message such as “继续生成”, “内容通过”, or “按当前内容生成PPT”. Load the approved `content_review.json`; do not silently weaken, generalize, or materially rewrite its visible content to make layout easier.

1. Record the approval with `scripts/record_content_approval.py`. The record locks the approved review by SHA-256 digest.
2. Inspect the reference style and create `deck_spec.json`. Add a top-level `workflow` object that points to the approved content review and approval record. Preserve each slide's `fact_ids` and `must_show_terms` in the deck spec.
3. Validate, build, and run style/editability QA as described below. The builder rejects a two-stage spec when the approval record is missing or its content digest no longer matches.
4. If the approved copy cannot fit without a material content change, return to Stage 1 with a proposed revision instead of changing the message during rendering.

## Non-negotiable outcome

- Never flatten a complete slide into a screenshot or generated full-slide image.
- Keep titles, body text, data, tables, timelines, processes, and diagrams as native PowerPoint objects.
- Generated raster art must contain no presentation text or data. It stays replaceable and removable.
- Prefer SVG or native shapes for icons and simple diagrams.
- Do not claim that pixels inside a photo or generated illustration are editable.
- Treat readability as a hard constraint. For ordinary 16:9 decks, keep titles at least 26 pt, takeaways 17 pt, module headings 15 pt, body text 13.5 pt, diagram text 12 pt, and auxiliary labels 11 pt unless the user explicitly chooses a different viewing context.
- Never solve overflow by silently shrinking below those floors. Compress redundant syntax, restructure, split the page, or return to Stage 1 when visible copy must change.
- Create hierarchy below the title: selectively bold conclusions, objects, mechanisms, metrics, deliverables, innovations, and decisive constraints. Do not bold whole paragraphs or give every block equal visual weight.

## Stage 2 production workflow

1. If the reference is PPTX, read [references/reference-style.md](references/reference-style.md), inspect rendered representative slides when rendering is available, group comparable pages into visual families, and run `scripts/extract_reference_style.py --slides ...` for each relevant family. Do not let a whole-deck Office theme override the visible palette of specialized research or technical pages. Reuse the master/theme by passing the reference to the builder when high fidelity matters.
2. If the reference is PDF or images, inspect representative rendered pages and reconstruct a style profile. Do not infer visible style from metadata alone.
3. Create `deck_spec.json` following [references/deck-spec.md](references/deck-spec.md). Translate the approved content into layouts without changing its claims, terminology, metrics, or deliverables. Add `readability_policy`, `emphasis_policy`, and `density_policy`, or consciously inherit their defaults. Use `**...**` markers for one to three important spans per block; they render as native bold runs and are not visible as markup.
4. Classify the relationship before selecting a visual. Use `decomposition` for parallel research tasks or capability groups; use `process` only for genuine order; use a semantic `diagram` for conditions, hierarchy, cycles, responsibilities, or system relationships; use `composite_diagram` for two or three coordinated subsystems with different internal logic. Add typed edges (`data`, `constraint`, `feedback`, and so on). Do not invent edges that the source does not support. Prefer native shapes for decision flows, swimlanes, architecture, and composite diagrams; request SmartArt only for simple process, hierarchy, or cycle diagrams.
5. Run `scripts/validate_deck_spec.py` and `scripts/qa_content.py` before building. Resolve errors. Review warnings for missing approved terms or facts, generic titles, repetition, duplicated structured content, false sequence, conflicting task numbers in notes, overlong bullets, excessive visible-character density, too many primary blocks, missing evidence, disconnected nodes, and weak decision or feedback branches.
6. In hybrid mode, generate only the illustration assets that the approved spec actually needs. Ask the image backend for no text and, when useful, transparent backgrounds. Save paths in the spec.
7. Build with `scripts/build_editable_ppt.py`. Use an output path different from the reference PPTX. When desktop PowerPoint is available and the spec requests `render: smartart`, add `--powerpoint-smartart`; otherwise the builder keeps an editable native-shape fallback.
8. When a reference exists, run `scripts/qa_style_fidelity.py` against comparable generated and reference page families. Then run `scripts/qa_editability.py` and `scripts/qa_readability.py --spec deck_spec.json --strict`. Inspect medium/high editing-complexity warnings and a rendered preview when PowerPoint, LibreOffice, or another renderer is available. Correct undersized text, missing emphasis, overflow, overlap, weak hierarchy, diagram crossings, competing focal points, or page-family style drift in the spec and rebuild only the changed deck.

## Commands

Use a Python interpreter that can import `python-pptx` and Pillow. If dependencies are missing, install `requirements.txt` only with the user's permission when network or environment changes require it.

```bash
python scripts/validate_content_review.py content_review.json --source-digest source_digest.json --markdown-out slide_content_review.md --json-out content_review_qa.json --strict
python scripts/record_content_approval.py content_review.json --instruction "用户明确同意继续生成" --out content_approval.json
python scripts/extract_reference_style.py reference.pptx --slides 10-13,16 --out technical_style.json
python scripts/validate_deck_spec.py deck_spec.json --check-assets
python scripts/qa_content.py deck_spec.json --json-out content_qa.json
python scripts/build_editable_ppt.py deck_spec.json --style-profile technical_style.json --reference-pptx reference.pptx --powerpoint-smartart --out final.pptx
python scripts/qa_style_fidelity.py final.pptx reference.pptx --generated-slides 2-3 --reference-slides 10-13 --json-out style_qa.json
python scripts/qa_editability.py final.pptx --json-out qa_report.json
python scripts/qa_readability.py final.pptx --spec deck_spec.json --json-out readability_qa.json --strict
```

When no reference exists, omit `--reference-pptx` and `--style-profile`; the builder uses a restrained editable theme. For decks with multiple visual families, add `style_families` to the spec and assign `style_family` per slide. Start with `assets/deck_spec.example.json`; use `assets/deck_spec.research.example.json` for parallel research tasks plus a composite technical architecture.

## Reference fidelity

- Editable reference PPTX: preserve its slide size, masters, layouts, and theme; remove source slides unless the user asks to keep them.
- Image-only reference: reproduce palette, typography hierarchy, grid, density, shapes, chart treatment, and image treatment as editable rules. Exact recovery of original objects is not possible.
- Missing fonts, proprietary effects, SmartArt, transitions, and animations may need explicit substitution or manual finishing. Report substitutions.

## Completion report

At the end of Stage 1, report the content-review paths, slide count, evidence coverage, unresolved warnings, and the full per-slide review copy; explicitly state that no PPTX has been generated yet and ask whether to continue.

At the end of Stage 2, report the PPTX path, slide count, mode, reference used, approved-content digest, style profile path, content-QA score, editability QA result, readability-QA result including minimum fonts and emphasis, semantic diagram count, SmartArt conversions versus native fallbacks, and which remaining assets are raster images. Do not describe a deck as fully editable when important information is baked into images.
