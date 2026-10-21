# Two-stage content approval workflow

Use this workflow for every new deck unless the user explicitly asks to skip content review and generate directly.

## Stage 1: content review only

Stage 1 ends with a user decision. It must not create a PPTX or presentation assets.

Create these files in the task's process directory:

1. `source_digest.json` — traceable source facts and terminology.
2. `storyline.json` — audience, purpose, slide questions, conclusions, and evidence IDs.
3. `content_review.json` — exact proposed visible copy and source mapping for every slide.
4. `content_review_qa.json` — machine-readable review results.
5. `slide_content_review.md` — human-readable content proof shown to the user.

The content review is not a short outline. It must let the user judge the finished message before layout exists. Each slide needs:

- a page role and the question it answers;
- the proposed title and takeaway;
- all visible headings, body copy, bullets, labels, callouts, metrics, and deliverables;
- the intended visual relationship, such as decomposition, sequence, architecture, or feedback cycle;
- `fact_ids` and source references;
- `must_show_terms` that protect source-specific terminology from later compression;
- a statement of what the audience should be able to explain after seeing the slide.

Use `content_blocks` in reading order. Each block must contain a heading and meaningful body text. For diagrams, include node and edge-label copy as blocks or in `diagram_copy`; the review must expose every important label that will appear in PowerPoint.

Use an evidence coverage table with one row per source fact:

```json
{
  "fact_id": "F4",
  "priority": "critical",
  "slides": [2],
  "visible": true,
  "rendered_as": "任务切分、PE阵列映射、数据流编排与配置生成",
  "source_location": "研究内容三（2）"
}
```

Priorities are `critical`, `high`, and `supporting`. All critical facts must be visible. At least 80% of high-priority facts must be visible. A fact present only in speaker notes does not count as visible.

Validate and render the review with:

```bash
python scripts/validate_content_review.py content_review.json \
  --source-digest source_digest.json \
  --markdown-out slide_content_review.md \
  --json-out content_review_qa.json \
  --strict
```

Then show the per-slide proof in the response and ask whether to continue. Stop. Do not inspect style merely to fill time, generate images, create `deck_spec.json`, or invoke the PPT builder while waiting.

If the user requests changes, revise the content review, rerun QA, show the new proof, and stop again. Never interpret requested revisions as approval to build.

## Stage 2: approved content to PPTX

Stage 2 begins only after the user explicitly approves the current content review. Short confirmations such as “继续”, “通过”, or “生成PPT” are sufficient when they clearly refer to the presented review.

Record the approval before creating the deck spec:

```bash
python scripts/record_content_approval.py content_review.json \
  --instruction "继续生成PPT" \
  --out content_approval.json
```

Add this object to `deck_spec.json`:

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

The paths are relative to `deck_spec.json`. Validation checks that the approval status is `approved` and that its SHA-256 digest matches the current content review. Editing the review after approval therefore invalidates the build until the user approves the revised content.

Carry each review slide's `fact_ids` and `must_show_terms` into the corresponding deck-spec slide. The content QA checks that required terms remain in visible slide text.

Layout may change line breaks, punctuation, and concise connective wording. It must not remove a required term, replace a precise mechanism with a generic phrase, change a number, add an unsupported claim, or move a critical fact to notes only. If content must change materially to fit, return to Stage 1.

## Stage 1 response pattern

Report:

- the number of proposed slides;
- the question and conclusion of each slide;
- the exact visible content in reading order;
- evidence coverage and warnings;
- paths to `slide_content_review.md`, `content_review.json`, and QA output;
- “尚未生成PPTX”。

End with one decision question: whether to continue generating the PPTX or revise the content.
