# Preserving a reference deck style

Read this before deriving style from a PPTX, PDF, or slide image.

## Editable PPTX reference

Use both sources of evidence:

1. Render and inspect representative slides to understand the visible system.
2. Group representative slides by page family, such as cover/section, ordinary content, research-task, and technical-architecture pages.
3. Extract theme colors, fonts, slide size, common fills, and typography with `scripts/extract_reference_style.py`.

Do not rely on whole-deck color frequency for a specialized technical page. Theme accents can reflect a generic Office theme while the visible page family uses navy/red, dark panels, or another local system. Extract local evidence with a one-based selection:

```bash
python scripts/extract_reference_style.py reference.pptx --slides 10-13,16 --out technical-style.json
```

With `--slides`, observed colors on those pages take priority over theme accents. Check `slide_numbers`, `local_palette_candidates`, and the rendered pages before assigning semantic roles such as `primary`, `accent`, and `accent_2`.

Pass the original PPTX to `build_editable_ppt.py --reference-pptx`. The builder opens that presentation, removes its source slides by default, and adds new native-object slides while retaining its theme, masters, and layouts. Write to a new output path.

The style profile controls newly created objects. If a deck contains multiple page families, put their overrides in top-level `style_families` and assign `style_family` per slide. A slide-level `style` override is the final, narrowest layer. If the extracted profile selects a decorative color as `primary`, correct the JSON rather than forcing the content to match a bad inference.

Match more than colors: preserve title position, rule treatment, margins, whitespace density, border radius, line weight, and whether technical pages use outlined zones or filled cards. Reusing the original master alone does not reproduce those visible conventions.

After building, compare like with like:

```bash
python scripts/qa_style_fidelity.py generated.pptx reference.pptx \
  --generated-slides 2-3 --reference-slides 10-13 --strict
```

This checks local palette coverage, heading/body fonts, median heading size, and background treatment. It is a gate for obvious drift, not a replacement for rendered visual inspection.

## PDF or image reference

Render or inspect representative pages. Reconstruct these rules in a style profile:

- canvas ratio and background treatment;
- title, subtitle, and body hierarchy;
- grid, margins, alignment, and whitespace density;
- palette roles rather than a raw color list;
- card geometry, border weight, radius impression, and shadow use;
- icon and illustration language;
- chart palette, axis, legend, and label treatment;
- image cropping, masks, overlays, and captions;
- recurring footer, logo, and page-number placement.

Do not use the reference content unless requested. Style reuse and content reuse are separate decisions.

## Fidelity tiers

- **Template fidelity:** original PPTX master/theme plus reconstructed native layouts. Best option.
- **System fidelity:** extracted palette, typography, grid, and page archetypes without the original template.
- **Visual approximation:** reconstruction from screenshots or PDF pages. Report missing fonts and effects.

Animations, transitions, SmartArt internals, proprietary fonts, embedded media, and third-party plugin effects are outside automatic parity. Prefer a stable editable substitute over a flattened screenshot.
