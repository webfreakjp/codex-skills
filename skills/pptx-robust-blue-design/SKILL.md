---
name: pptx-robust-blue-design
description: "Create new PowerPoint decks or restyle existing .pptx files in a robust blue presentation design derived from a PDF design template: deep blue section/cover slides, white content slides, bold blue headings, pale-blue cards, light gray rules, and process bars. Use when Codex needs to generate, edit, or redesign presentation slides with this specific clean guidebook-like style while avoiding any source-document organization labels or agency branding."
---

# PPTX Robust Blue Design

## Overview

Use this skill to create or restyle PowerPoint decks in the guidebook style captured in `references/design-system.md`. Preserve the user's content, remove source-template branding, and produce an editable `.pptx`.

## Workflow

1. Read `references/design-system.md` before making visual decisions.
2. Choose the operation:
   - Create a new deck from an outline or slide JSON.
   - Restyle an existing `.pptx` while preserving text.
   - Manually edit a `.pptx` using the same tokens and layout rules.
3. Pick slide layouts by message type before generating:
   - `message`: one central claim with 2-3 supporting points.
   - `compare`: contrasting responsibilities, options, before/after, or tradeoffs.
   - `process`: a short sequence where the order matters. Add a short explanation for every step.
   - `roadmap`: phased migration or time-based rollout.
   - `challenge-solution`: paired risks, issues, and mitigations.
   - `decision`: final choices, asks, or next actions.
4. Prefer `scripts/guidebook_pptx.py` for repeatable generation or restyling.
5. Validate the result by opening it with `python-pptx` and running `unzip -t`.
6. Check the deck for unwanted source-template labels before finalizing. Do not add organization names, agency labels, logos, or copied footer branding unless the user explicitly provides replacement branding.

## Script Usage

Install the dependency if needed:

```bash
python -m pip install --user python-pptx
```

Create a new deck from JSON:

```bash
python /path/to/skill/scripts/guidebook_pptx.py create \
  --slides slides.json \
  --output output.pptx
```

Restyle an existing deck:

```bash
python /path/to/skill/scripts/guidebook_pptx.py restyle \
  --input input.pptx \
  --output output.pptx
```

Create a small sample deck:

```bash
python /path/to/skill/scripts/guidebook_pptx.py sample --output sample.pptx
```

## Slide JSON

Use this structure for `create`:

```json
{
  "title": "Robust Blue Presentation",
  "subtitle": "Working draft",
  "slides": [
    {"type": "cover", "title": "Robust Blue Presentation", "subtitle": "Working draft"},
    {"type": "section", "title": "Requirements", "items": ["Define purpose", "Understand constraints"]},
    {"type": "content", "title": "Define the purpose", "body": "Clarify the decision this dashboard should support."},
    {"type": "process", "title": "Creation flow", "steps": ["Requirements", "Prototype", "Build"], "active": 0}
  ]
}
```

Supported slide types are `cover`, `section`, `content`, `message`, `two-column`, `compare`, `process`, `roadmap`, `challenge-solution`, and `decision`. If a type is missing, use `content`.

## Editing Rules

- Keep the deck 16:9.
- Use deep blue section slides sparingly for major dividers.
- Use white content slides for most information.
- Keep title text large, blue, and left aligned on content slides.
- Use pale-blue cards only for callouts, definitions, selected states, or grouped points that are introduced by the slide title or body. Do not drop unexplained cards onto a slide.
- Use light gray lines and boxes for structure.
- Avoid progress/status bars by default. Use them only when they add information that the main slide body does not already show.
- For any `process` or step-based slide, explain every step with a short phrase. Do not show step labels only.
- Keep parallel items visually equal by default. If one card, row, column, or step is colored differently, explain why it is selected, current, recommended, or risky.
- Do not reuse the same layout for every slide. Match the layout to the slide's communication job.
- Avoid copied source branding. The style is reusable, but the original organization's visible label is not.
