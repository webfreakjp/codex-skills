# Robust Blue PPTX Design System

## Design Tokens

- Canvas: 16:9 widescreen.
- Primary blue: `#0031D8`.
- Secondary blue: `#264AF4`.
- Pale blue: `#E8F1FE`.
- Text: `#1A1A1C`.
- Secondary text: `#626264`.
- Rule gray: `#E6E6E6`.
- Mid gray: `#949494`.
- Background: `#FFFFFF`.
- Font: use a modern sans-serif. Prefer `Yu Gothic` for Japanese text and `Aptos` or `Arial` for Latin text.

## Layout Patterns

### Cover

- Full-slide primary blue background.
- Large white title placed left of center.
- Optional subtitle below the title.
- Optional small outlined date/status badge below subtitle.
- Do not add logos, agency names, or source-template labels.

### Section Divider

- Full-slide primary blue background.
- Large white section title near the horizontal center.
- Thin white horizontal rule extending to the right of the title.
- Optional short item list or subtitle to the right of the rule.
- Do not add logos, agency names, or source-template labels.

### Content

- White background.
- Blue bold title in the upper-left.
- Thin gray rule below title with a short blue accent segment.
- Body text in dark gray or black.
- Right-side pale-blue rounded card for summaries, definitions, or supporting details.
- Use generous whitespace. Do not fill every corner.

### Process

- White background with the same heading treatment as content slides.
- Use three to five horizontal step blocks.
- Add one short explanatory phrase to every step block.
- Active/current step uses secondary blue with white text.
- Inactive steps use light gray or pale blue with dark text.
- Avoid footer progress bars by default. They should appear only when the main slide is not already showing the same stages.

### Message-Specific Layouts

- Use `message` for a strong claim supported by two or three proof points.
- Use `compare` when the audience must understand a contrast.
- Use `roadmap` when the slide is about phases, maturity, or rollout.
- Use `challenge-solution` when risks and mitigations must be paired row by row.
- Use `decision` when the slide should end with asks, actions, or choices.

## Composition Rules

- Keep rectangles at small corner radius only.
- Use cards for actual grouped content, not page sections.
- Prefer thin rules, simple boxes, and restrained callout cards.
- Choose layout from the message first, then place content. Avoid defaulting every slide to body text plus a side card.
- Explain visual devices before reusing them. Cards need a clear grouping label, and status bars need an introduced stage model.
- Do not duplicate a process as both main step cards and a footer progress bar unless the footer is explicitly useful for orientation across a later section.
- Treat side-by-side cards, columns, rows, and steps as equal unless the copy explicitly states that one is selected, current, recommended, or high risk.
- Keep any footer text generic or user-provided.
- Do not include source-document organization names, copied logos, or visible attribution from the PDF template.
