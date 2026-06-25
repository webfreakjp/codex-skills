#!/usr/bin/env python3
"""Create or restyle PPTX decks in a dashboard guidebook visual style."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError as exc:
    raise SystemExit("Missing dependency: install with `python -m pip install --user python-pptx`.") from exc


BLUE = RGBColor(0, 49, 216)
BLUE_2 = RGBColor(38, 74, 244)
PALE_BLUE = RGBColor(232, 241, 254)
TEXT = RGBColor(26, 26, 28)
GRAY = RGBColor(148, 148, 148)
LIGHT_GRAY = RGBColor(230, 230, 230)
MID_GRAY = RGBColor(98, 98, 100)
WHITE = RGBColor(255, 255, 255)
CARD_BORDER = RGBColor(190, 205, 235)
FONT_JP = "Yu Gothic"


def text_of_slide(slide) -> list[str]:
    texts: list[str] = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text.strip()
            if text:
                texts.append(text)
    return texts


def clear_slides(prs: Presentation) -> None:
    while len(prs.slides):
        r_id = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[0]


def set_run(run, size: float, color: RGBColor, bold: bool = False) -> None:
    run.font.name = FONT_JP
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold


def textbox(slide, left, top, width, height, text: str, size: float, color: RGBColor,
            bold: bool = False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    frame.vertical_anchor = anchor
    para = frame.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = str(text or "")
    set_run(run, size, color, bold)
    return box


def rect(slide, left, top, width, height, fill: RGBColor, line: RGBColor | None = None,
         rounded: bool = False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.8)
    return shape


def add_heading(slide, title: str) -> None:
    textbox(slide, Inches(0.55), Inches(0.52), Inches(7.2), Inches(0.5), title, 23, BLUE, True)
    rect(slide, Inches(0.55), Inches(1.17), Inches(4.2), Inches(0.012), LIGHT_GRAY)
    rect(slide, Inches(0.55), Inches(1.17), Inches(1.0), Inches(0.025), BLUE)


def add_footer(slide, page_no: int, steps: list[str] | None = None, active: int | None = None) -> None:
    labels = steps or ["Requirements", "Prototype", "Build"]
    labels = labels[:5]
    if not labels:
        return
    left = Inches(0.55)
    top = Inches(6.25)
    gap = Inches(0.05)
    width = Inches(11.5 / len(labels)) - gap
    for idx, label in enumerate(labels):
        color = BLUE_2 if active == idx else LIGHT_GRAY
        x = left + idx * (width + gap)
        rect(slide, x, top, width, Inches(0.16), color)
        textbox(slide, x, top - Inches(0.18), width, Inches(0.12), label, 5.5,
                BLUE if active == idx else GRAY, active == idx, PP_ALIGN.CENTER)
    textbox(slide, Inches(11.35), Inches(6.88), Inches(1.05), Inches(0.12),
            str(page_no), 6, GRAY, align=PP_ALIGN.RIGHT)


def cover_slide(prs: Presentation, data: dict[str, Any], page_no: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, prs.slide_width, prs.slide_height, BLUE)
    textbox(slide, Inches(0.72), Inches(1.82), Inches(8.2), Inches(1.15),
            data.get("title", "Untitled"), 34, WHITE, True)
    if data.get("subtitle"):
        textbox(slide, Inches(0.74), Inches(3.28), Inches(6.6), Inches(0.38),
                data["subtitle"], 14, WHITE)
    if data.get("badge"):
        badge = rect(slide, Inches(0.74), Inches(4.08), Inches(2.25), Inches(0.34), BLUE, WHITE)
        badge.fill.transparency = 20
        textbox(slide, Inches(0.86), Inches(4.17), Inches(2.0), Inches(0.12),
                data["badge"], 6.5, WHITE)
    rect(slide, Inches(5.45), Inches(3.06), Inches(5.95), Inches(0.012), WHITE)
    textbox(slide, Inches(11.35), Inches(6.82), Inches(1.05), Inches(0.12),
            str(page_no), 6, WHITE, align=PP_ALIGN.RIGHT)


def section_slide(prs: Presentation, data: dict[str, Any], page_no: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, prs.slide_width, prs.slide_height, BLUE)
    textbox(slide, Inches(2.2), Inches(3.05), Inches(4.55), Inches(0.62),
            data.get("title", "Section"), 31, WHITE, True, PP_ALIGN.RIGHT)
    rect(slide, Inches(7.0), Inches(3.42), Inches(2.55), Inches(0.012), WHITE)
    items = data.get("items") or []
    right_text = data.get("subtitle") or "\n".join(str(item) for item in items[:4])
    if right_text:
        textbox(slide, Inches(9.85), Inches(3.04), Inches(2.35), Inches(0.8),
                right_text, 12.5, WHITE, True)
    textbox(slide, Inches(11.35), Inches(6.82), Inches(1.05), Inches(0.12),
            str(page_no), 6, WHITE, align=PP_ALIGN.RIGHT)


def content_slide(prs: Presentation, data: dict[str, Any], page_no: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = data.get("title", "Untitled")
    body = data.get("body", "")
    add_heading(slide, title)
    textbox(slide, Inches(0.62), Inches(1.55), Inches(5.25), Inches(3.35), body, 13.5, TEXT)
    rect(slide, Inches(6.28), Inches(1.42), Inches(5.55), Inches(2.75), PALE_BLUE, CARD_BORDER, True)
    textbox(slide, Inches(6.58), Inches(1.78), Inches(1.0), Inches(0.25),
            f"{page_no:02d}", 14, BLUE, True)
    textbox(slide, Inches(7.08), Inches(1.76), Inches(4.25), Inches(0.32), title, 15, TEXT, True)
    rect(slide, Inches(6.58), Inches(2.34), Inches(4.9), Inches(0.01), CARD_BORDER)
    card_text = data.get("callout") or body
    textbox(slide, Inches(6.58), Inches(2.70), Inches(4.8), Inches(0.95), card_text, 12, TEXT)
    add_footer(slide, page_no, data.get("steps"), data.get("active"))


def two_column_slide(prs: Presentation, data: dict[str, Any], page_no: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_heading(slide, data.get("title", "Untitled"))
    cols = data.get("columns") or []
    while len(cols) < 2:
        cols.append({})
    for idx, col in enumerate(cols[:2]):
        x = Inches(0.72 + idx * 5.85)
        fill = PALE_BLUE if idx == 0 else WHITE
        rect(slide, x, Inches(1.65), Inches(5.05), Inches(3.5), fill, CARD_BORDER, True)
        textbox(slide, x + Inches(0.32), Inches(1.95), Inches(4.4), Inches(0.3),
                col.get("title", ""), 15, BLUE if idx == 0 else TEXT, True)
        textbox(slide, x + Inches(0.32), Inches(2.45), Inches(4.35), Inches(2.05),
                col.get("body", ""), 12.5, TEXT)
    add_footer(slide, page_no, data.get("steps"), data.get("active"))


def process_slide(prs: Presentation, data: dict[str, Any], page_no: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_heading(slide, data.get("title", "Process"))
    steps = [str(step) for step in data.get("steps", ["Requirements", "Prototype", "Build"])]
    active = int(data.get("active", 0))
    count = max(1, len(steps))
    width = Inches(11.3 / count)
    for idx, step in enumerate(steps):
        x = Inches(0.72) + idx * width
        fill = BLUE_2 if idx == active else (PALE_BLUE if idx > active else LIGHT_GRAY)
        text_color = WHITE if idx == active else TEXT
        rect(slide, x, Inches(2.85), width - Inches(0.08), Inches(0.82), fill, CARD_BORDER, True)
        textbox(slide, x + Inches(0.12), Inches(3.12), width - Inches(0.32), Inches(0.2),
                step, 12, text_color, True, PP_ALIGN.CENTER)
    if data.get("body"):
        textbox(slide, Inches(0.74), Inches(4.25), Inches(10.8), Inches(0.7), data["body"], 12.5, TEXT)
    add_footer(slide, page_no, steps, active)


def build_deck(spec: dict[str, Any], output: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    clear_slides(prs)
    slides = spec.get("slides") or [
        {"type": "cover", "title": spec.get("title", "Untitled"), "subtitle": spec.get("subtitle", "")}
    ]
    builders = {
        "cover": cover_slide,
        "section": section_slide,
        "content": content_slide,
        "two-column": two_column_slide,
        "process": process_slide,
    }
    for page_no, slide_spec in enumerate(slides, 1):
        builder = builders.get(slide_spec.get("type", "content"), content_slide)
        builder(prs, slide_spec, page_no)
    prs.save(output)


def restyle(input_path: Path, output: Path) -> None:
    source = Presentation(input_path)
    extracted = [text_of_slide(slide) for slide in source.slides]
    slides: list[dict[str, Any]] = []
    for idx, texts in enumerate(extracted):
        title = texts[0] if texts else f"Slide {idx + 1}"
        body = "\n".join(texts[1:])
        if idx == 0:
            slides.append({"type": "cover", "title": title, "subtitle": body})
        elif idx == len(extracted) - 1 and len(extracted) > 2:
            slides.append({"type": "section", "title": title, "subtitle": body})
        else:
            slides.append({"type": "content", "title": title, "body": body})
    build_deck({"slides": slides}, output)


def sample(output: Path) -> None:
    build_deck(
        {
            "slides": [
                {
                    "type": "cover",
                    "title": "Robust Blue Presentation",
                    "subtitle": "Editable PowerPoint template",
                    "badge": "Draft",
                },
                {
                    "type": "section",
                    "title": "Requirements",
                    "items": ["Define purpose", "Understand constraints", "Prepare worksheet"],
                },
                {
                    "type": "content",
                    "title": "Define the purpose",
                    "body": "Clarify the user, decision, and action that the dashboard should support.",
                    "callout": "A good dashboard starts from the decision it enables.",
                    "active": 0,
                },
                {
                    "type": "process",
                    "title": "Creation flow",
                    "steps": ["Requirements", "Prototype", "Build"],
                    "active": 1,
                    "body": "Move from purpose and constraints into layout exploration before implementation.",
                },
            ]
        },
        output,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    create = sub.add_parser("create")
    create.add_argument("--slides", required=True, type=Path, help="JSON slide specification.")
    create.add_argument("--output", required=True, type=Path)

    restyle_cmd = sub.add_parser("restyle")
    restyle_cmd.add_argument("--input", required=True, type=Path)
    restyle_cmd.add_argument("--output", required=True, type=Path)

    sample_cmd = sub.add_parser("sample")
    sample_cmd.add_argument("--output", required=True, type=Path)

    args = parser.parse_args()
    if args.command == "create":
        build_deck(json.loads(args.slides.read_text(encoding="utf-8")), args.output)
    elif args.command == "restyle":
        restyle(args.input, args.output)
    elif args.command == "sample":
        sample(args.output)


if __name__ == "__main__":
    main()
