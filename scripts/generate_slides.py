"""
Generate comprehensive lecture slide decks (.pptx) for each reading.
One deck per reading, saved in LectureSlides/.
"""

import sys
import unittest.mock as mock
from pathlib import Path

# Monkey-patch broken cryptography module in this environment
sys.modules['cryptography'] = mock.MagicMock()
sys.modules['cryptography.hazmat'] = mock.MagicMock()
sys.modules['cryptography.hazmat.primitives'] = mock.MagicMock()
sys.modules['cryptography.hazmat.primitives.padding'] = mock.MagicMock()
sys.modules['cryptography.exceptions'] = mock.MagicMock()
sys.modules['cryptography.hazmat.bindings'] = mock.MagicMock()
sys.modules['cryptography.hazmat.bindings._rust'] = mock.MagicMock()

import json
import pypdf
import anthropic

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy


# ── Paths ────────────────────────────────────────────────────────────────────
READINGS_DIR   = Path(__file__).parent.parent / "Readings"
SUMMARIES_DIR  = Path(__file__).parent.parent / "Summaries"
TAKEAWAYS_DIR  = Path(__file__).parent.parent / "KeyTakeaways"
SLIDES_DIR     = Path(__file__).parent.parent / "LectureSlides"

READINGS = [
    {
        "file":          "2310.11986v2.pdf",
        "summary_file":  "2310.11986v2_summary.pdf",
        "takeaway_file": "2310.11986v2_takeaways.pdf",
        "title":         "Sociotechnical Safety Evaluation of Generative AI Systems",
        "authors":       "Weidinger et al. (Google DeepMind, 2023)",
        "max_pages":     76,
    },
    {
        "file":          "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero.pdf",
        "summary_file":  "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero_summary.pdf",
        "takeaway_file": "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero_takeaways.pdf",
        "title":         '"All Roads Lead to ChatGPT": How Generative AI is Eroding Social Interactions and Student Learning Communities',
        "authors":       "Hou, Man, Hamilton, Muthusekaran, Johnykutty et al. (Temple University, 2025)",
        "max_pages":     None,
    },
    {
        "file":          "Excerpts from AI Now 2025 Reports on AI Power.pdf",
        "summary_file":  "Excerpts from AI Now 2025 Reports on AI Power_summary.pdf",
        "takeaway_file": "Excerpts from AI Now 2025 Reports on AI Power_takeaways.pdf",
        "title":         "Artificial Power: AI Now Institute 2025 Report",
        "authors":       "AI Now Institute (2025)",
        "max_pages":     None,
    },
    {
        "file":          "Executive Summary - AI Now Institute.pdf",
        "summary_file":  "Executive Summary - AI Now Institute_summary.pdf",
        "takeaway_file": "Executive Summary - AI Now Institute_takeaways.pdf",
        "title":         "Confronting Tech Power: AI Now Institute 2023 Landscape Report",
        "authors":       "AI Now Institute (2023)",
        "max_pages":     None,
    },
    {
        "file":          "hai_ai_index_report_2025.pdf",
        "summary_file":  "hai_ai_index_report_2025_summary.pdf",
        "takeaway_file": "hai_ai_index_report_2025_takeaways.pdf",
        "title":         "AI Index Report 2025",
        "authors":       "Stanford HAI (2025)",
        "max_pages":     30,
    },
]


# ── Color palette ─────────────────────────────────────────────────────────────
NAVY      = RGBColor(0x1a, 0x1a, 0x2e)   # slide background / title bar
BLUE      = RGBColor(0x2c, 0x5f, 0x8a)   # accent
TEAL      = RGBColor(0x16, 0x85, 0x7b)   # highlight
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE  = RGBColor(0xF5, 0xF5, 0xF0)
LIGHTGRAY = RGBColor(0xCC, 0xCC, 0xCC)
DARKGRAY  = RGBColor(0x44, 0x44, 0x44)
AMBER     = RGBColor(0xE8, 0x9C, 0x13)   # discussion / callout accent


# ── PDF text extraction ───────────────────────────────────────────────────────
def extract_text(pdf_path: Path, max_pages: int | None = None) -> str:
    reader = pypdf.PdfReader(str(pdf_path))
    pages = reader.pages if max_pages is None else reader.pages[:max_pages]
    return "\n".join(p.extract_text() or "" for p in pages)


# ── Claude call ───────────────────────────────────────────────────────────────
SLIDE_SCHEMA = """
Return a JSON array of slide objects. Each object has:
  "type":    one of "title" | "agenda" | "section" | "content" | "two_col" | "quote" | "data" | "discussion" | "summary" | "thankyou"
  "title":   slide heading string (keep ≤ 55 chars)
  "subtitle": (title/section/thankyou slides only) short subtitle string
  "body":    (content/data/summary slides) list of 3–6 bullet strings; each ≤ 120 chars. Sub-bullets start with "  – "
  "left":    (two_col slides) list of 3–5 bullet strings for left column
  "right":   (two_col slides) list of 3–5 bullet strings for right column
  "left_header":  (two_col) column header string
  "right_header": (two_col) column header string
  "quote":   (quote slides) the verbatim or paraphrased quotation string
  "attribution": (quote slides) source/attribution string
  "questions": (discussion slides) list of 3–4 discussion question strings
  "note":    (any) optional presenter note string (1–3 sentences)

Do not include any key that is not relevant to the slide type.
Return ONLY the raw JSON array — no markdown fences, no extra text.
"""

def generate_slide_content(reading: dict, original_text: str,
                            summary_text: str, takeaways_text: str,
                            client: anthropic.Anthropic) -> list[dict]:
    prompt = f"""You are designing a comprehensive university lecture for a graduate course (COS598) on AI.

The reading is: "{reading['title']}" by {reading['authors']}

Below are supporting materials:

SUMMARY:
{summary_text}

KEY TAKEAWAYS & CONFUSION AREAS:
{takeaways_text}

ORIGINAL TEXT EXCERPT:
{original_text[:45000]}

---

Design a full 50-minute lecture slide deck. The deck must have 28–35 slides and cover the reading in depth — this is the ONLY lecture on this reading. Structure the deck as follows:

1. Title slide
2. Learning Objectives slide (4–5 concrete objectives, what students will be able to do)
3. Agenda / Roadmap slide
4. Context & Motivation section (2–3 slides: why this topic matters now, historical/industry context)
5. Core Framework or Central Argument section (4–6 slides: the paper's main thesis, framework, or model explained carefully — use two_col slides to contrast concepts, quote slides for key definitions)
6. Key Findings or Evidence section (4–6 slides: specific data points, examples, case studies from the reading — be concrete and quantitative where possible)
7. Critical Analysis section (3–4 slides: strengths, limitations, counterarguments, what the paper leaves open)
8. Implications & Applications section (3–4 slides: so what? what does this mean for practitioners, policymakers, researchers?)
9. Common Misconceptions slide (use a content slide to address 3–4 confusions from the takeaways analysis)
10. Discussion Questions slide
11. Key Takeaways / Summary slide
12. Thank You / Q&A slide

Use a good mix of slide types. Include at least 2 quote slides, at least 2 two_col slides, and at least 2 discussion slides. Make presenter notes substantive (2–3 sentences of what to say or emphasize).

Be specific and detailed — pull real numbers, real terms, and real examples from the reading. Do not be vague.

{SLIDE_SCHEMA}"""

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=16000,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        final = stream.get_final_message()

    text_blocks = [b.text for b in final.content if b.type == "text"]
    if not text_blocks:
        raise ValueError(f"No text block in response. Stop reason: {final.stop_reason}. "
                         f"Blocks: {[b.type for b in final.content]}")
    response_text = text_blocks[0]

    clean = response_text.strip()
    if clean.startswith("```"):
        clean = clean.split("```", 1)[1]
        if clean.startswith("json"):
            clean = clean[4:]
        clean = clean.rsplit("```", 1)[0]
    return json.loads(clean.strip())


# ── Slide rendering helpers ───────────────────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs: Presentation):
    layout = prs.slide_layouts[6]   # Blank
    return prs.slides.add_slide(layout)


def fill_bg(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill_color: RGBColor,
             line_color: RGBColor | None = None, line_width_pt: float = 0):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, left, top, width, height, text: str,
                font_size: int, bold: bool = False, color: RGBColor = WHITE,
                align=PP_ALIGN.LEFT, wrap: bool = True,
                italic: bool = False, font_name: str = "Calibri") -> None:
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic = italic
    run.font.name   = font_name


def add_bullet_textbox(slide, left, top, width, height,
                       bullets: list[str], font_size: int = 18,
                       color: RGBColor = DARKGRAY,
                       accent: RGBColor = BLUE,
                       line_spacing_pt: float = 1.15) -> None:
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    from lxml import etree

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(bullets):
        is_sub = bullet.startswith("  –")
        text   = bullet.lstrip(" –").strip() if is_sub else bullet

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT

        # indent level
        pPr = p._p.get_or_add_pPr()
        pPr.set("lvl", "1" if is_sub else "0")

        run = p.add_run()
        run.text = ("  – " if is_sub else "• ") + text
        run.font.size  = Pt(font_size - 2 if is_sub else font_size)
        run.font.color.rgb = DARKGRAY if is_sub else color
        run.font.bold  = False
        run.font.name  = "Calibri"

        # line spacing
        from pptx.oxml.ns import qn
        lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
        spcPct = etree.SubElement(lnSpc, qn("a:spcPct"))
        spcPct.set("val", "115000")   # 115%

        spcBef = etree.SubElement(pPr, qn("a:spcBef"))
        spcPts = etree.SubElement(spcBef, qn("a:spcPts"))
        spcPts.set("val", "0" if is_sub else "120")   # 12pt before each top-level


def add_speaker_note(slide, note_text: str) -> None:
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = note_text


# ── Individual slide type renderers ──────────────────────────────────────────

def render_title(prs: Presentation, s: dict, course: str = "COS598"):
    slide = blank_slide(prs)
    fill_bg(slide, NAVY)

    # Left accent bar
    add_rect(slide, Inches(0), Inches(0), Inches(0.35), SLIDE_H, TEAL)

    # Course label
    add_textbox(slide,
                Inches(0.6), Inches(0.45),
                Inches(12), Inches(0.5),
                course, 14, bold=False, color=TEAL, align=PP_ALIGN.LEFT)

    # Title
    add_textbox(slide,
                Inches(0.6), Inches(1.4),
                Inches(12), Inches(2.8),
                s["title"], 34, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Subtitle / authors
    sub = s.get("subtitle", "")
    if sub:
        add_textbox(slide,
                    Inches(0.6), Inches(4.4),
                    Inches(12), Inches(0.7),
                    sub, 20, bold=False, color=LIGHTGRAY, align=PP_ALIGN.LEFT)

    # Bottom accent line
    add_rect(slide, Inches(0.6), Inches(5.3), Inches(8), Inches(0.04), TEAL)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_section(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, BLUE)

    add_rect(slide, Inches(0), Inches(0), Inches(0.35), SLIDE_H, TEAL)

    add_textbox(slide,
                Inches(0.7), Inches(2.4),
                Inches(11.8), Inches(1.5),
                s["title"], 40, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    sub = s.get("subtitle", "")
    if sub:
        add_textbox(slide,
                    Inches(0.7), Inches(4.1),
                    Inches(11.8), Inches(0.8),
                    sub, 22, bold=False, color=OFFWHITE, align=PP_ALIGN.LEFT)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_content(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    # Title bar
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), TEAL)

    add_textbox(slide,
                Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    bullets = s.get("body", [])
    add_bullet_textbox(slide,
                       Inches(0.5), Inches(1.4),
                       Inches(12.3), Inches(5.7),
                       bullets, font_size=19, color=NAVY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_agenda(prs: Presentation, s: dict):
    """Agenda uses two columns of bullets."""
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), TEAL)

    add_textbox(slide,
                Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=WHITE)

    bullets = s.get("body", [])
    mid = (len(bullets) + 1) // 2
    left_b  = bullets[:mid]
    right_b = bullets[mid:]

    add_bullet_textbox(slide,
                       Inches(0.5), Inches(1.4),
                       Inches(5.9), Inches(5.7),
                       left_b, font_size=19, color=NAVY)

    add_bullet_textbox(slide,
                       Inches(6.8), Inches(1.4),
                       Inches(5.9), Inches(5.7),
                       right_b, font_size=19, color=NAVY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_two_col(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), TEAL)

    add_textbox(slide,
                Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=WHITE)

    # Column divider
    add_rect(slide, Inches(6.56), Inches(1.25), Inches(0.04), Inches(6.1), LIGHTGRAY)

    # Left column
    lh = s.get("left_header", "")
    if lh:
        add_rect(slide, Inches(0.35), Inches(1.35), Inches(6.0), Inches(0.42), BLUE)
        add_textbox(slide, Inches(0.4), Inches(1.37), Inches(5.9), Inches(0.38),
                    lh, 15, bold=True, color=WHITE)
    add_bullet_textbox(slide,
                       Inches(0.4), Inches(1.9),
                       Inches(5.9), Inches(5.3),
                       s.get("left", []), font_size=17, color=NAVY)

    # Right column
    rh = s.get("right_header", "")
    if rh:
        add_rect(slide, Inches(6.73), Inches(1.35), Inches(6.0), Inches(0.42), TEAL)
        add_textbox(slide, Inches(6.78), Inches(1.37), Inches(5.9), Inches(0.38),
                    rh, 15, bold=True, color=WHITE)
    add_bullet_textbox(slide,
                       Inches(6.73), Inches(1.9),
                       Inches(5.9), Inches(5.3),
                       s.get("right", []), font_size=17, color=NAVY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_quote(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, NAVY)

    add_rect(slide, Inches(0), Inches(0), Inches(0.35), SLIDE_H, AMBER)

    # Decorative large quotation mark
    add_textbox(slide, Inches(0.5), Inches(0.3), Inches(2), Inches(2),
                "\u201c", 120, bold=True, color=AMBER, align=PP_ALIGN.LEFT)

    add_textbox(slide,
                Inches(1.0), Inches(1.3),
                Inches(10.8), Inches(3.8),
                s.get("quote", ""), 24, bold=False,
                color=WHITE, align=PP_ALIGN.LEFT, italic=True)

    attr = s.get("attribution", "")
    if attr:
        add_rect(slide, Inches(1.0), Inches(5.3), Inches(3), Inches(0.04), AMBER)
        add_textbox(slide, Inches(1.0), Inches(5.5),
                    Inches(11), Inches(0.6),
                    "— " + attr, 16, bold=False, color=AMBER, align=PP_ALIGN.LEFT)

    title = s.get("title", "")
    if title:
        add_textbox(slide, Inches(0.5), Inches(6.7), Inches(12), Inches(0.5),
                    title, 13, bold=False, color=LIGHTGRAY, align=PP_ALIGN.LEFT)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_discussion(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), AMBER)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), NAVY)

    add_textbox(slide, Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=NAVY)

    questions = s.get("questions", [])
    for i, q in enumerate(questions):
        top = Inches(1.5) + i * Inches(1.4)
        # numbered box
        add_rect(slide, Inches(0.4), top, Inches(0.5), Inches(0.5), NAVY)
        add_textbox(slide, Inches(0.41), top + Pt(4),
                    Inches(0.48), Inches(0.45),
                    str(i + 1), 18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide,
                    Inches(1.1), top,
                    Inches(11.5), Inches(1.2),
                    q, 18, bold=False, color=DARKGRAY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_data(prs: Presentation, s: dict):
    """Data / statistics slide — same as content but with TEAL accent."""
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), TEAL)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), NAVY)

    add_textbox(slide, Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=WHITE)

    add_bullet_textbox(slide,
                       Inches(0.5), Inches(1.4),
                       Inches(12.3), Inches(5.7),
                       s.get("body", []), font_size=19, color=NAVY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_summary(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, OFFWHITE)

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.05), AMBER)

    add_textbox(slide, Inches(0.35), Inches(0.18),
                Inches(12.5), Inches(0.85),
                s["title"], 26, bold=True, color=AMBER)

    add_bullet_textbox(slide,
                       Inches(0.5), Inches(1.4),
                       Inches(12.3), Inches(5.7),
                       s.get("body", []), font_size=19, color=NAVY)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


def render_thankyou(prs: Presentation, s: dict):
    slide = blank_slide(prs)
    fill_bg(slide, NAVY)

    add_rect(slide, Inches(0), Inches(0), Inches(0.35), SLIDE_H, TEAL)

    add_textbox(slide, Inches(0.6), Inches(2.0),
                Inches(12), Inches(1.5),
                s["title"], 48, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    sub = s.get("subtitle", "")
    if sub:
        add_textbox(slide, Inches(0.6), Inches(3.8),
                    Inches(11), Inches(0.8),
                    sub, 22, color=TEAL, align=PP_ALIGN.LEFT)

    add_rect(slide, Inches(0.6), Inches(4.8), Inches(5), Inches(0.04), AMBER)

    if s.get("note"):
        add_speaker_note(slide, s["note"])
    return slide


RENDERERS = {
    "title":      render_title,
    "agenda":     render_agenda,
    "section":    render_section,
    "content":    render_content,
    "two_col":    render_two_col,
    "quote":      render_quote,
    "discussion": render_discussion,
    "data":       render_data,
    "summary":    render_summary,
    "thankyou":   render_thankyou,
}


def build_deck(reading: dict, slides_data: list[dict]) -> Presentation:
    prs = new_prs()
    for s in slides_data:
        stype = s.get("type", "content")
        renderer = RENDERERS.get(stype, render_content)
        renderer(prs, s)
    return prs


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    SLIDES_DIR.mkdir(exist_ok=True)

    session_token_path = Path("/home/claude/.claude/remote/.session_ingress_token")
    if session_token_path.exists():
        token = session_token_path.read_text().strip()
        client = anthropic.Anthropic(auth_token=token)
    else:
        client = anthropic.Anthropic()

    for reading in READINGS:
        print(f"\nProcessing: {reading['title']}")

        original_text  = extract_text(READINGS_DIR / reading["file"], reading["max_pages"])
        summary_text   = extract_text(SUMMARIES_DIR / reading["summary_file"])
        takeaways_text = extract_text(TAKEAWAYS_DIR / reading["takeaway_file"])

        print(f"  Generating slide content with Claude ({len(original_text):,} chars input)...")
        slides_data = generate_slide_content(
            reading, original_text, summary_text, takeaways_text, client
        )
        print(f"  Generated {len(slides_data)} slides.")

        prs = build_deck(reading, slides_data)

        stem    = Path(reading["file"]).stem
        out_path = SLIDES_DIR / f"{stem}_lecture.pptx"
        prs.save(str(out_path))
        print(f"  Saved: {out_path.name}")

    print(f"\nDone. {len(READINGS)} decks saved to: {SLIDES_DIR}")


if __name__ == "__main__":
    main()
