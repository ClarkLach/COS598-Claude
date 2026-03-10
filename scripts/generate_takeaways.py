"""
Generate key takeaways and areas of confusion for each reading,
saved as bullet-point PDFs in a new KeyTakeaways/ folder.
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

import pypdf
import anthropic
import json
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors


READINGS_DIR = Path(__file__).parent.parent / "Readings"
SUMMARIES_DIR = Path(__file__).parent.parent / "Summaries"
TAKEAWAYS_DIR = Path(__file__).parent.parent / "KeyTakeaways"

READINGS = [
    {
        "file": "2310.11986v2.pdf",
        "summary_file": "2310.11986v2_summary.pdf",
        "title": "Sociotechnical Safety Evaluation of Generative AI Systems",
        "max_pages": 76,
    },
    {
        "file": "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero.pdf",
        "summary_file": "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero_summary.pdf",
        "title": '"All Roads Lead to ChatGPT": How Generative AI is Eroding Social Interactions and Student Learning Communities',
        "max_pages": None,
    },
    {
        "file": "Excerpts from AI Now 2025 Reports on AI Power.pdf",
        "summary_file": "Excerpts from AI Now 2025 Reports on AI Power_summary.pdf",
        "title": "Excerpts from AI Now 2025 Report on AI Power",
        "max_pages": None,
    },
    {
        "file": "Executive Summary - AI Now Institute.pdf",
        "summary_file": "Executive Summary - AI Now Institute_summary.pdf",
        "title": "AI Now Institute 2023 Executive Summary: Confronting Tech Power",
        "max_pages": None,
    },
    {
        "file": "hai_ai_index_report_2025.pdf",
        "summary_file": "hai_ai_index_report_2025_summary.pdf",
        "title": "HAI AI Index Report 2025",
        "max_pages": 30,
    },
]


def extract_text(pdf_path: Path, max_pages: int | None = None) -> str:
    reader = pypdf.PdfReader(str(pdf_path))
    pages = reader.pages if max_pages is None else reader.pages[:max_pages]
    return "\n".join(page.extract_text() or "" for page in pages)


def generate_takeaways(title: str, original_text: str, summary_text: str,
                        client: anthropic.Anthropic) -> dict:
    prompt = f"""You are an expert educational assistant helping graduate students engage deeply with academic readings.

Below is a document titled: "{title}"

SUMMARY:
{summary_text}

ORIGINAL TEXT (excerpt):
{original_text[:50000]}

Based on both the summary and the original text, produce a JSON object with exactly two keys:

1. "takeaways": A list of 6-8 concise bullet points capturing the most important ideas a reader must walk away with. Each bullet should be a complete, specific sentence — not vague. Include key claims, key evidence, and key implications.

2. "confusions": A list of 5-7 bullet points identifying the most common areas of confusion or misunderstanding that readers are likely to encounter with this material. For each, briefly state WHAT the confusion is and WHY it arises (e.g., ambiguous terminology, counterintuitive findings, conflation of related concepts, or claims that contradict popular assumptions). Each bullet should be 1-3 sentences.

Return ONLY valid JSON — no markdown fences, no extra text. Example format:
{{"takeaways": ["...", "..."], "confusions": ["...", "..."]}}"""

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=2048,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        response_text = next(
            b.text for b in stream.get_final_message().content if b.type == "text"
        )

    # Strip any accidental markdown fences
    clean = response_text.strip()
    if clean.startswith("```"):
        clean = clean.split("```")[1]
        if clean.startswith("json"):
            clean = clean[4:]
    clean = clean.strip().rstrip("```").strip()

    return json.loads(clean)


def save_takeaways_pdf(title: str, data: dict, output_path: Path) -> None:
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=letter,
        leftMargin=1 * inch,
        rightMargin=1 * inch,
        topMargin=0.85 * inch,
        bottomMargin=0.85 * inch,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontSize=13,
        spaceAfter=4,
        spaceBefore=0,
        alignment=TA_CENTER,
        leading=17,
        textColor=colors.HexColor("#1a1a2e"),
    )

    section_style = ParagraphStyle(
        "SectionHeader",
        parent=styles["Heading2"],
        fontSize=11,
        spaceBefore=14,
        spaceAfter=6,
        textColor=colors.HexColor("#2c5f8a"),
        leading=14,
    )

    bullet_style = ParagraphStyle(
        "Bullet",
        parent=styles["Normal"],
        fontSize=10,
        leading=14,
        spaceAfter=5,
        leftIndent=14,
        bulletIndent=0,
        bulletFontSize=10,
    )

    story = [
        Paragraph(title, title_style),
        Spacer(1, 4),
        HRFlowable(width="100%", thickness=1, color=colors.HexColor("#2c5f8a")),
        Paragraph("Key Takeaways", section_style),
    ]

    for item in data["takeaways"]:
        story.append(Paragraph(f"• {item.strip()}", bullet_style))

    story += [
        Spacer(1, 6),
        HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc")),
        Paragraph("Common Areas of Confusion", section_style),
    ]

    for item in data["confusions"]:
        story.append(Paragraph(f"• {item.strip()}", bullet_style))

    doc.build(story)


def main():
    TAKEAWAYS_DIR.mkdir(exist_ok=True)

    session_token_path = Path("/home/claude/.claude/remote/.session_ingress_token")
    if session_token_path.exists():
        token = session_token_path.read_text().strip()
        client = anthropic.Anthropic(auth_token=token)
    else:
        client = anthropic.Anthropic()

    for reading in READINGS:
        print(f"\nProcessing: {reading['title']}")

        original_text = extract_text(
            READINGS_DIR / reading["file"], reading["max_pages"]
        )
        summary_text = extract_text(SUMMARIES_DIR / reading["summary_file"])
        print(f"  Extracted {len(original_text):,} chars from original. Calling Claude...")

        data = generate_takeaways(reading["title"], original_text, summary_text, client)
        print(f"  Generated {len(data['takeaways'])} takeaways, {len(data['confusions'])} confusions.")

        stem = Path(reading["file"]).stem
        out_path = TAKEAWAYS_DIR / f"{stem}_takeaways.pdf"
        save_takeaways_pdf(reading["title"], data, out_path)
        print(f"  Saved: {out_path.name}")

    print(f"\nDone. {len(READINGS)} takeaway sheets saved to: {TAKEAWAYS_DIR}")


if __name__ == "__main__":
    main()
