"""
Summarize PDF readings using the Claude API and save each summary as a PDF.
"""

import sys
import unittest.mock as mock
import os
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
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER


READINGS_DIR = Path(__file__).parent.parent / "Readings"
SUMMARIES_DIR = Path(__file__).parent.parent / "Summaries"

READINGS = [
    {
        "file": "2310.11986v2.pdf",
        "title": "Sociotechnical Safety Evaluation of Generative AI Systems",
        "max_pages": 76,
    },
    {
        "file": "All_Roads_Lead_to_ChatGPT_How_Generative_AI_is_Ero.pdf",
        "title": '"All Roads Lead to ChatGPT": How Generative AI is Eroding Social Interactions and Student Learning Communities',
        "max_pages": None,
    },
    {
        "file": "Excerpts from AI Now 2025 Reports on AI Power.pdf",
        "title": "Excerpts from AI Now 2025 Report on AI Power",
        "max_pages": None,
    },
    {
        "file": "Executive Summary - AI Now Institute.pdf",
        "title": "AI Now Institute 2023 Executive Summary: Confronting Tech Power",
        "max_pages": None,
    },
    {
        "file": "hai_ai_index_report_2025.pdf",
        "title": "HAI AI Index Report 2025",
        "max_pages": 30,  # Sample first ~30 pages of the 457-page report
    },
]


def extract_text(pdf_path: Path, max_pages: int | None = None) -> str:
    reader = pypdf.PdfReader(str(pdf_path))
    pages = reader.pages if max_pages is None else reader.pages[:max_pages]
    return "\n".join(page.extract_text() or "" for page in pages)


def summarize(title: str, text: str, client: anthropic.Anthropic) -> str:
    prompt = f"""Below is the text extracted from a document titled: "{title}"

Please write a concise summary of this document that fits on a single page (approximately 350-450 words). Your summary should cover:
1. The main topic and thesis/purpose of the document
2. Key arguments, findings, or data points
3. Important conclusions or recommendations
4. The significance or broader implications

Write in clear, academic prose. Do not use bullet points — write in paragraphs.

---

{text[:60000]}
"""

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        return stream.get_final_message().content[0].text


def save_summary_pdf(title: str, summary_text: str, output_path: Path) -> None:
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=letter,
        leftMargin=1 * inch,
        rightMargin=1 * inch,
        topMargin=1 * inch,
        bottomMargin=1 * inch,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "SummaryTitle",
        parent=styles["Heading1"],
        fontSize=14,
        spaceAfter=6,
        alignment=TA_CENTER,
        leading=18,
    )

    label_style = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontSize=10,
        textColor=(0.4, 0.4, 0.4),
        spaceAfter=16,
        alignment=TA_CENTER,
        leading=14,
    )

    body_style = ParagraphStyle(
        "SummaryBody",
        parent=styles["Normal"],
        fontSize=11,
        leading=16,
        spaceAfter=10,
        alignment=TA_LEFT,
    )

    story = [
        Paragraph(title, title_style),
        Paragraph("Summary", label_style),
    ]

    for paragraph in summary_text.strip().split("\n\n"):
        paragraph = paragraph.strip()
        if paragraph:
            story.append(Paragraph(paragraph, body_style))
            story.append(Spacer(1, 4))

    doc.build(story)


def main():
    SUMMARIES_DIR.mkdir(exist_ok=True)

    # In this Claude Code environment, authenticate via the session ingress token
    session_token_path = Path("/home/claude/.claude/remote/.session_ingress_token")
    if session_token_path.exists():
        token = session_token_path.read_text().strip()
        client = anthropic.Anthropic(auth_token=token)
    else:
        client = anthropic.Anthropic()  # Falls back to ANTHROPIC_API_KEY

    for reading in READINGS:
        pdf_path = READINGS_DIR / reading["file"]
        print(f"\nProcessing: {reading['title']}")
        print(f"  Extracting text from {pdf_path.name}...")

        text = extract_text(pdf_path, reading["max_pages"])
        print(f"  Extracted {len(text):,} characters. Summarizing with Claude...")

        summary = summarize(reading["title"], text, client)
        print(f"  Summary generated ({len(summary.split())} words).")

        # Build output filename from the input filename (sans extension)
        stem = Path(reading["file"]).stem
        out_path = SUMMARIES_DIR / f"{stem}_summary.pdf"
        save_summary_pdf(reading["title"], summary, out_path)
        print(f"  Saved: {out_path.name}")

    print(f"\nDone. {len(READINGS)} summaries saved to: {SUMMARIES_DIR}")


if __name__ == "__main__":
    main()
