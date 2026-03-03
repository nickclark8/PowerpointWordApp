import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from app import render_index
from services.ppt_summarizer import build_exec_summary, extract_slide_content, generate_exec_docx


def _build_fake_pptx() -> bytes:
    slide1 = """<p:sld xmlns:a='a'><a:t>Q4 Business Review</a:t><a:t>Revenue grew 18% year over year</a:t><a:t>Operating margin improved by 3 points</a:t></p:sld>"""
    slide2 = """<p:sld xmlns:a='a'><a:t>Priorities</a:t><a:t>Expand enterprise accounts in EMEA and APAC</a:t><a:t>Reduce onboarding time by two weeks</a:t></p:sld>"""
    notes1 = """<p:notes xmlns:a='a'><a:t>Focus on top-tier segments first</a:t></p:notes>"""

    out = BytesIO()
    with ZipFile(out, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("ppt/slides/slide1.xml", slide1)
        zf.writestr("ppt/slides/slide2.xml", slide2)
        zf.writestr("ppt/notesSlides/notesSlide1.xml", notes1)
    return out.getvalue()


def test_extract_and_summarize_and_docx_generation():
    slides = extract_slide_content(_build_fake_pptx())
    assert len(slides) == 2
    assert slides[0].title == "Q4 Business Review"

    summary = build_exec_summary(slides, "demo.pptx")
    assert summary.source_name == "demo.pptx"
    assert summary.key_highlights

    docx_bytes = generate_exec_docx(summary)
    assert docx_bytes.startswith(b"PK")
    assert len(docx_bytes) > 500


def test_render_index_contains_form():
    html = render_index().decode("utf-8")
    assert "<form" in html
    assert "PowerPoint to Executive Word Summary" in html
