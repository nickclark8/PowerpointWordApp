import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from app import render_index
from services.ppt_summarizer import (
    build_business_case,
    extract_slide_content,
    extract_supporting_text,
    generate_business_case_docx,
)


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


def test_extract_and_build_business_case_and_docx_generation():
    slides = extract_slide_content(_build_fake_pptx())
    assert len(slides) == 2
    assert slides[0].title == "Q4 Business Review"

    business_case = build_business_case(slides, "demo.pptx", knowledge_context="Align to FY strategic priority")
    assert business_case.source_name == "demo.pptx"
    assert business_case.sections
    assert any(section.title == "Financial Impact" for section in business_case.sections)

    docx_bytes = generate_business_case_docx(business_case)
    assert docx_bytes.startswith(b"PK")
    assert len(docx_bytes) > 500


def test_render_index_contains_form():
    html = render_index().decode("utf-8")
    assert "<form" in html
    assert "PowerPoint to Branded Business Case" in html
    assert 'name="knowledge_file"' in html


def test_extract_supporting_text_from_txt_and_docx():
    txt = extract_supporting_text(b"Strategic priority: improve retention", "txt")
    assert "improve retention" in txt

    minimal_docx = BytesIO()
    with ZipFile(minimal_docx, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr(
            "word/document.xml",
            """<?xml version='1.0' encoding='UTF-8'?><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>Board objective</w:t></w:r></w:p><w:p><w:r><w:t>Reduce churn by 2%</w:t></w:r></w:p></w:body></w:document>""",
        )
    docx_text = extract_supporting_text(minimal_docx.getvalue(), "docx")
    assert "Board objective" in docx_text
