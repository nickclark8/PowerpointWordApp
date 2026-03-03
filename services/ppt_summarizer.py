from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from io import BytesIO
import re
from typing import Iterable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

TEXT_PATTERN = re.compile(r"<a:t>(.*?)</a:t>", flags=re.DOTALL)
TAG_STRIP_PATTERN = re.compile(r"<[^>]+>")


@dataclass
class SlideContent:
    index: int
    title: str
    bullets: list[str]
    notes: list[str]


@dataclass
class ExecSummary:
    source_name: str
    generated_on: str
    key_highlights: list[str]
    strategic_implications: list[str]
    recommended_actions: list[str]
    slide_snapshots: list[str]


def _normalize_whitespace(value: str) -> str:
    return " ".join(TAG_STRIP_PATTERN.sub("", value).split())


def _read_slide_texts(xml_content: str) -> list[str]:
    return [
        _normalize_whitespace(chunk)
        for chunk in TEXT_PATTERN.findall(xml_content)
        if _normalize_whitespace(chunk)
    ]


def extract_slide_content(file_bytes: bytes) -> list[SlideContent]:
    slides: list[SlideContent] = []
    with ZipFile(BytesIO(file_bytes), "r") as archive:
        slide_files = sorted(
            [name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)),
        )

        for idx, slide_path in enumerate(slide_files, start=1):
            slide_xml = archive.read(slide_path).decode("utf-8", errors="ignore")
            texts = _read_slide_texts(slide_xml)
            title = texts[0] if texts else f"Slide {idx}"
            bullets = texts[1:] if len(texts) > 1 else []

            notes_path = f"ppt/notesSlides/notesSlide{idx}.xml"
            notes: list[str] = []
            if notes_path in archive.namelist():
                notes_xml = archive.read(notes_path).decode("utf-8", errors="ignore")
                notes = _read_slide_texts(notes_xml)

            slides.append(SlideContent(index=idx, title=title, bullets=bullets, notes=notes))

    return slides


def _trim_sentence(text: str, max_words: int = 18) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]).rstrip(".,;:") + "..."


def build_exec_summary(slides: list[SlideContent], source_name: str) -> ExecSummary:
    all_points: list[str] = []
    for slide in slides:
        all_points.extend(slide.bullets)
        all_points.extend(slide.notes)

    seen: set[str] = set()
    unique_points = []
    for point in all_points:
        normalized = point.lower()
        if len(point.split()) >= 4 and normalized not in seen:
            unique_points.append(point)
            seen.add(normalized)

    key_highlights = [_trim_sentence(p) for p in unique_points[:5]]
    token_counter = Counter(
        token.strip(".,:;!?()[]{}\"'").lower()
        for point in unique_points
        for token in point.split()
        if len(token) > 4
    )
    themes = [word for word, _ in token_counter.most_common(3)]

    strategic_implications = [
        (
            f"Core narrative emphasizes {', '.join(themes)}; align decisions and communication around these themes."
            if themes
            else "Presentation is descriptive; tighten the strategic narrative for executive stakeholders."
        ),
        "Dependencies and sequencing should be tracked with clear ownership across functions.",
        "Translate commitments into measurable outcomes before the next review cadence.",
    ]

    recommended_actions = [
        "Confirm top priorities and attach a KPI to each.",
        "Document top risks with mitigation owners and deadlines.",
        "Prepare a 30-60-90 day execution plan for leadership sign-off.",
    ]

    slide_snapshots = []
    for slide in slides:
        top_bullets = "; ".join(_trim_sentence(b, 10) for b in slide.bullets[:2]) or "No key bullets captured."
        slide_snapshots.append(f"Slide {slide.index} — {slide.title}: {top_bullets}")

    return ExecSummary(
        source_name=source_name,
        generated_on=date.today().isoformat(),
        key_highlights=key_highlights or ["No clear highlights extracted from slide text."],
        strategic_implications=strategic_implications,
        recommended_actions=recommended_actions,
        slide_snapshots=slide_snapshots,
    )


def _paragraphs_to_document_xml(paragraphs: Iterable[str]) -> str:
    body = []
    for text in paragraphs:
        body.append(f"<w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p>")
    body_xml = "".join(body)
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        f"<w:body>{body_xml}<w:sectPr/></w:body></w:document>"
    )


def generate_exec_docx(summary: ExecSummary) -> bytes:
    paragraphs = [
        "Executive Summary",
        f"Source presentation: {summary.source_name}",
        f"Generated on: {summary.generated_on}",
        "",
        "1) Key Highlights",
        *[f"• {item}" for item in summary.key_highlights],
        "",
        "2) Strategic Implications",
        *[f"• {item}" for item in summary.strategic_implications],
        "",
        "3) Recommended Actions",
        *[f"• {item}" for item in summary.recommended_actions],
        "",
        "4) Slide-by-Slide Snapshot",
        *[f"{i+1}. {item}" for i, item in enumerate(summary.slide_snapshots)],
    ]

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    package_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", package_rels)
        docx.writestr("word/document.xml", _paragraphs_to_document_xml(paragraphs))

    return output.getvalue()
