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
class BusinessCaseSection:
    title: str
    bullets: list[str]


@dataclass
class BusinessCase:
    source_name: str
    generated_on: str
    brand_name: str
    sections: list[BusinessCaseSection]


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


def extract_supporting_text(file_bytes: bytes, extension: str) -> str:
    normalized_ext = extension.lower().lstrip(".")
    if normalized_ext == "txt":
        return file_bytes.decode("utf-8", errors="ignore").strip()

    if normalized_ext == "docx":
        with ZipFile(BytesIO(file_bytes), "r") as archive:
            if "word/document.xml" not in archive.namelist():
                return ""
            doc_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
            return "\n".join(_read_slide_texts(doc_xml)).strip()

    raise ValueError("Unsupported supporting document format")


def _trim_sentence(text: str, max_words: int = 18) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]).rstrip(".,;:") + "..."


def _extract_unique_points(slides: list[SlideContent]) -> list[str]:
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
    return unique_points


def _top_themes(points: list[str], max_themes: int = 4) -> list[str]:
    token_counter = Counter(
        token.strip(".,:;!?()[]{}\"'").lower()
        for point in points
        for token in point.split()
        if len(token) > 4
    )
    return [word for word, _ in token_counter.most_common(max_themes)]


def build_business_case(slides: list[SlideContent], source_name: str, knowledge_context: str = "") -> BusinessCase:
    unique_points = _extract_unique_points(slides)
    themes = _top_themes(unique_points)
    context_snippet = _trim_sentence(knowledge_context.strip(), 24) if knowledge_context.strip() else ""

    default_financial_placeholders = [
        "Revenue impact: [Insert annual uplift estimate, assumptions, and confidence range]",
        "Cost impact: [Split fixed vs variable and one-time vs recurring costs]",
        "Investment required: [Capex/Opex totals and funding profile]",
        "ROI / NPV / Payback: [Insert horizon, discount rate, and scenario outputs]",
    ]

    sections = [
        BusinessCaseSection(
            "Executive Summary",
            [
                f"This business case synthesizes findings from {source_name} for leadership decision-making.",
                (
                    f"Primary strategic themes observed: {', '.join(themes)}."
                    if themes
                    else "Primary strategic themes were not explicit; narrative should be tightened before approval."
                ),
                "Recommendation is to proceed conditionally with stage-gate governance and quantified targets.",
            ],
        ),
        BusinessCaseSection(
            "Problem Statement / Opportunity",
            [
                *([f"Knowledge input considered: {context_snippet}"] if context_snippet else []),
                "Current-state pain points and growth opportunities should be validated with stakeholder owners.",
                "Define baseline performance metrics to measure incremental impact.",
            ],
        ),
        BusinessCaseSection(
            "Strategic Rationale",
            [
                "Align the initiative to explicit corporate priorities and the accountable decision-maker.",
                "Clarify how this initiative differentiates versus competing investments.",
                "Document dependencies across teams, systems, and third parties.",
            ],
        ),
        BusinessCaseSection(
            "Options Considered (including Do Nothing)",
            [
                "Option A: Full-scope implementation with accelerated timeline.",
                "Option B: Phased implementation prioritizing highest-value segments.",
                "Option C (Do Nothing): Maintain current state and accept opportunity cost/risk exposure.",
                "Evaluate trade-offs for cost, speed, execution risk, and strategic fit.",
            ],
        ),
        BusinessCaseSection("Financial Impact", default_financial_placeholders),
        BusinessCaseSection(
            "Key Assumptions",
            [
                "Assumption 1: Adoption rate trajectory and ramp period.",
                "Assumption 2: Unit economics remain within historical variance bands.",
                "Assumption 3: Required capabilities and staffing can be mobilized on schedule.",
            ],
        ),
        BusinessCaseSection(
            "Risks and Mitigations",
            [
                "Execution risk: Stage-gate delivery with named owners and milestone controls.",
                "Financial risk: Scenario modeling (best/base/worst) with trigger thresholds.",
                "Stakeholder risk: Governance cadence with escalation paths and decision rights.",
            ],
        ),
        BusinessCaseSection(
            "Implementation Plan",
            [
                "0-30 days: Confirm scope, governance, and KPI baseline.",
                "31-90 days: Pilot execution, capture learnings, and refine forecasts.",
                "90+ days: Scale rollout subject to gate approval and KPI attainment.",
            ],
        ),
        BusinessCaseSection(
            "Success Metrics / KPIs",
            [
                "Financial: Revenue uplift, gross margin impact, and payback progress.",
                "Operational: Cycle time, quality, and throughput improvements.",
                "Adoption: User/segment activation and retention against targets.",
            ],
        ),
        BusinessCaseSection(
            "Recommendation",
            [
                "Proceed with a phased implementation in base case, contingent on validated financial model inputs.",
                "Escalate to leadership if assumptions move outside predefined tolerance bands.",
            ],
        ),
    ]

    return BusinessCase(
        source_name=source_name,
        generated_on=date.today().isoformat(),
        brand_name="Business Case Copilot",
        sections=sections,
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


def generate_business_case_docx(case: BusinessCase) -> bytes:
    paragraphs = [
        f"{case.brand_name} | Board-Ready Business Case",
        f"Source presentation: {case.source_name}",
        f"Generated on: {case.generated_on}",
        "",
    ]

    for i, section in enumerate(case.sections, start=1):
        paragraphs.append(f"{i}) {section.title}")
        paragraphs.extend([f"• {item}" for item in section.bullets])
        paragraphs.append("")

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
