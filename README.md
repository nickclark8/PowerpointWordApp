# PowerPoint Word App

A lightweight Python web app that lets users upload a PowerPoint presentation (`.pptx`), extracts slide content, and returns a branded, executive-ready business case as a downloadable Word document (`.docx`).

## Features
- Upload `.pptx` presentations via web UI.
- Extract text from slide titles, body shapes, and notes.
- Generate a **Business Case Copilot** output with standard sections:
  - Executive Summary
  - Problem Statement / Opportunity
  - Strategic Rationale
  - Options Considered (including Do Nothing)
  - Financial Impact placeholders (revenue, cost, investment, ROI/NPV/payback)
  - Key Assumptions
  - Risks and Mitigations
  - Implementation Plan
  - Success Metrics / KPIs
  - Recommendation
- Optional context field to include pasted strategic/knowledge notes.
- Optional supporting knowledge file upload (`.txt` or `.docx`) to enrich the business case context.
- Download the result as a branded Word document.

## Quickstart

```bash
python -m venv .venv
source .venv/bin/activate
python app.py
```

Then open `http://localhost:5000`.

## Run tests

```bash
pytest
```
