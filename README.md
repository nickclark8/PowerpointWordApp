# PowerPoint Word App

A lightweight Python web app that lets users upload a PowerPoint presentation (`.pptx`), extracts slide content, creates an executive-style summary, and returns it as a downloadable Word document (`.docx`).

## Features
- Upload `.pptx` presentations via web UI.
- Extract text from slide titles, body shapes, and notes.
- Generate an executive summary with:
  - Key highlights
  - Strategic implications
  - Recommended actions
  - Slide-by-slide snapshot
- Download summary as a Word document.

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
