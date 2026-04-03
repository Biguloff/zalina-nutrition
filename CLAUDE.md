# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Nutritionist recommendation PDF generator for Залина Бигулова (a naturopathic nutritionist). Each client gets a personalized multi-page A4 PDF with blood test analysis, supplement recommendations, and dietary protocols. There is also a web version (`index.html`) serving as a landing page.

## Generating PDFs

Each client has a standalone Python script that produces a styled PDF:

```bash
python3 generate_pdf.py                # → Рекомендации_нутрициолога_Бигулова.pdf
python3 generate_pdf_alekhina.py       # → Рекомендации_нутрициолога_Алехина_Стояна.pdf
python3 generate_pdf_filonova.py       # → Рекомендации_нутрициолога_Филонова_Татьяна.pdf
```

**Dependency:** `reportlab` — install with `pip3 install reportlab`.

## Architecture

- **Per-client scripts** — each `generate_pdf_*.py` is a self-contained script with its own data, styles, and layout logic. There is no shared library; each script duplicates the ReportLab boilerplate, color constants, style definitions, and page template callbacks. When making style or layout changes, they must be applied to each script individually.
- **Fonts** — Georgia (serif, headings) and Arial (sans-serif, body) loaded from macOS system fonts at `/System/Library/Fonts/Supplemental/`. Scripts will not run on non-macOS systems without adjusting `FONT_DIR`.
- **Design spec** — `docs/superpowers/specs/2026-03-30-pdf-redesign-design.md` defines the "Warm Organic" aesthetic (cream backgrounds, sage/copper accents, Georgia typography).
- **Web page** — `index.html` is a standalone Russian-language page with the same sage/cream/copper color palette. No build tooling; opens directly in a browser.

## Design System (shared across all scripts)

| Token | Hex | Role |
|-------|-----|------|
| SAGE | #5A7A60 | Primary accent, section headers |
| COPPER | #C17744 | Secondary accent, subheadings |
| CREAM | #FDF8F0 | Page background |
| CHARCOAL | #2D3436 | Body text |
| Status colors | red #DC4E42 / amber #D4A04B / green #3D8B47 | Blood test result markers |

## Key Conventions

- All content is in **Russian** (UTF-8). Use Russian text for any new content.
- PDF layout uses ReportLab's `platypus` flowable system (`SimpleDocTemplate`, `Paragraph`, `Table`, `KeepTogether`).
- Custom header/footer drawn via `canvas` callbacks in each script's `build_pdf()` or equivalent function.
- Blood test tables use a consistent 4-column format: name, result, reference range, status.
- Supplement protocol tables use 4 columns: supplement, dosage, timing, course duration.
