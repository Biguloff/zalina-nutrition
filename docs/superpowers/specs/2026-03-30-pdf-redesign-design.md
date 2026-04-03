# PDF Redesign — Nutritionist Letterhead for Zalina Bigulova

## Overview

Redesign the 10-page nutritionist recommendation PDF to use a "Warm Organic" aesthetic — warm cream tones, serif typography, soft natural feel. The document should look like a personalized letter from a trusted naturopathic specialist, not a generic clinical report.

## Design Decisions

| Element | Choice | Rationale |
|---------|--------|-----------|
| Aesthetic | Warm Organic | Cream tones, Georgia serif, natural feel — matches naturopathic brand |
| Header | Centered, elegant | Name centered, thin ornamental line, credentials below — like a high-end invitation |
| Tables | Classic with dark header + row alternation + icons | Familiar medical format with warm color palette |
| Footer | Minimal | Thin line, disclaimer, name + page number — doesn't compete with content |

## Color Palette

| Token | Hex | Usage |
|-------|-----|-------|
| `bg-page` | #faf6f0 | Page background (warm cream) |
| `bg-alt-row` | #f5efe5 | Alternating table rows |
| `text-primary` | #3d3225 | Body text (dark chocolate) |
| `text-secondary` | #5a4a35 | Secondary text, table header bg |
| `text-muted` | #8a7a60 | Captions, notes |
| `accent-gold` | #a08860 | Accent color, section labels, ornaments |
| `border` | #d4c4a8 | Dividers, table borders |
| `border-light` | #ece5d8 | Subtle card borders |
| `status-alert` | #c0392b | Deficit / elevated markers |
| `status-warn` | #b8860b | Suboptimal markers |
| `status-ok` | #2e7d32 | Normal markers |
| `white` | #ffffff | Table row background |

## Typography

- **Headings:** Georgia (serif) — document title, section headers, recommendation values
- **Body:** Arial (sans-serif) — table cells, bullet points, body text, labels
- **Accent text:** Georgia Italic — subtitle, specialist note, footer name

## Page Structure (A4, 210x297mm)

### Header (every page)
- Centered layout
- Top: "НУТРИЦИОЛОГ-НАТУРОПАТ" — small caps, `accent-gold`, letter-spacing 5px
- Name: "Залина Бигулова" — Georgia Bold, 22px, `text-primary`
- Decorative line: 50px, 1.5px, `accent-gold`, centered
- Credentials: small text, `accent-gold`
- Contact info: small text, `text-muted`
- Bottom border: 1px solid `border`

### Footer (every page)
- Thin line: 1px solid `border`
- Disclaimer text: 7px, `text-muted`, centered
- Bottom row: "Залина Бигулова | Нутрициолог" left, "Стр. X из Y" right

### Content Area
- Margins: 20mm left/right, header ~32mm top, footer ~18mm bottom
- Section headers: Georgia Bold, 14px, `text-primary`, preceded by thin `border` line
- Section number labels: small caps, `accent-gold`, letter-spacing 2px

### Analysis Tables
- Header row: `text-secondary` (#5a4a35) background, cream text
- Body rows: alternating `white` / `bg-alt-row`
- Cell borders: 1px solid `border-light`
- Status text: colored bold (`status-alert`, `status-warn`, `status-ok`)
- Column widths: name 40%, result 22%, norm 20%, status 18%

### Recommendation Blocks
- Background: rgba(90,74,53,0.04)
- Border: 1px solid `border`
- Border-radius: 6px
- Icon circle: 20px, `accent-gold` background, white letter (D, Fe, I, Zn)
- Label: 7px uppercase, `accent-gold`, letter-spacing 2px
- Value: Georgia Bold, 13px, `text-primary`
- Note: 8px, `text-muted`

### Protocol Table (final section)
- Same style as analysis tables
- 4 columns: supplement, dosage, timing, course

### Signature Block (last page)
- "Нутрициолог:" left
- Signature line: thin `text-muted` underline
- "Бигулова З.К." right, Georgia Italic below: "подпись"

## Content (unchanged from current version)
1. Title page — client info + key findings summary table
2. Full analysis results — hormones, thyroid, lipids, vitamins/minerals, other markers
3. Deficit correction — vitamin D, ferritin, iodine, zinc with dose boxes
4. Cholesterol management — perimenopause connection, omega-3, fiber, sterols
5. Hormonal support — FSH/LH explanation, phytoestrogens, adaptogens
6. Weight loss plan — 6 steps: protein, Mediterranean diet, meal timing, exercise, sleep, safe pace
7. Supplement protocol table + control analyses + sources
8. Signature page

## What Changes vs Current Version
- **Remove:** green header bar, copper accent line, sage/green color scheme
- **Add:** warm cream background, centered header, gold accents, Georgia serif headings
- **Restyle:** tables from green headers to dark brown (#5a4a35) headers with cream alternating rows
- **Restyle:** recommendation boxes from green-bordered to gold-icon bordered cards
- **Restyle:** footer from green-themed to minimal warm line
- **Keep:** all content, page structure, 15 sources, signature block
