#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate professional nutritionist recommendation PDF
for Zalina Bigulova
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import os

# ─── Colors ───
SAGE = HexColor("#5A7A60")
SAGE_LIGHT = HexColor("#A8C5AE")
SAGE_BG = HexColor("#EDF5EF")
COPPER = HexColor("#C17744")
CREAM = HexColor("#FDF8F0")
CREAM_DARK = HexColor("#F0EAE0")
CHARCOAL = HexColor("#2D3436")
GRAY = HexColor("#6B7280")
LIGHT_GRAY = HexColor("#F3F4F6")
RED_SOFT = HexColor("#DC4E42")
RED_BG = HexColor("#FEF2F2")
AMBER = HexColor("#D4A04B")
AMBER_BG = HexColor("#FFFBEB")
GREEN = HexColor("#3D8B47")
GREEN_BG = HexColor("#F0FFF4")
WHITE = HexColor("#FFFFFF")
BORDER_LIGHT = HexColor("#E5E7EB")

# ─── Register Fonts ───
FONT_DIR = "/System/Library/Fonts/Supplemental/"
pdfmetrics.registerFont(TTFont("Georgia", os.path.join(FONT_DIR, "Georgia.ttf")))
pdfmetrics.registerFont(TTFont("Georgia-Bold", os.path.join(FONT_DIR, "Georgia Bold.ttf")))
pdfmetrics.registerFont(TTFont("Georgia-Italic", os.path.join(FONT_DIR, "Georgia Italic.ttf")))
pdfmetrics.registerFont(TTFont("Georgia-BoldItalic", os.path.join(FONT_DIR, "Georgia Bold Italic.ttf")))
pdfmetrics.registerFont(TTFont("Arial", os.path.join(FONT_DIR, "Arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Bold", os.path.join(FONT_DIR, "Arial Bold.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Italic", os.path.join(FONT_DIR, "Arial Narrow Italic.ttf")))

pdfmetrics.registerFontFamily(
    "Georgia", normal="Georgia", bold="Georgia-Bold",
    italic="Georgia-Italic", boldItalic="Georgia-BoldItalic"
)
pdfmetrics.registerFontFamily(
    "Arial", normal="Arial", bold="Arial-Bold", italic="Arial-Italic"
)

W, H = A4  # 210 x 297 mm

# ─── Styles ───
def make_styles():
    s = {}
    s["title"] = ParagraphStyle(
        "title", fontName="Georgia-Bold", fontSize=18, leading=24,
        textColor=CHARCOAL, alignment=TA_CENTER, spaceAfter=4*mm
    )
    s["subtitle"] = ParagraphStyle(
        "subtitle", fontName="Georgia-Italic", fontSize=11, leading=15,
        textColor=SAGE, alignment=TA_CENTER, spaceAfter=6*mm
    )
    s["h1"] = ParagraphStyle(
        "h1", fontName="Georgia-Bold", fontSize=15, leading=20,
        textColor=SAGE, spaceBefore=8*mm, spaceAfter=4*mm
    )
    s["h2"] = ParagraphStyle(
        "h2", fontName="Georgia-Bold", fontSize=12, leading=16,
        textColor=CHARCOAL, spaceBefore=5*mm, spaceAfter=3*mm
    )
    s["h3"] = ParagraphStyle(
        "h3", fontName="Arial-Bold", fontSize=10, leading=14,
        textColor=COPPER, spaceBefore=4*mm, spaceAfter=2*mm
    )
    s["body"] = ParagraphStyle(
        "body", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, alignment=TA_JUSTIFY, spaceAfter=2*mm
    )
    s["body_bold"] = ParagraphStyle(
        "body_bold", fontName="Arial-Bold", fontSize=9, leading=13.5,
        textColor=CHARCOAL, spaceAfter=2*mm
    )
    s["small"] = ParagraphStyle(
        "small", fontName="Arial", fontSize=7.5, leading=11,
        textColor=GRAY, spaceAfter=1*mm
    )
    s["small_italic"] = ParagraphStyle(
        "small_italic", fontName="Georgia-Italic", fontSize=7.5, leading=11,
        textColor=GRAY, spaceAfter=1*mm
    )
    s["table_header"] = ParagraphStyle(
        "table_header", fontName="Arial-Bold", fontSize=8, leading=11,
        textColor=WHITE
    )
    s["table_cell"] = ParagraphStyle(
        "table_cell", fontName="Arial", fontSize=8, leading=11.5,
        textColor=CHARCOAL
    )
    s["table_cell_bold"] = ParagraphStyle(
        "table_cell_bold", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=CHARCOAL
    )
    s["alert_cell"] = ParagraphStyle(
        "alert_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=RED_SOFT
    )
    s["warn_cell"] = ParagraphStyle(
        "warn_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=AMBER
    )
    s["ok_cell"] = ParagraphStyle(
        "ok_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=GREEN
    )
    s["bullet"] = ParagraphStyle(
        "bullet", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, leftIndent=12, spaceAfter=1.5*mm,
        bulletIndent=0, bulletFontSize=9
    )
    s["dose_label"] = ParagraphStyle(
        "dose_label", fontName="Arial-Bold", fontSize=7.5, leading=10,
        textColor=SAGE
    )
    s["dose_value"] = ParagraphStyle(
        "dose_value", fontName="Georgia-Bold", fontSize=11, leading=15,
        textColor=CHARCOAL
    )
    s["dose_note"] = ParagraphStyle(
        "dose_note", fontName="Arial", fontSize=8, leading=12,
        textColor=GRAY
    )
    s["footer"] = ParagraphStyle(
        "footer", fontName="Arial", fontSize=7, leading=10,
        textColor=GRAY, alignment=TA_CENTER
    )
    s["center"] = ParagraphStyle(
        "center", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, alignment=TA_CENTER, spaceAfter=2*mm
    )
    return s

ST = make_styles()

# ─── Header/Footer on every page ───
class LetterheadCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self.pages)
        for i, page in enumerate(self.pages):
            self.__dict__.update(page)
            self._draw_letterhead(i + 1, num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_letterhead(self, page_num, total_pages):
        c = self
        w, h = A4

        # ─── TOP BAR (sage green) ───
        c.setFillColor(SAGE)
        c.rect(0, h - 28*mm, w, 28*mm, fill=1, stroke=0)

        # Decorative copper accent line
        c.setStrokeColor(COPPER)
        c.setLineWidth(1.5)
        c.line(0, h - 28*mm, w, h - 28*mm)

        # Name / Title in header
        c.setFillColor(WHITE)
        c.setFont("Georgia-Bold", 14)
        c.drawString(20*mm, h - 12*mm, "Залина Бигулова")

        c.setFont("Georgia-Italic", 9)
        c.drawString(20*mm, h - 18*mm, "Нутрициолог-натуропат  |  Антиэйдж менеджмент  |  Спортивная диетология")

        c.setFont("Arial", 7.5)
        c.drawString(20*mm, h - 23.5*mm, "Эксперт МУНН  |  Член Национального общества нутрициологии")

        # Contact info (right side)
        c.setFont("Arial", 7.5)
        c.setFillColor(HexColor("#D4E8D7"))
        c.drawRightString(w - 20*mm, h - 12*mm, "+7 918 822 7716  |  b-zalina@mail.ru")
        c.drawRightString(w - 20*mm, h - 18*mm, "taplink.cc/zalina_bigulova")

        # ─── FOOTER ───
        # Thin sage line
        c.setStrokeColor(SAGE_LIGHT)
        c.setLineWidth(0.5)
        c.line(20*mm, 14*mm, w - 20*mm, 14*mm)

        c.setFillColor(GRAY)
        c.setFont("Arial", 6.5)
        c.drawString(20*mm, 10*mm,
            "Данные рекомендации носят информационный характер. Перед приёмом добавок проконсультируйтесь с врачом.")
        c.drawRightString(w - 20*mm, 10*mm, f"Стр. {page_num} из {total_pages}")


# ─── Helper functions ───
def colored_line():
    return HRFlowable(
        width="100%", thickness=0.5, color=SAGE_LIGHT,
        spaceBefore=2*mm, spaceAfter=2*mm
    )

def copper_line():
    return HRFlowable(
        width="40%", thickness=1, color=COPPER,
        spaceBefore=3*mm, spaceAfter=3*mm
    )

def section_header(number, title):
    """Creates a section header with number and title"""
    return [
        Spacer(1, 3*mm),
        Paragraph(f'<font color="{COPPER.hexval()}" size="8">{number}</font>', ST["center"]),
        Paragraph(title, ST["h1"]),
        colored_line(),
    ]

def bullet(text):
    return Paragraph(f'<bullet>&bull;</bullet> {text}', ST["bullet"])

def dose_box(label, value, note=""):
    """Creates a styled dose recommendation box"""
    data = [[
        Paragraph(f'<font size="7" color="{SAGE.hexval()}"><b>{label}</b></font>', ST["dose_label"]),
    ], [
        Paragraph(f'<b>{value}</b>', ST["dose_value"]),
    ]]
    if note:
        data.append([Paragraph(note, ST["dose_note"])])

    t = Table(data, colWidths=[155*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), SAGE_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("LINEBEFOREDECL", (0, 0), (0, -1), 3, SAGE),
    ]))
    return t

def status_tag(text, status="ok"):
    colors = {
        "ok": (GREEN, GREEN_BG),
        "warn": (AMBER, AMBER_BG),
        "alert": (RED_SOFT, RED_BG),
    }
    fg, bg = colors.get(status, colors["ok"])
    style = {"ok": ST["ok_cell"], "warn": ST["warn_cell"], "alert": ST["alert_cell"]}[status]
    return Paragraph(text, style)

def analysis_table(title, rows):
    """
    rows: list of (name, result, norm, status_text, status_type)
    """
    header = [
        Paragraph("Показатель", ST["table_header"]),
        Paragraph("Результат", ST["table_header"]),
        Paragraph("Норма", ST["table_header"]),
        Paragraph("Статус", ST["table_header"]),
    ]
    data = [header]
    for name, result, norm, st_text, st_type in rows:
        data.append([
            Paragraph(f'<b>{name}</b>', ST["table_cell_bold"]),
            Paragraph(result, ST["table_cell"]),
            Paragraph(norm, ST["table_cell"]),
            status_tag(st_text, st_type),
        ])

    col_w = [60*mm, 35*mm, 35*mm, 25*mm]
    t = Table(data, colWidths=col_w, repeatRows=1)

    style_cmds = [
        # Header
        ("BACKGROUND", (0, 0), (-1, 0), SAGE),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        # Body
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        # Grid
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT),
        ("LINEBELOW", (0, 0), (-1, 0), 1, SAGE),
        # Padding
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]
    # Alternate row colors
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))

    t.setStyle(TableStyle(style_cmds))
    return KeepTogether([
        Paragraph(title, ST["h2"]),
        Spacer(1, 1*mm),
        t,
        Spacer(1, 3*mm),
    ])


def protocol_table():
    header = [
        Paragraph("Добавка", ST["table_header"]),
        Paragraph("Дозировка", ST["table_header"]),
        Paragraph("Когда", ST["table_header"]),
        Paragraph("Курс", ST["table_header"]),
    ]
    rows_data = [
        ("Витамин D3", "5 000 МЕ", "Утро, с жирной пищей", "8-12 нед., затем 2 000 МЕ"),
        ("Витамин K2 (МК-7)", "200 мкг", "Утро, вместе с D3", "Постоянно с D3"),
        ("Магний (цитрат)", "300-400 мг", "Вечер, перед сном", "Постоянно"),
        ("Омега-3 (EPA+DHA)", "2 000-3 000 мг", "С едой (обед/ужин)", "Постоянно"),
        ("Железо (бисглицинат)", "25-36 мг", "Утро натощак + вит. C", "2-3 мес., контроль"),
        ("Цинк (пиколинат)", "15-25 мг", "Вечер, отдельно от Fe", "2-3 мес."),
        ("Йод (калия йодид)", "150-200 мкг", "Утро", "3 мес., контроль"),
        ("Семена льна (молотые)", "1-2 ст. л.", "С едой (каша, салат)", "Постоянно"),
    ]
    data = [header]
    for name, dose, when, course in rows_data:
        data.append([
            Paragraph(f'<b>{name}</b>', ST["table_cell_bold"]),
            Paragraph(dose, ST["table_cell"]),
            Paragraph(when, ST["table_cell"]),
            Paragraph(course, ST["table_cell"]),
        ])

    col_w = [42*mm, 32*mm, 42*mm, 39*mm]
    t = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), SAGE),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT),
        ("LINEBELOW", (0, 0), (-1, 0), 1, SAGE),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))
    t.setStyle(TableStyle(style_cmds))
    return t


# ─── Build the document ───
def build_pdf(output_path):
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=34*mm, bottomMargin=20*mm,
        leftMargin=20*mm, rightMargin=20*mm,
        title="Рекомендации нутрициолога — Залина Бигулова",
        author="Залина Бигулова",
        subject="Персональные рекомендации по результатам анализов",
    )

    story = []

    # ═══════════════════════════════════════════
    # PAGE 1: TITLE + CLIENT INFO + OVERVIEW
    # ═══════════════════════════════════════════
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph("Персональные рекомендации", ST["title"]))
    story.append(Paragraph("по результатам лабораторных исследований", ST["subtitle"]))
    story.append(copper_line())

    # Client info box
    client_data = [
        [Paragraph('<b>Дата анализов:</b>', ST["table_cell_bold"]),
         Paragraph('14.02.2025, лаборатория HELIX, Владикавказ', ST["table_cell"])],
        [Paragraph('<b>Дата рекомендаций:</b>', ST["table_cell_bold"]),
         Paragraph('Март 2025', ST["table_cell"])],
        [Paragraph('<b>Специалист:</b>', ST["table_cell_bold"]),
         Paragraph('Бигулова Залина Константиновна, нутрициолог-натуропат', ST["table_cell"])],
    ]
    client_t = Table(client_data, colWidths=[45*mm, 110*mm])
    client_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CREAM_DARK),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(client_t)
    story.append(Spacer(1, 5*mm))

    # Overview summary cards as table
    story.append(Paragraph("Ключевые находки", ST["h2"]))
    story.append(Spacer(1, 1*mm))

    overview_data = [
        [Paragraph('<b>Витамин D</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">25.97 нг/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 30-100', ST["small"]),
         status_tag("ДЕФИЦИТ", "alert")],
        [Paragraph('<b>Холестерин общий</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">5.92 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;5.2', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Холестерин не-ЛПВП</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">3.97 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;3.4', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>ФСГ</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">27.80 мМЕ/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 3.5-12.5', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>ЛГ</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">15.70 мМЕ/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 2.4-12.6', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Ферритин</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">38.4 мкг/л</font></b>', ST["table_cell"]),
         Paragraph('оптимум 50-80', ST["small"]),
         status_tag("Субоптимально", "warn")],
        [Paragraph('<b>Йод</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">32.47 мкг/л</font></b>', ST["table_cell"]),
         Paragraph('норма 30-60', ST["small"]),
         status_tag("Нижняя граница", "warn")],
        [Paragraph('<b>Цинк</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">739.7 мкг/л</font></b>', ST["table_cell"]),
         Paragraph('норма 700-1140', ST["small"]),
         status_tag("Нижняя граница", "warn")],
    ]
    overview_t = Table(overview_data, colWidths=[45*mm, 35*mm, 35*mm, 40*mm])
    overview_t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 4), RED_BG),
        ("BACKGROUND", (0, 5), (-1, 7), AMBER_BG),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]))
    story.append(overview_t)
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        '<b>Также в норме (21 из 29 показателей):</b> эстрадиол, пролактин, кортизол, ТТГ, Т4, Т3, '
        'анти-ТПО, антитела к тиреоглобулину, ЛПНП, ЛПВП (отличный!), ЛПОНП, триглицериды, '
        'коэф. атерогенности (2.04 - отлично), ДЭА-SO4, витамин B9, витамин B12, селен, медь, '
        'гомоцистеин, С-реактивный белок.',
        ST["small"]
    ))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 2: FULL ANALYSIS TABLES
    # ═══════════════════════════════════════════
    story.extend(section_header("01", "Результаты анализов — полная расшифровка"))

    # Hormones
    story.append(analysis_table("Гормональная панель", [
        ("ФСГ — фолликулостимулирующий\nгормон", "27.80 мМЕ/мл", "3.50-12.50", "Повышен", "alert"),
        ("ЛГ — лютеинизирующий гормон", "15.70 мМЕ/мл", "2.40-12.60", "Повышен", "alert"),
        ("Эстрадиол", "35.80 пг/мл", "12.40-233.00", "Норма", "ok"),
        ("Пролактин", "272.00 мкМЕ/мл", "102.00-496.00", "Норма", "ok"),
        ("Кортизол (утро)", "343.0 нмоль/л", "166-507", "Норма", "ok"),
        ("ДЭА-SO4", "76.8 мкг/дл", "35.4-256.0", "Ниже оптимума", "warn"),
    ]))

    # Thyroid
    story.append(analysis_table("Щитовидная железа", [
        ("ТТГ — тиреотропный гормон", "0.978 мкМЕ/мл", "0.270-4.200", "Норма", "ok"),
        ("Т4 свободный", "17.90 пмоль/л", "10.80-22.00", "Норма", "ok"),
        ("Т3 свободный", "4.17 пмоль/л", "3.10-6.80", "Норма", "ok"),
        ("Анти-ТПО", "<9.00 МЕ/мл", "0.00-34.00", "Норма", "ok"),
        ("Антитела к тиреоглобулину", "19.90 МЕ/мл", "0.00-115.00", "Норма", "ok"),
    ]))

    # Lipids
    story.append(analysis_table("Липидный профиль", [
        ("Холестерин общий", "5.92 ммоль/л", "<5.20", "Повышен", "alert"),
        ("Холестерин не-ЛПВП", "3.97 ммоль/л", "<3.40", "Повышен", "alert"),
        ("ЛПНП", "1.86 ммоль/л", "<3.00", "Норма", "ok"),
        ("ЛПВП (\"хороший\")", "1.95 ммоль/л", ">1.20", "Отличный", "ok"),
        ("ЛПОНП", "0.43 ммоль/л", "<0.80", "Норма", "ok"),
        ("Триглицериды", "0.94 ммоль/л", "<1.70", "Норма", "ok"),
        ("Коэф. атерогенности", "2.04", "2.20-3.50", "Отлично", "ok"),
    ]))

    # Vitamins & Minerals
    story.append(analysis_table("Витамины и минералы", [
        ("Витамин D (25-гидрокси)", "25.97 нг/мл", "30.00-100.00", "Дефицит", "alert"),
        ("Ферритин", "38.4 мкг/л", "10.0-120.0", "Субоптимально", "warn"),
        ("Железо сыворотки", "14.75 мкмоль/л", "6.60-26.00", "Ниже оптимума", "warn"),
        ("Витамин B9 (фолат)", "14.50 нмоль/л", "7.00-39.70", "Норма", "ok"),
        ("Витамин B12 активный", "129.00 пмоль/л", "25.10-165.00", "Норма", "ok"),
        ("Йод в плазме", "32.47 мкг/л", "30.00-60.00", "Нижняя граница", "warn"),
        ("Цинк в сыворотке", "739.70 мкг/л", "700.00-1140.00", "Нижняя граница", "warn"),
        ("Селен в плазме", "122.00 мкг/л", "23.00-190.00", "Норма", "ok"),
        ("Медь в плазме", "1109.00 мкг/л", "575.00-1175.00", "Норма", "ok"),
    ]))

    # Other markers
    story.append(analysis_table("Прочие маркеры", [
        ("Гомоцистеин", "6.90 мкмоль/л", "0.00-15.00", "Норма", "ok"),
        ("С-реактивный белок", "0.65 мг/л", "0.00-5.00", "Норма", "ok"),
    ]))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 4: VITAMIN D
    # ═══════════════════════════════════════════
    story.extend(section_header("02", "Восполнение дефицитов"))

    story.append(Paragraph("Витамин D — устранить дефицит", ST["h2"]))
    story.append(Paragraph(
        '<font color="#DC4E42"><b>25.97 нг/мл</b></font> при норме 30-100. Цель: 50-60 нг/мл.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>Почему важно:</b> витамин D — прогормон, регулирующий усвоение кальция, иммунитет, '
        'настроение, обмен веществ и здоровье костей. В перименопаузе дефицит ускоряет потерю '
        'костной массы и повышает риск остеопороза. Также влияет на инсулиночувствительность и вес.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Витамин D3 (холекальциферол) — 5 000 МЕ/день",
        "Курс 8-12 недель, затем контроль 25(OH)D. После достижения 50+ нг/мл — поддерживающая доза 2 000 МЕ/день."
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph('<b>Обязательные кофакторы:</b>', ST["body_bold"]))
    story.append(bullet('<b>Витамин K2 (МК-7)</b> — 100-200 мкг/день. Направляет кальций в кости, а не в сосуды'))
    story.append(bullet('<b>Магний (цитрат или глицинат)</b> — 300-400 мг/день. Нужен для активации витамина D'))
    story.append(Spacer(1, 1*mm))

    story.append(Paragraph(
        '<b>Продукты:</b> жирная рыба (лосось, сельдь, скумбрия), яичный желток, печень трески. '
        'Принимать D3 с жирной пищей в первой половине дня.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    # FERRITIN
    story.append(Paragraph("Ферритин и железо — довести до оптимума", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>Ферритин 38.4 мкг/л</b></font> — функциональный оптимум 50-80. '
        'При уровне ниже 50 мкг/л часто наблюдаются: хроническая усталость, выпадение волос, '
        'сухость кожи, ломкость ногтей. Восполнение до 50+ снижает усталость на 48% за 12 недель.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Хелатное железо (бисглицинат) — 25-36 мг/день",
        "Натощак или через 2 часа после еды + витамин С (80-100 мг). Курс 2-3 мес. с контролем ферритина."
    ))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        '<b>Источники:</b> говядина, печень, морепродукты, шпинат, чечевица, гречка, гранат. '
        '<b>Не сочетать</b> с кальцием, цинком, чаем и кофе — разнести на 2-3 часа.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    # IODINE
    story.append(Paragraph("Йод — усилить поступление", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>32.47 мкг/л</b></font> — нижняя граница нормы (30-60). '
        'Йод — строительный материал для гормонов щитовидной железы. Щитовидка сейчас работает '
        'нормально, но пограничный йод при длительном дефиците может замедлить метаболизм. '
        'Хороший уровень селена (122 мкг/л) делает приём йода безопасным.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Калия йодид — 150-200 мкг/день",
        "Источники: морская капуста, морская рыба, креветки, йодированная соль."
    ))

    story.append(Spacer(1, 4*mm))

    # ZINC
    story.append(Paragraph("Цинк — поднять уровень", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>739.7 мкг/л</b></font> — нижняя граница (700-1140). '
        'Цинк участвует в 300+ ферментных реакциях: иммунитет, кожа, волосы, синтез гормонов, '
        'инсулиночувствительность. При пограничном цинке и высокой меди (1109) есть дисбаланс.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Цинк (пиколинат или бисглицинат) — 15-25 мг/день",
        "Вечером, отдельно от железа и кальция. Курс 2-3 месяца. Источники: говядина, тыквенные семечки, устрицы."
    ))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 5: CHOLESTEROL
    # ═══════════════════════════════════════════
    story.extend(section_header("03", "Липидный профиль — снижение холестерина"))

    story.append(Paragraph(
        'Общий холестерин 5.92 и не-ЛПВП 3.97 повышены. Лаборатория HELIX установила '
        'гиперлипопротеинемию IIа типа. <b>Хорошая новость:</b> ЛПВП 1.95 — отличный, '
        'коэффициент атерогенности 2.04 — ниже нормы, воспаления нет (СРБ 0.65).',
        ST["body"]
    ))

    story.append(Paragraph("Почему холестерин повысился", ST["h3"]))
    story.append(Paragraph(
        'В перименопаузе (ФСГ 27.8, ЛГ 15.7) снижается защитное действие эстрогена на липидный '
        'обмен. Эстроген помогал удерживать ЛПНП низким и ЛПВП высоким. При его колебаниях общий '
        'холестерин растёт на 10-15%. У 77% женщин 45-64 лет наблюдается повышенный холестерин. '
        'Это физиологически объяснимо, но требует коррекции.',
        ST["body"]
    ))

    story.append(Paragraph("Как снизить без лекарств", ST["h3"]))

    story.append(Paragraph('<b>1. Омега-3 жирные кислоты (EPA + DHA):</b>', ST["body_bold"]))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Омега-3 — 2 000-3 000 мг EPA+DHA в день",
        "Снижает триглицериды на 20-30%, улучшает функцию ЛПВП, противовоспалительное действие. С едой."
    ))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph('<b>2. Растворимая клетчатка</b> — 10-25 г/день. Связывает холестерин и выводит:', ST["body_bold"]))
    story.append(bullet('Овсянка (бета-глюканы), яблоки (пектин), семена льна'))
    story.append(bullet('Чечевица, фасоль, псиллиум (подорожник), авокадо'))

    story.append(Paragraph('<b>3. Растительные стеролы</b> — 2 г/день. Снижают ЛПНП на 6-15%.', ST["body_bold"]))

    story.append(Paragraph('<b>4. Полезные жиры — заменить, а не убрать:</b>', ST["body_bold"]))
    story.append(bullet('Оливковое масло extra virgin — 2-3 ст. ложки в день'))
    story.append(bullet('Орехи (грецкие, миндаль) — 30 г/день'))
    story.append(bullet('Авокадо — половинка в день'))

    story.append(Paragraph('<b>5. Исключить или ограничить:</b>', ST["body_bold"]))
    story.append(bullet('Трансжиры (маргарин, фаст-фуд, промышленная выпечка)'))
    story.append(bullet('Рафинированный сахар и белую муку'))
    story.append(bullet('Избыток насыщенных жиров (жирное мясо, сливочное масло >10 г/день)'))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 6: HORMONES + PERIMENOPAUSE
    # ═══════════════════════════════════════════
    story.extend(section_header("04", "Гормональный фон — перименопауза"))

    story.append(Paragraph(
        'ФСГ 27.8 и ЛГ 15.7 подтверждают начало перименопаузального перехода. Это не болезнь, а новый этап.',
        ST["body"]
    ))

    story.append(Paragraph("Что это значит", ST["h3"]))
    story.append(Paragraph(
        '<b>ФСГ (27.8 при норме до 12.5)</b> — гипофиз «стимулирует» яичники работать активнее, но они '
        'постепенно снижают выработку эстрогена. Чем выше ФСГ — тем ближе менопауза.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>ЛГ (15.7 при норме до 12.6)</b> — повышается параллельно с ФСГ, подтверждая переходный период.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>Хорошие новости:</b> эстрадиол пока в рабочем диапазоне (35.8), щитовидка работает прекрасно, '
        'кортизол в норме, воспаления нет.',
        ST["body"]
    ))

    story.append(Paragraph("Нутритивная поддержка", ST["h3"]))
    story.append(Paragraph('<b>Фитоэстрогены</b> — мягко поддерживают гормональный баланс:', ST["body_bold"]))
    story.append(bullet('<b>Семена льна</b> — 1-2 ст. ложки молотых в день (лигнаны)'))
    story.append(bullet('<b>Изофлавоны сои</b> — тофу, эдамамэ, мисо 2-3 раза в неделю'))

    story.append(Paragraph('<b>Адаптогены</b> для надпочечников (ДЭА-SO4 на нижней границе):', ST["body_bold"]))
    story.append(bullet('<b>Ашваганда (KSM-66)</b> — 300-600 мг/день (снижает кортизол, поддерживает ДЭА)'))
    story.append(bullet('<b>Родиола розовая</b> — 200-400 мг/день (энергия и стрессоустойчивость)'))

    story.append(Paragraph('<b>Кальций + магний</b> — профилактика остеопороза:', ST["body_bold"]))
    story.append(bullet('Кальций из продуктов — 1000-1200 мг/день (кунжут, зелень, молочные)'))
    story.append(bullet('Магний — 300-400 мг/день (уже входит в протокол с витамином D)'))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 7: WEIGHT LOSS
    # ═══════════════════════════════════════════
    story.extend(section_header("05", "Снижение веса — системный подход"))

    story.append(Paragraph(
        'В перименопаузе метаболизм замедляется, меняется распределение жира, '
        'растёт инсулинорезистентность. Но снизить вес реально при системном подходе.',
        ST["body"]
    ))

    story.append(Paragraph("1. Белок — основа каждого приёма пищи", ST["h3"]))
    story.append(bullet('<b>1.2-1.5 г белка на 1 кг веса</b> — минимум для сохранения мышц'))
    story.append(bullet('В каждом приёме: мясо, рыба, яйца, творог или бобовые'))
    story.append(bullet('Белок усиливает термогенез и даёт сытость на 3-4 часа'))

    story.append(Paragraph("2. Средиземноморский стиль питания", ST["h3"]))
    story.append(bullet('<b>Овощи</b> — 400-500 г в день (половина тарелки)'))
    story.append(bullet('<b>Жирная рыба</b> — 3-4 раза в неделю'))
    story.append(bullet('<b>Оливковое масло</b> — основной жир'))
    story.append(bullet('<b>Цельнозерновые</b> — бурый рис, гречка, овсянка'))
    story.append(bullet('<b>Ограничить:</b> сахар, белый хлеб, колбасы, полуфабрикаты'))

    story.append(Paragraph("3. Режим питания — 3 приёма, без перекусов", ST["h3"]))
    story.append(bullet('<b>Завтрак</b> (до 10:00) — белок + жиры + клетчатка'))
    story.append(bullet('<b>Обед</b> (13:00-14:00) — самый объёмный приём'))
    story.append(bullet('<b>Ужин</b> (до 19:00) — лёгкий, белок + овощи'))
    story.append(bullet('Пищевое окно 10-12 часов. <b>Вода</b> — 30 мл на 1 кг веса в день'))

    story.append(Paragraph("4. Физическая активность", ST["h3"]))
    story.append(bullet('<b>Силовые тренировки</b> — 2-3 раза/нед (сохранение мышц = быстрый метаболизм)'))
    story.append(bullet('<b>Ходьба</b> — 8 000-10 000 шагов в день'))
    story.append(bullet('<b>Кардио</b> — 2-3 раза/нед по 30-40 мин (плавание, велосипед)'))
    story.append(bullet('Избегать чрезмерных нагрузок — повышают кортизол'))

    story.append(Paragraph("5. Сон и стресс — скрытые факторы веса", ST["h3"]))
    story.append(bullet('<b>Сон 7-8 часов</b> — недосып повышает грелин и снижает лептин'))
    story.append(bullet('<b>Магний перед сном</b> — улучшает сон и снижает тревожность'))
    story.append(bullet('Практики: дыхательные техники, прогулки на природе, йога'))

    story.append(Paragraph("6. Безопасный темп", ST["h3"]))
    story.append(bullet('<b>2-4 кг в месяц</b> — физиологичный темп. Дефицит не более 300-500 ккал'))
    story.append(bullet('Контролировать <b>объёмы</b>, а не только вес. Не голодать!'))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 8: PROTOCOL TABLE + SOURCES
    # ═══════════════════════════════════════════
    story.extend(section_header("06", "Сводный протокол приёма"))

    story.append(Paragraph(
        'Что принимать, когда и в какой дозе. Разнесено по времени для оптимального усвоения.',
        ST["body"]
    ))
    story.append(Spacer(1, 2*mm))
    story.append(protocol_table())

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Контрольные анализы", ST["h3"]))
    story.append(Paragraph(
        'Через <b>3 месяца</b> повторить: витамин D (25-OH), ферритин, липидный профиль, '
        'йод, цинк. Рекомендуется консультация терапевта/кардиолога (по заключению лаборатории '
        'HELIX) и гинеколога-эндокринолога для обсуждения менопаузальной гормональной терапии.',
        ST["body"]
    ))

    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Источники", ST["h2"]))
    sources = [
        "Российская ассоциация эндокринологов — клинические рекомендации по витамину D",
        "AHA (American Heart Association) — Omega-3 and Dyslipidemia Meta-Analysis, 2023",
        "NIH — Omega-3 Fatty Acids Fact Sheet; Vitamin D Health Professional Fact Sheet",
        "Mayo Clinic Proceedings — Omega-3 Dosage and Cardiovascular Outcomes",
        "NAMS (North American Menopause Society) — Menopause Management Guidelines",
        "PMC/NIH — N-3 Fatty Acids and Cardiovascular Health; Iron Deficiency Reviews",
        "NICE Guidelines — Menopause: Diagnosis and Management (NG23)",
        "Docma.ru — Витамин D рекомендации 2025",
        "Гемотест, СМ-Клиника, ЦНМТ, СИНЭКС — справочные материалы",
        "НМИЦ ТПМ — Гормональные сдвиги при менопаузе и сердечно-сосудистая система",
        "Frontiers in Endocrinology — Gut Microbiota and Menopausal Health, 2025",
        "Stanford Medicine — Fermented Food Study, 2021",
        "Examine.com — Evidence-based Supplement Reviews",
        "PubMed — Систематические обзоры по адаптогенам, фитоэстрогенам, клетчатке",
        "WHO — Physical Activity Guidelines, 2020",
    ]
    for i, src in enumerate(sources, 1):
        story.append(Paragraph(f'{i}. {src}', ST["small"]))

    story.append(Spacer(1, 4*mm))
    story.append(colored_line())
    story.append(Spacer(1, 2*mm))

    # Signature block
    sig_data = [
        [Paragraph('<b>Нутрициолог:</b>', ST["body_bold"]),
         Paragraph('', ST["body"]),
         Paragraph('Бигулова З.К.', ST["body_bold"])],
        [Paragraph('', ST["small"]),
         Paragraph('', ST["small"]),
         Paragraph('подпись', ST["small_italic"])],
    ]
    sig_t = Table(sig_data, colWidths=[40*mm, 65*mm, 50*mm])
    sig_t.setStyle(TableStyle([
        ("LINEBELOW", (1, 0), (1, 0), 0.5, GRAY),
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(sig_t)

    # Build
    doc.build(story, canvasmaker=LetterheadCanvas)
    print(f"PDF saved to: {output_path}")


if __name__ == "__main__":
    output = "/Users/serzhbigulov/Documents/Zalina/Рекомендации_нутрициолога_Бигулова.pdf"
    build_pdf(output)
