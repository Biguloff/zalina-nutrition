#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate professional nutritionist recommendation PDF
for Алехина Стояна Игоревна
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
PURPLE = HexColor("#7C3AED")
PURPLE_BG = HexColor("#F5F3FF")

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
    s["info_cell"] = ParagraphStyle(
        "info_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=PURPLE
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
    s["alert_box"] = ParagraphStyle(
        "alert_box", fontName="Arial", fontSize=9, leading=13.5,
        textColor=RED_SOFT, alignment=TA_LEFT, spaceAfter=2*mm
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
    return [
        Spacer(1, 3*mm),
        Paragraph(f'<font color="{COPPER.hexval()}" size="8">{number}</font>', ST["center"]),
        Paragraph(title, ST["h1"]),
        colored_line(),
    ]

def bullet(text):
    return Paragraph(f'<bullet>&bull;</bullet> {text}', ST["bullet"])

def dose_box(label, value, note=""):
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

def alert_box(text):
    data = [[Paragraph(text, ST["alert_box"])]]
    t = Table(data, colWidths=[155*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), RED_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("LINEBEFOREDECL", (0, 0), (0, -1), 3, RED_SOFT),
    ]))
    return t

def status_tag(text, status="ok"):
    colors = {
        "ok": (GREEN, GREEN_BG),
        "warn": (AMBER, AMBER_BG),
        "alert": (RED_SOFT, RED_BG),
        "info": (PURPLE, PURPLE_BG),
    }
    fg, bg = colors.get(status, colors["ok"])
    style = {
        "ok": ST["ok_cell"],
        "warn": ST["warn_cell"],
        "alert": ST["alert_cell"],
        "info": ST["info_cell"],
    }[status]
    return Paragraph(text, style)

def analysis_table(title, rows):
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
        ("BACKGROUND", (0, 0), (-1, 0), SAGE),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT),
        ("LINEBELOW", (0, 0), (-1, 0), 1, SAGE),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]
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
        ("Витамин D3", "4 000 МЕ", "Утро, с жирной пищей", "8-12 нед., затем 2 000 МЕ"),
        ("Витамин K2 (МК-7)", "100-200 мкг", "Утро, вместе с D3", "Постоянно с D3"),
        ("Магний (цитрат/глицинат)", "300-400 мг", "Вечер, перед сном", "Постоянно"),
        ("Железо (бисглицинат)", "25 мг", "Утро натощак + вит. C", "2-3 мес., контроль"),
        ("Цинк (пиколинат)", "15-25 мг", "Вечер, отдельно от Fe", "2-3 мес."),
        ("Витамин B9 (метилфолат)", "400-800 мкг", "Утро, с едой", "3 мес., контроль"),
        ("Омега-3 (EPA+DHA)", "1 000-2 000 мг", "С едой (обед/ужин)", "Постоянно"),
    ]
    data = [header]
    for name, dose, when, course in rows_data:
        data.append([
            Paragraph(f'<b>{name}</b>', ST["table_cell_bold"]),
            Paragraph(dose, ST["table_cell"]),
            Paragraph(when, ST["table_cell"]),
            Paragraph(course, ST["table_cell"]),
        ])

    col_w = [45*mm, 30*mm, 42*mm, 38*mm]
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
        title="Рекомендации нутрициолога — Алехина Стояна Игоревна",
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
        [Paragraph('<b>Пациент:</b>', ST["table_cell_bold"]),
         Paragraph('Алехина Стояна Игоревна, 19 лет (04.07.2006)', ST["table_cell"])],
        [Paragraph('<b>Дата анализов:</b>', ST["table_cell_bold"]),
         Paragraph('15.03.2026, лаборатория HELIX', ST["table_cell"])],
        [Paragraph('<b>Дата рекомендаций:</b>', ST["table_cell_bold"]),
         Paragraph('Апрель 2026', ST["table_cell"])],
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

    # ALERT BOX — prolactin
    story.append(alert_box(
        '<b>ВАЖНО: Пролактин значительно повышен (960 мкМЕ/мл при норме до 496). '
        'Необходима срочная консультация эндокринолога и МРТ гипофиза с контрастом.</b>'
    ))
    story.append(Spacer(1, 3*mm))

    # Overview summary cards
    story.append(Paragraph("Ключевые находки", ST["h2"]))
    story.append(Spacer(1, 1*mm))

    overview_data = [
        [Paragraph('<b>Пролактин</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">960.10 мкМЕ/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 102-496', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Пролактин мономерн.</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">678.40 мкМЕ/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 75-381', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Витамин D</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">20.05 нг/мл</font></b>', ST["table_cell"]),
         Paragraph('норма 30-100', ST["small"]),
         status_tag("НЕДОСТАТОЧН.", "alert")],
        [Paragraph('<b>Фибриноген</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">1.98 г/л</font></b>', ST["table_cell"]),
         Paragraph('норма 2-3.9', ST["small"]),
         status_tag("Снижен", "warn")],
        [Paragraph('<b>Фосфор</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">1.55 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма 0.81-1.45', ST["small"]),
         status_tag("Повышен", "warn")],
        [Paragraph('<b>Ферритин</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">28.3 мкг/л</font></b>', ST["table_cell"]),
         Paragraph('оптимум 40-80', ST["small"]),
         status_tag("Субоптимально", "warn")],
        [Paragraph('<b>Цинк</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">721.20 мкг/л</font></b>', ST["table_cell"]),
         Paragraph('норма 700-1140', ST["small"]),
         status_tag("Нижняя граница", "warn")],
        [Paragraph('<b>Витамин B9</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">9.42 нмоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма 7.0-39.7', ST["small"]),
         status_tag("Нижняя граница", "warn")],
        [Paragraph('<b>Ген LCT (CC)</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#7C3AED">Генотип CC</font></b>', ST["table_cell"]),
         Paragraph('', ST["small"]),
         status_tag("Непереносим.", "info")],
    ]
    overview_t = Table(overview_data, colWidths=[45*mm, 35*mm, 35*mm, 40*mm])
    overview_t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 2), RED_BG),
        ("BACKGROUND", (0, 3), (-1, 7), AMBER_BG),
        ("BACKGROUND", (0, 8), (-1, 8), PURPLE_BG),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]))
    story.append(overview_t)
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph(
        '<b>В норме:</b> липидный профиль (идеальный! коэф. атерогенности 1.08), '
        'щитовидная железа (ТТГ, анти-ТПО, антитела к тиреоглобулину), углеводный обмен '
        '(глюкоза, инсулин, HOMA-IR), печёночные пробы (АЛТ, АСТ, билирубин), '
        'почки (креатинин, СКФ, мочевина), гомоцистеин, СРБ, витамин B12, селен, медь, '
        'кальций, калий, коэнзим Q10, коагулограмма (кроме фибриногена), ОАК.',
        ST["small"]
    ))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 2-3: FULL ANALYSIS TABLES
    # ═══════════════════════════════════════════
    story.extend(section_header("01", "Результаты анализов — полная расшифровка"))

    # Hormones
    story.append(analysis_table("Гормональная панель", [
        ("Пролактин", "960.10 мкМЕ/мл", "102-496", "Повышен", "alert"),
        ("Пролактин мономерный\n(пост-ПЭГ)", "678.40 мкМЕ/мл", "75-381", "Повышен", "alert"),
        ("Макропролактин (пост-ПЭГ)%", "70.66 %", ">60", "Норма", "ok"),
        ("Инсулин", "7.60 мкЕд/мл", "2.6-24.9", "Норма", "ok"),
        ("HOMA-IR", "1.51", "0-2.7", "Норма", "ok"),
        ("Кальцитонин", "1.27 пг/мл", "0-6.40", "Норма", "ok"),
        ("Паратиреоидный гормон", "32.70 пг/мл", "12-88", "Норма", "ok"),
    ]))

    # Thyroid
    story.append(analysis_table("Щитовидная железа", [
        ("ТТГ", "2.250 мкМЕ/мл", "0.51-4.30", "Норма", "ok"),
        ("Анти-ТПО", "7.52 МЕ/мл", "0-34", "Норма", "ok"),
        ("Антитела к тиреоглобулину", "18.13 МЕ/мл", "0-115", "Норма", "ok"),
    ]))

    # Lipids
    story.append(analysis_table("Липидный профиль", [
        ("Холестерин общий", "4.12 ммоль/л", "<5.2", "Норма", "ok"),
        ("ЛПНП", "2.07 ммоль/л", "<3.0", "Норма", "ok"),
        ("ЛПВП", "1.98 ммоль/л", ">1.2", "Отличный", "ok"),
        ("ЛПОНП", "<0.1 ммоль/л", "<0.8", "Норма", "ok"),
        ("Холестерин не-ЛПВП", "2.14 ммоль/л", "<3.4", "Норма", "ok"),
        ("Триглицериды", "0.58 ммоль/л", "<1.7", "Норма", "ok"),
        ("Коэф. атерогенности", "1.08", "2.2-3.5", "Отлично", "ok"),
    ]))

    # Vitamins & Minerals
    story.append(analysis_table("Витамины и минералы", [
        ("Витамин D (25-гидрокси)", "20.05 нг/мл", "30-100", "Недостаточность", "alert"),
        ("Ферритин", "28.3 мкг/л", "10-120", "Субоптимально", "warn"),
        ("Железо сыворотки", "12.70 мкмоль/л", "6.6-26", "Норма", "ok"),
        ("Цинк в сыворотке", "721.20 мкг/л", "700-1140", "Нижняя граница", "warn"),
        ("Витамин B9 (фолат)", "9.42 нмоль/л", "7.0-39.7", "Нижняя граница", "warn"),
        ("Витамин B12 активный", "68.70 пмоль/л", "25.10-165.0", "Норма", "ok"),
        ("Селен в плазме", "106.96 мкг/л", "23-190", "Норма", "ok"),
        ("Медь в плазме", "769.16 мкг/л", "575-1725", "Норма", "ok"),
        ("Фосфор", "1.55 ммоль/л", "0.81-1.45", "Повышен", "warn"),
        ("Кальций в сыворотке", "2.35 ммоль/л", "2.15-2.50", "Норма", "ok"),
        ("Кальций ионизированный", "1.32 ммоль/л", "1.16-1.32", "Верхняя граница", "warn"),
        ("Калий", "3.97 ммоль/л", "3.5-5.1", "Норма", "ok"),
        ("Коэнзим Q10", "817.00 нг/мл", "433-1532", "Норма", "ok"),
        ("Йод в моче", "258.63 мкг/л", "28-544", "Норма", "ok"),
    ]))

    story.append(PageBreak())

    # Liver, kidney, other
    story.append(analysis_table("Печень и поджелудочная", [
        ("АЛТ", "13.4 Ед/л", "0-33", "Норма", "ok"),
        ("АСТ", "13.1 Ед/л", "0-32", "Норма", "ok"),
        ("Билирубин общий", "6.90 мкмоль/л", "<21", "Норма", "ok"),
        ("Билирубин прямой", "3.20 мкмоль/л", "<5.0", "Норма", "ok"),
        ("Амилаза панкреатическая", "18 Ед/л", "0-53", "Норма", "ok"),
        ("Белок общий", "70.6 г/л", "64-83", "Норма", "ok"),
    ]))

    story.append(analysis_table("Почки и обмен веществ", [
        ("Креатинин", "64.00 мкмоль/л", "44-80", "Норма", "ok"),
        ("СКФ (CKD-EPI)", "120.98 мл/мин", ">60", "Норма", "ok"),
        ("Мочевина", "5.90 ммоль/л", "2.9-7.5", "Норма", "ok"),
        ("Мочевая кислота", "155.00 мкмоль/л", "142.8-339.2", "Норма", "ok"),
        ("Гомоцистеин", "8.50 мкмоль/л", "0-15", "Норма", "ok"),
        ("С-реактивный белок", "<0.6 мг/л", "0-5", "Норма", "ok"),
        ("Глюкоза", "4.47 ммоль/л", "4.1-6.1", "Норма", "ok"),
    ]))

    story.append(analysis_table("Коагулограмма", [
        ("D-димер", "<150 нг/мл", "<243", "Норма", "ok"),
        ("Антитромбин III", "103.00 %", "83-128", "Норма", "ok"),
        ("АЧТВ", "33.1 сек", "25.1-36.5", "Норма", "ok"),
        ("Протромбин (по Квику)", "88.00 %", "70-120", "Норма", "ok"),
        ("МНО", "1.08", "0.8-1.2", "Норма", "ok"),
        ("Фибриноген", "1.98 г/л", "2-3.9", "Снижен", "warn"),
    ]))

    story.append(analysis_table("Общий анализ крови (выборочно)", [
        ("Гемоглобин", "122 г/л", "117-155", "Норма", "ok"),
        ("Эритроциты", "4.11 *10<super>12</super>/л", "3.8-5.1", "Норма", "ok"),
        ("Лейкоциты", "5.68 *10<super>9</super>/л", "4.0-10.0", "Норма", "ok"),
        ("Тромбоциты", "183 *10<super>9</super>/л", "150-400", "Норма", "ok"),
        ("Лимфоциты %", "37.6 %", "19-37", "Чуть повышены", "warn"),
        ("СОЭ", "4 мм/ч", "2-20", "Норма", "ok"),
        ("Гематокрит", "35.8 %", "35-45", "Нижняя граница", "ok"),
    ]))

    story.append(analysis_table("Анализ мочи (отклонения)", [
        ("Белок в моче", "0.16 г/л", "0-0.15", "Чуть повышен", "warn"),
        ("Кристаллы: ураты", "большое кол-во", "отсутствуют", "Отклонение", "alert"),
        ("Прозрачность", "мутная", "прозрачная", "Отклонение", "warn"),
    ]))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PROLACTIN
    # ═══════════════════════════════════════════
    story.extend(section_header("02", "Гиперпролактинемия — требует обследования"))

    story.append(alert_box(
        '<b>Пролактин 960.10 мкМЕ/мл — почти в 2 раза выше верхней границы нормы (496). '
        'Мономерный пролактин (пост-ПЭГ) 678.40 мкМЕ/мл также значительно повышен (норма до 381). '
        'Макропролактин НЕ обнаружен — это ИСТИННАЯ гиперпролактинемия.</b>'
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Что это значит", ST["h3"]))
    story.append(Paragraph(
        'Пролактин — гормон гипофиза. В норме он повышается при беременности и кормлении '
        'грудью. Повышение у молодой девушки вне беременности может указывать на:',
        ST["body"]
    ))
    story.append(bullet('<b>Микроаденому гипофиза (пролактиному)</b> — наиболее частая причина при таком уровне'))
    story.append(bullet('<b>Функциональную гиперпролактинемию</b> — стресс, нарушение сна, приём некоторых лекарств'))
    story.append(bullet('<b>Гипотиреоз</b> — исключён (ТТГ 2.25 — норма)'))

    story.append(Paragraph("Возможные симптомы", ST["h3"]))
    story.append(bullet('Нарушения менструального цикла (задержки, нерегулярность, аменорея)'))
    story.append(bullet('Головные боли, нарушения зрения (при аденоме)'))
    story.append(bullet('Галакторея (выделения из молочных желёз)'))
    story.append(bullet('Акне, выпадение волос'))

    story.append(Paragraph("Обязательные обследования", ST["h3"]))
    story.append(Spacer(1, 1*mm))
    story.append(dose_box(
        "НЕОБХОДИМО СДЕЛАТЬ",
        "1. МРТ гипофиза с контрастом",
        "Для исключения микро-/макроаденомы гипофиза. Направление от эндокринолога."
    ))
    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "КОНСУЛЬТАЦИИ",
        "2. Эндокринолог + гинеколог-эндокринолог",
        "Для определения причины и тактики лечения. Повторный анализ пролактина — натощак, "
        "утром (8-10 ч), в покое, без стресса, исключив за сутки физ. нагрузки и половые контакты."
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        '<b>Нутритивная поддержка при гиперпролактинемии</b> (не заменяет лечение у врача):',
        ST["body_bold"]
    ))
    story.append(bullet('<b>Витамин B6 (пиридоксин)</b> — 50-100 мг/день. Участвует в синтезе дофамина, '
                        'который подавляет выработку пролактина'))
    story.append(bullet('<b>Витекс (Vitex agnus-castus)</b> — 20-40 мг стандартизированного экстракта. '
                        'Обладает дофаминергическим действием, снижает пролактин. '
                        'Только после консультации с эндокринологом!'))
    story.append(bullet('<b>Цинк</b> — 15-25 мг/день. Снижает повышенный пролактин'))
    story.append(bullet('<b>Магний</b> — 300-400 мг/день. Поддерживает нервную систему, снижает стресс'))
    story.append(bullet('Управление стрессом: сон 8-9 часов, снижение нагрузки, прогулки'))

    story.append(Spacer(1, 6*mm))

    # ═══════════════════════════════════════════
    # VITAMIN D
    # ═══════════════════════════════════════════
    story.extend(section_header("03", "Восполнение дефицитов"))

    story.append(Paragraph("Витамин D — устранить недостаточность", ST["h2"]))
    story.append(Paragraph(
        '<font color="#DC4E42"><b>20.05 нг/мл</b></font> при норме 30-100. '
        'По комментарию лаборатории: &lt;20 — дефицит, 20-30 — недостаточность. '
        'Цель: 40-60 нг/мл.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>Почему важно в 19 лет:</b> витамин D критически важен для формирования пиковой костной массы '
        '(до 25 лет), иммунитета, настроения, здоровья кожи и волос. '
        'Дефицит витамина D ассоциирован с повышенным пролактином и нарушениями менструального цикла.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Витамин D3 (холекальциферол) — 4 000 МЕ/день",
        "Курс 8-12 недель, затем контроль 25(OH)D. После достижения 40-60 нг/мл — "
        "поддерживающая доза 2 000 МЕ/день."
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph('<b>Обязательные кофакторы:</b>', ST["body_bold"]))
    story.append(bullet('<b>Витамин K2 (МК-7)</b> — 100-200 мкг/день. Направляет кальций в кости, а не в сосуды. '
                        'Особенно важно, т.к. кальций ионизированный уже на верхней границе (1.32)'))
    story.append(bullet('<b>Магний (цитрат или глицинат)</b> — 300-400 мг/день. Нужен для активации витамина D'))
    story.append(Spacer(1, 1*mm))
    story.append(Paragraph(
        '<b>Продукты:</b> жирная рыба (лосось, сельдь, скумбрия), яичный желток, печень трески. '
        'Принимать D3 с жирной пищей в первой половине дня.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    # FERRITIN
    story.append(Paragraph("Ферритин — довести до оптимума", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>Ферритин 28.3 мкг/л</b></font> — формально в норме (10-120), но '
        'функциональный оптимум для молодой женщины — 40-80 мкг/л. '
        'При уровне ниже 40 мкг/л часто наблюдаются: усталость, выпадение волос, '
        'сухость кожи, ломкость ногтей, снижение концентрации.',
        ST["body"]
    ))
    story.append(Paragraph(
        'Железо сыворотки 12.70 мкмоль/л — в нижней трети нормы, что подтверждает '
        'субоптимальный статус железа. Гемоглобин 122 — пока в норме, но при дальнейшем снижении '
        'ферритина может развиться железодефицитная анемия.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Хелатное железо (бисглицинат) — 25 мг/день",
        "Утром натощак + витамин С (80-100 мг) для усвоения. Курс 2-3 мес. с контролем ферритина."
    ))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        '<b>Источники:</b> красное мясо, печень, морепродукты, гречка, чечевица, шпинат, гранат. '
        '<b>Не сочетать</b> с кальцием, цинком, чаем и кофе — разнести на 2-3 часа.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    # ZINC
    story.append(Paragraph("Цинк — поднять уровень", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>721.20 мкг/л</b></font> — нижняя граница (700-1140). '
        'Цинк участвует в 300+ ферментных реакциях: иммунитет, здоровье кожи и волос, '
        'синтез гормонов, подавление избыточного пролактина.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Цинк (пиколинат или бисглицинат) — 15-25 мг/день",
        "Вечером, отдельно от железа и кальция. Курс 2-3 месяца. "
        "Источники: говядина, тыквенные семечки, устрицы, кешью."
    ))

    story.append(Spacer(1, 4*mm))

    # B9
    story.append(Paragraph("Витамин B9 (фолат) — увеличить поступление", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>9.42 нмоль/л</b></font> — нижняя граница (7.0-39.7). '
        'Фолат критически важен для молодой женщины: участвует в делении клеток, '
        'кроветворении, синтезе ДНК, здоровье нервной системы. Низкий фолат при '
        'нормальном B12 и гомоцистеине может говорить о недостаточном поступлении с пищей.',
        ST["body"]
    ))

    story.append(Spacer(1, 2*mm))
    story.append(dose_box(
        "РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Метилфолат (5-МТГФ) — 400-800 мкг/день",
        "Предпочтительнее синтетической фолиевой кислоты. Утром, с едой. Курс 3 мес."
    ))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        '<b>Источники:</b> тёмно-зелёные листовые овощи (шпинат, брокколи, руккола), '
        'авокадо, спаржа, чечевица, нут, печень.',
        ST["body"]
    ))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # URINE + PHOSPHORUS + FIBRINOGEN
    # ═══════════════════════════════════════════
    story.extend(section_header("04", "Дополнительные находки"))

    story.append(Paragraph("Анализ мочи — ураты и белок", ST["h2"]))
    story.append(Paragraph(
        'В моче обнаружено <b>большое количество кристаллов уратов</b> (в норме — отсутствуют), '
        'моча мутная, белок незначительно повышен (0.16 г/л при норме до 0.15). '
        'Мочевая кислота в крови 155 мкмоль/л — в нижней трети нормы (142.8-339.2).',
        ST["body"]
    ))
    story.append(Paragraph(
        'Ураты в моче при нормальной мочевой кислоте в крови чаще всего связаны с:',
        ST["body"]
    ))
    story.append(bullet('<b>Недостаточным потреблением воды</b> — концентрированная моча'))
    story.append(bullet('<b>Кислой реакцией мочи</b> (pH 5.0 — нижняя граница) — ураты выпадают в осадок'))
    story.append(bullet('Избытком белковой пищи или мясных продуктов'))

    story.append(Paragraph("Рекомендации:", ST["h3"]))
    story.append(bullet('<b>Увеличить потребление воды</b> — минимум 30-35 мл на 1 кг веса в день'))
    story.append(bullet('<b>Ощелачивание:</b> увеличить долю овощей, фруктов, зелени в рационе'))
    story.append(bullet('Ограничить красное мясо до 2-3 раз в неделю, субпродукты — до 1 раза'))
    story.append(bullet('Исключить сладкие газированные напитки'))
    story.append(bullet('Пересдать ОАМ через 2-4 недели при соблюдении водного режима'))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Фосфор — незначительно повышен", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>1.55 ммоль/л</b></font> при норме 0.81-1.45. '
        'Небольшое превышение. В сочетании с кальцием ионизированным на верхней границе (1.32) '
        'и нормальным паратгормоном (32.70) — минеральный обмен в целом компенсирован.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>Рекомендации:</b> ограничить продукты с фосфатными добавками (колбасы, '
        'газированные напитки, плавленые сыры, полуфабрикаты). Повторить контроль через 3 мес.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Фибриноген — незначительно снижен", ST["h2"]))
    story.append(Paragraph(
        '<font color="#D4A04B"><b>1.98 г/л</b></font> при норме 2-3.9. '
        'Минимальное отклонение от нижней границы. Остальные показатели коагулограммы '
        '(D-димер, антитромбин III, АЧТВ, протромбин, МНО) — в норме. '
        'Скорее всего, не имеет клинического значения, но при частых носовых кровотечениях, '
        'обильных менструациях или длительных кровотечениях из ран — обратиться к гематологу.',
        ST["body"]
    ))
    story.append(Paragraph(
        '<b>Продукты для поддержки:</b> белковая пища (мясо, рыба, яйца), '
        'зелёные овощи (витамин K), гранат.',
        ST["body"]
    ))

    story.append(Spacer(1, 4*mm))

    # GENETIC
    story.append(Paragraph("Генетика — лактазная недостаточность", ST["h2"]))
    story.append(Paragraph(
        '<font color="#7C3AED"><b>Ген LCT — генотип CC</b></font> — первичная (генетическая) '
        'лактазная недостаточность. Это означает, что организм не вырабатывает достаточно '
        'фермента лактазы для переваривания молочного сахара (лактозы).',
        ST["body"]
    ))
    story.append(Paragraph("Что делать:", ST["h3"]))
    story.append(bullet('<b>Исключить или резко ограничить цельное молоко</b> и продукты с высоким '
                        'содержанием лактозы'))
    story.append(bullet('<b>Допустимы:</b> кисломолочные продукты (кефир, йогурт, творог) в небольших '
                        'количествах — в них лактоза частично расщеплена'))
    story.append(bullet('<b>Допустимы:</b> твёрдые и выдержанные сыры (пармезан, чеддер) — почти не содержат лактозы'))
    story.append(bullet('<b>При необходимости:</b> принимать фермент лактазу перед употреблением молочных продуктов'))
    story.append(bullet('<b>Альтернативные источники кальция:</b> кунжут (тахини), миндаль, брокколи, '
                        'листовая зелень, тофу, сардины с костями, обогащённое растительное молоко'))
    story.append(Paragraph(
        '<b>Важно:</b> поскольку молочные продукты ограничены, следить за достаточным '
        'поступлением кальция (1000 мг/день) из других источников.',
        ST["body"]
    ))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # NUTRITION
    # ═══════════════════════════════════════════
    story.extend(section_header("05", "Питание и образ жизни"))

    story.append(Paragraph(
        'Общая картина здоровья у Стояны хорошая: прекрасный липидный профиль, нет воспаления, '
        'щитовидка работает нормально, углеводный обмен без нарушений. Основные задачи — '
        'восполнить дефициты, разобраться с пролактином и скорректировать рацион с учётом '
        'лактазной недостаточности.',
        ST["body"]
    ))

    story.append(Paragraph("Принципы питания", ST["h3"]))
    story.append(bullet('<b>Белок в каждом приёме пищи</b> — 1.0-1.2 г на 1 кг веса: мясо, рыба, яйца, '
                        'бобовые, тофу'))
    story.append(bullet('<b>Овощи и зелень</b> — 400-500 г в день (источник фолата, калия, клетчатки)'))
    story.append(bullet('<b>Жирная рыба</b> — 2-3 раза в неделю (омега-3, витамин D)'))
    story.append(bullet('<b>Орехи и семена</b> — 30 г/день (цинк, магний, селен)'))
    story.append(bullet('<b>Безлактозные источники кальция</b> — кунжут, миндаль, тофу, листовая зелень, '
                        'сардины'))
    story.append(bullet('<b>Ограничить:</b> сахар, фаст-фуд, газировку, полуфабрикаты, '
                        'продукты с фосфатными добавками'))

    story.append(Paragraph("Водный режим", ST["h3"]))
    story.append(Paragraph(
        'Минимум <b>30-35 мл воды на 1 кг веса</b> в день. Это важно для профилактики '
        'образования уратных камней. Предпочтение — чистая вода, травяные чаи, вода с лимоном. '
        'Ограничить крепкий чай и кофе (не более 1-2 чашек, не натощак, не с приёмом железа).',
        ST["body"]
    ))

    story.append(Paragraph("Сон и стресс", ST["h3"]))
    story.append(Paragraph(
        'Для молодого организма и для нормализации пролактина крайне важны:',
        ST["body"]
    ))
    story.append(bullet('<b>Сон 8-9 часов</b> — ложиться до 23:00, исключить экраны за 1 час до сна'))
    story.append(bullet('<b>Управление стрессом</b> — стресс повышает пролактин'))
    story.append(bullet('Прогулки на свежем воздухе — минимум 30-40 мин/день (также витамин D от солнца)'))
    story.append(bullet('Умеренная физическая активность — 3-4 раза в неделю'))
    story.append(bullet('Избегать чрезмерных физических нагрузок — они повышают пролактин'))

    story.append(Paragraph("Что хорошо", ST["h3"]))
    story.append(Paragraph(
        'Стоит отметить отличные показатели, которые говорят о хорошей базе здоровья:',
        ST["body"]
    ))
    story.append(bullet('<b>Липидный профиль идеальный</b> — коэф. атерогенности 1.08 (при норме 2.2-3.5), '
                        'ЛПВП 1.98 — великолепный уровень'))
    story.append(bullet('<b>Нет воспаления</b> — СРБ &lt;0.6, гомоцистеин 8.50'))
    story.append(bullet('<b>Щитовидная железа</b> — все маркеры в норме'))
    story.append(bullet('<b>Углеводный обмен</b> — глюкоза, инсулин, HOMA-IR — в норме'))
    story.append(bullet('<b>Печень и почки</b> — все показатели в норме'))
    story.append(bullet('<b>Коэнзим Q10, селен, витамин B12</b> — в норме'))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PROTOCOL + CONTROL
    # ═══════════════════════════════════════════
    story.extend(section_header("06", "Сводный протокол приёма"))

    story.append(Paragraph(
        'Что принимать, когда и в какой дозе. Разнесено по времени для оптимального усвоения.',
        ST["body"]
    ))
    story.append(Spacer(1, 2*mm))
    story.append(protocol_table())

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Важные правила приёма", ST["h3"]))
    story.append(bullet('<b>Железо</b> — утром натощак, с витамином C. Не сочетать с цинком, кальцием, чаем, кофе'))
    story.append(bullet('<b>Цинк</b> — вечером, отдельно от железа (разнос минимум 4 часа)'))
    story.append(bullet('<b>Витамин D3 + K2</b> — утром, с жирной пищей'))
    story.append(bullet('<b>Магний</b> — вечером перед сном (улучшает сон)'))
    story.append(bullet('<b>Омега-3</b> — с едой (обед или ужин)'))
    story.append(bullet('<b>Метилфолат (B9)</b> — утром, с едой'))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Контрольные анализы", ST["h3"]))
    story.append(Paragraph(
        'Через <b>3 месяца</b> повторить:',
        ST["body"]
    ))
    story.append(bullet('Пролактин (после обследования у эндокринолога)'))
    story.append(bullet('Витамин D (25-OH)'))
    story.append(bullet('Ферритин, железо сыворотки'))
    story.append(bullet('Цинк в сыворотке'))
    story.append(bullet('Витамин B9 (фолат)'))
    story.append(bullet('Фосфор'))
    story.append(bullet('Общий анализ мочи'))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Приоритет действий", ST["h3"]))

    priority_data = [
        [Paragraph('<b>Срочно</b>', ST["table_cell_bold"]),
         Paragraph('Эндокринолог + МРТ гипофиза (пролактин)', ST["table_cell"])],
        [Paragraph('<b>В первую очередь</b>', ST["table_cell_bold"]),
         Paragraph('Витамин D3 + K2 + магний; увеличить потребление воды', ST["table_cell"])],
        [Paragraph('<b>Во вторую очередь</b>', ST["table_cell_bold"]),
         Paragraph('Железо, цинк, метилфолат, омега-3', ST["table_cell"])],
        [Paragraph('<b>Постоянно</b>', ST["table_cell_bold"]),
         Paragraph('Безлактозная диета, водный режим, сон, управление стрессом', ST["table_cell"])],
    ]
    priority_t = Table(priority_data, colWidths=[45*mm, 110*mm])
    priority_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), RED_BG),
        ("BACKGROUND", (0, 1), (0, 1), AMBER_BG),
        ("BACKGROUND", (0, 2), (0, 2), SAGE_BG),
        ("BACKGROUND", (0, 3), (0, 3), LIGHT_GRAY),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]))
    story.append(priority_t)

    story.append(Spacer(1, 5*mm))

    story.append(Paragraph("Источники", ST["h2"]))
    sources = [
        "Российская ассоциация эндокринологов — клинические рекомендации по гиперпролактинемии, 2023",
        "Российская ассоциация эндокринологов — клинические рекомендации по витамину D",
        "NIH — Vitamin D, Zinc, Folate Health Professional Fact Sheets",
        "Endocrine Society — Diagnosis and Treatment of Hyperprolactinemia, 2011",
        "NICE Guidelines — Prolactinoma: Diagnosis and Management",
        "PubMed — Vitamin B6 and dopamine synthesis; Vitex agnus-castus in hyperprolactinemia",
        "WHO — Guideline on Daily Iron Supplementation in Women, 2016",
        "Examine.com — Evidence-based Supplement Reviews (Zinc, Magnesium, Folate)",
        "EFSA — Scientific opinion on dietary reference values for phosphorus",
        "Российские клинические рекомендации — Лактазная недостаточность у взрослых",
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
    output = "/Users/serzhbigulov/Documents/Zalina/Рекомендации_нутрициолога_Алехина_Стояна.pdf"
    build_pdf(output)
