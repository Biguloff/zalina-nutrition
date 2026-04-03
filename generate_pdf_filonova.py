#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate professional nutritionist recommendation PDF
for Филонова Татьяна Александровна
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
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

pdfmetrics.registerFontFamily("Georgia", normal="Georgia", bold="Georgia-Bold",
    italic="Georgia-Italic", boldItalic="Georgia-BoldItalic")
pdfmetrics.registerFontFamily("Arial", normal="Arial", bold="Arial-Bold", italic="Arial-Italic")

W, H = A4

# ─── Styles ───
def make_styles():
    s = {}
    s["title"] = ParagraphStyle("title", fontName="Georgia-Bold", fontSize=18, leading=24,
        textColor=CHARCOAL, alignment=TA_CENTER, spaceAfter=4*mm)
    s["subtitle"] = ParagraphStyle("subtitle", fontName="Georgia-Italic", fontSize=11, leading=15,
        textColor=SAGE, alignment=TA_CENTER, spaceAfter=6*mm)
    s["h1"] = ParagraphStyle("h1", fontName="Georgia-Bold", fontSize=15, leading=20,
        textColor=SAGE, spaceBefore=8*mm, spaceAfter=4*mm)
    s["h2"] = ParagraphStyle("h2", fontName="Georgia-Bold", fontSize=12, leading=16,
        textColor=CHARCOAL, spaceBefore=5*mm, spaceAfter=3*mm)
    s["h3"] = ParagraphStyle("h3", fontName="Arial-Bold", fontSize=10, leading=14,
        textColor=COPPER, spaceBefore=4*mm, spaceAfter=2*mm)
    s["body"] = ParagraphStyle("body", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, alignment=TA_JUSTIFY, spaceAfter=2*mm)
    s["body_bold"] = ParagraphStyle("body_bold", fontName="Arial-Bold", fontSize=9, leading=13.5,
        textColor=CHARCOAL, spaceAfter=2*mm)
    s["small"] = ParagraphStyle("small", fontName="Arial", fontSize=7.5, leading=11,
        textColor=GRAY, spaceAfter=1*mm)
    s["small_italic"] = ParagraphStyle("small_italic", fontName="Georgia-Italic", fontSize=7.5, leading=11,
        textColor=GRAY, spaceAfter=1*mm)
    s["table_header"] = ParagraphStyle("table_header", fontName="Arial-Bold", fontSize=8, leading=11,
        textColor=WHITE)
    s["table_cell"] = ParagraphStyle("table_cell", fontName="Arial", fontSize=8, leading=11.5,
        textColor=CHARCOAL)
    s["table_cell_bold"] = ParagraphStyle("table_cell_bold", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=CHARCOAL)
    s["alert_cell"] = ParagraphStyle("alert_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=RED_SOFT)
    s["warn_cell"] = ParagraphStyle("warn_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=AMBER)
    s["ok_cell"] = ParagraphStyle("ok_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=GREEN)
    s["info_cell"] = ParagraphStyle("info_cell", fontName="Arial-Bold", fontSize=8, leading=11.5,
        textColor=PURPLE)
    s["bullet"] = ParagraphStyle("bullet", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, leftIndent=12, spaceAfter=1.5*mm, bulletIndent=0, bulletFontSize=9)
    s["dose_label"] = ParagraphStyle("dose_label", fontName="Arial-Bold", fontSize=7.5, leading=10,
        textColor=SAGE)
    s["dose_value"] = ParagraphStyle("dose_value", fontName="Georgia-Bold", fontSize=11, leading=15,
        textColor=CHARCOAL)
    s["dose_note"] = ParagraphStyle("dose_note", fontName="Arial", fontSize=8, leading=12,
        textColor=GRAY)
    s["center"] = ParagraphStyle("center", fontName="Arial", fontSize=9, leading=13.5,
        textColor=CHARCOAL, alignment=TA_CENTER, spaceAfter=2*mm)
    s["alert_box"] = ParagraphStyle("alert_box", fontName="Arial", fontSize=9, leading=13.5,
        textColor=RED_SOFT, alignment=TA_LEFT, spaceAfter=2*mm)
    return s

ST = make_styles()


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
        c.setFillColor(SAGE)
        c.rect(0, h - 28*mm, w, 28*mm, fill=1, stroke=0)
        c.setStrokeColor(COPPER)
        c.setLineWidth(1.5)
        c.line(0, h - 28*mm, w, h - 28*mm)
        c.setFillColor(WHITE)
        c.setFont("Georgia-Bold", 14)
        c.drawString(20*mm, h - 12*mm, "Залина Бигулова")
        c.setFont("Georgia-Italic", 9)
        c.drawString(20*mm, h - 18*mm, "Нутрициолог-натуропат  |  Антиэйдж менеджмент  |  Спортивная диетология")
        c.setFont("Arial", 7.5)
        c.drawString(20*mm, h - 23.5*mm, "Эксперт МУНН  |  Член Национального общества нутрициологии")
        c.setFont("Arial", 7.5)
        c.setFillColor(HexColor("#D4E8D7"))
        c.drawRightString(w - 20*mm, h - 12*mm, "+7 918 822 7716  |  b-zalina@mail.ru")
        c.drawRightString(w - 20*mm, h - 18*mm, "taplink.cc/zalina_bigulova")
        c.setStrokeColor(SAGE_LIGHT)
        c.setLineWidth(0.5)
        c.line(20*mm, 14*mm, w - 20*mm, 14*mm)
        c.setFillColor(GRAY)
        c.setFont("Arial", 6.5)
        c.drawString(20*mm, 10*mm,
            "Данные рекомендации носят информационный характер. Перед приёмом добавок проконсультируйтесь с врачом.")
        c.drawRightString(w - 20*mm, 10*mm, f"Стр. {page_num} из {total_pages}")


# ─── Helpers ───
def colored_line():
    return HRFlowable(width="100%", thickness=0.5, color=SAGE_LIGHT, spaceBefore=2*mm, spaceAfter=2*mm)

def copper_line():
    return HRFlowable(width="40%", thickness=1, color=COPPER, spaceBefore=3*mm, spaceAfter=3*mm)

def section_header(number, title):
    return [Spacer(1, 3*mm),
        Paragraph(f'<font color="{COPPER.hexval()}" size="8">{number}</font>', ST["center"]),
        Paragraph(title, ST["h1"]), colored_line()]

def bullet(text):
    return Paragraph(f'<bullet>&bull;</bullet> {text}', ST["bullet"])

def dose_box(label, value, note=""):
    data = [[Paragraph(f'<font size="7" color="{SAGE.hexval()}"><b>{label}</b></font>', ST["dose_label"])],
            [Paragraph(f'<b>{value}</b>', ST["dose_value"])]]
    if note:
        data.append([Paragraph(note, ST["dose_note"])])
    t = Table(data, colWidths=[155*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), SAGE_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (0, 0), 8), ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]), ("LINEBEFOREDECL", (0, 0), (0, -1), 3, SAGE)]))
    return t

def alert_box(text):
    data = [[Paragraph(text, ST["alert_box"])]]
    t = Table(data, colWidths=[155*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), RED_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]), ("LINEBEFOREDECL", (0, 0), (0, -1), 3, RED_SOFT)]))
    return t

def status_tag(text, status="ok"):
    colors = {"ok": (GREEN, GREEN_BG), "warn": (AMBER, AMBER_BG),
              "alert": (RED_SOFT, RED_BG), "info": (PURPLE, PURPLE_BG)}
    style = {"ok": ST["ok_cell"], "warn": ST["warn_cell"],
             "alert": ST["alert_cell"], "info": ST["info_cell"]}[status]
    return Paragraph(text, style)

def analysis_table(title, rows):
    header = [Paragraph("Показатель", ST["table_header"]), Paragraph("Результат", ST["table_header"]),
              Paragraph("Норма", ST["table_header"]), Paragraph("Статус", ST["table_header"])]
    data = [header]
    for name, result, norm, st_text, st_type in rows:
        data.append([Paragraph(f'<b>{name}</b>', ST["table_cell_bold"]),
            Paragraph(result, ST["table_cell"]), Paragraph(norm, ST["table_cell"]),
            status_tag(st_text, st_type)])
    col_w = [60*mm, 35*mm, 35*mm, 25*mm]
    t = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), SAGE), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"), ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT), ("LINEBELOW", (0, 0), (-1, 0), 1, SAGE),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROUNDEDCORNERS", [3, 3, 3, 3])]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))
    t.setStyle(TableStyle(style_cmds))
    return KeepTogether([Paragraph(title, ST["h2"]), Spacer(1, 1*mm), t, Spacer(1, 3*mm)])

def protocol_table():
    header = [Paragraph("Добавка", ST["table_header"]), Paragraph("Дозировка", ST["table_header"]),
              Paragraph("Когда", ST["table_header"]), Paragraph("Курс", ST["table_header"])]
    rows_data = [
        ("Омега-3 (EPA+DHA)", "2 000-3 000 мг", "С едой (обед/ужин)", "Постоянно"),
        ("Берберин", "500 мг x 2-3 р/д", "За 30 мин до еды", "3 мес., контроль"),
        ("Растворимая клетчатка\n(псиллиум)", "5-10 г", "Перед сном, с водой", "Постоянно"),
        ("Магний (цитрат/глицинат)", "300-400 мг", "Вечер, перед сном", "Постоянно"),
        ("Витамин D3", "2 000 МЕ", "Утро, с жирной пищей", "Постоянно (поддержка)"),
        ("Витамин K2 (МК-7)", "100-200 мкг", "Утро, с D3", "Постоянно с D3"),
        ("Кальций (из пищи)", "1 000-1 200 мг", "С приёмами пищи", "Постоянно"),
        ("Противогрибковые\nнатуральные средства", "по протоколу", "С едой", "6-8 нед., см. раздел"),
    ]
    data = [header]
    for name, dose, when, course in rows_data:
        data.append([Paragraph(f'<b>{name}</b>', ST["table_cell_bold"]),
            Paragraph(dose, ST["table_cell"]), Paragraph(when, ST["table_cell"]),
            Paragraph(course, ST["table_cell"])])
    col_w = [45*mm, 30*mm, 40*mm, 40*mm]
    t = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), SAGE), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT), ("LINEBELOW", (0, 0), (-1, 0), 1, SAGE),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROUNDEDCORNERS", [3, 3, 3, 3])]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))
    t.setStyle(TableStyle(style_cmds))
    return t


# ─── Build ───
def build_pdf(output_path):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
        topMargin=34*mm, bottomMargin=20*mm, leftMargin=20*mm, rightMargin=20*mm,
        title="Рекомендации нутрициолога — Филонова Татьяна Александровна",
        author="Залина Бигулова",
        subject="Персональные рекомендации по результатам анализов")

    story = []

    # ═══════════════════════════════════════════
    # PAGE 1: TITLE + OVERVIEW
    # ═══════════════════════════════════════════
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph("Персональные рекомендации", ST["title"]))
    story.append(Paragraph("по результатам лабораторных исследований", ST["subtitle"]))
    story.append(copper_line())

    client_data = [
        [Paragraph('<b>Пациент:</b>', ST["table_cell_bold"]),
         Paragraph('Филонова Татьяна Александровна, 54 года (04.04.1971)', ST["table_cell"])],
        [Paragraph('<b>Дата анализов:</b>', ST["table_cell_bold"]),
         Paragraph('21.03.2026, лаборатория LabQuest, Москва', ST["table_cell"])],
        [Paragraph('<b>Дата рекомендаций:</b>', ST["table_cell_bold"]),
         Paragraph('Апрель 2026', ST["table_cell"])],
        [Paragraph('<b>Специалист:</b>', ST["table_cell_bold"]),
         Paragraph('Бигулова Залина Константиновна, нутрициолог-натуропат', ST["table_cell"])],
    ]
    client_t = Table(client_data, colWidths=[45*mm, 110*mm])
    client_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CREAM_DARK),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story.append(client_t)
    story.append(Spacer(1, 5*mm))

    story.append(alert_box(
        '<b>ВАЖНО: Выраженная дислипидемия — холестерин общий 8.74 ммоль/л, ЛПНП 6.3 ммоль/л '
        '(в 2.4 раза выше нормы). Необходима консультация кардиолога/терапевта для оценки '
        'сердечно-сосудистого риска и решения вопроса о медикаментозной терапии.</b>'
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Ключевые находки", ST["h2"]))
    story.append(Spacer(1, 1*mm))

    overview_data = [
        [Paragraph('<b>Холестерин общий</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">8.74 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;5.18', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>ЛПНП-холестерин</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">6.3 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;2.6', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Не-ЛПВП-холестерин</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">7.05 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;3.4', ST["small"]),
         status_tag("ПОВЫШЕН", "alert")],
        [Paragraph('<b>Коэф. атерогенности</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">4.17</font></b>', ST["table_cell"]),
         Paragraph('норма &lt;4.0', ST["small"]),
         status_tag("Повышен", "alert")],
        [Paragraph('<b>Грибковая нагрузка</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">7 831</font></b>', ST["table_cell"]),
         Paragraph('норма 2 332', ST["small"]),
         status_tag("x3.4 выше", "alert")],
        [Paragraph('<b>Вирусы (HHV-1,2)</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">3 730</font></b>', ST["table_cell"]),
         Paragraph('норма 800', ST["small"]),
         status_tag("x4.7 выше", "alert")],
        [Paragraph('<b>Плазмалоген</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#DC4E42">17</font></b>', ST["table_cell"]),
         Paragraph('норма 50', ST["small"]),
         status_tag("Снижен", "alert")],
        [Paragraph('<b>Натрий</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">134 ммоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма 136-145', ST["small"]),
         status_tag("Чуть снижен", "warn")],
        [Paragraph('<b>Витамин B12 актив.</b>', ST["table_cell_bold"]),
         Paragraph('<b><font color="#D4A04B">191 пмоль/л</font></b>', ST["table_cell"]),
         Paragraph('норма 25-165', ST["small"]),
         status_tag("Повышен", "warn")],
    ]
    overview_t = Table(overview_data, colWidths=[45*mm, 35*mm, 35*mm, 40*mm])
    overview_t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 6), RED_BG),
        ("BACKGROUND", (0, 7), (-1, 8), AMBER_BG),
        ("ROUNDEDCORNERS", [3, 3, 3, 3])]))
    story.append(overview_t)
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph(
        '<b>В норме:</b> витамин D (55.30 — отличный!), ЛПВП (1.69), триглицериды (0.92), '
        'гомоцистеин (8.96), фолиевая кислота, витамин A, кальций, калий, фосфор, хлор, '
        'щелочная фосфатаза, железо, инсулин, пролактин, АКТГ, эстрогены и прогестерон '
        '(постменопауза).',
        ST["small"]))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PAGE 2-3: ANALYSIS TABLES
    # ═══════════════════════════════════════════
    story.extend(section_header("01", "Результаты анализов — полная расшифровка"))

    story.append(analysis_table("Липидный профиль", [
        ("Холестерин общий", "8.74 ммоль/л", "<5.18", "Повышен", "alert"),
        ("ЛПНП-холестерин", "6.3 ммоль/л", "<2.6", "Повышен", "alert"),
        ("Не-ЛПВП-холестерин", "7.05 ммоль/л", "<3.4", "Повышен", "alert"),
        ("Коэф. атерогенности", "4.17", "<4.0", "Повышен", "alert"),
        ("ЛПВП-холестерин", "1.69 ммоль/л", ">1.55", "Норма", "ok"),
        ("Триглицериды", "0.92 ммоль/л", "<1.70", "Норма", "ok"),
    ]))

    story.append(analysis_table("Витамины", [
        ("Витамин D, 25-OH", "55.30 нг/мл", "20-100", "Отличный", "ok"),
        ("Витамин B12 активный", "191 пмоль/л", "25-165", "Повышен", "warn"),
        ("Фолиевая кислота", "7.7 нг/мл", "3.1-20.5", "Норма", "ok"),
        ("Витамин A (ретинол)", "0.52 мкг/мл", "0.30-1.20", "Норма", "ok"),
    ]))

    story.append(analysis_table("Минералы и электролиты", [
        ("Кальций", "2.42 ммоль/л", "2.15-2.5", "Норма", "ok"),
        ("Кальций ионизированный", "1.26 ммоль/л", "1.18-1.32", "Норма", "ok"),
        ("Фосфор", "1.31 ммоль/л", "0.78-1.42", "Норма", "ok"),
        ("Калий", "3.93 ммоль/л", "3.50-5.10", "Норма", "ok"),
        ("Натрий", "134 ммоль/л", "136-145", "Чуть снижен", "warn"),
        ("Хлор", "105 ммоль/л", "98-107", "Норма", "ok"),
        ("Железо", "13.6 мкмоль/л", "9-30.4", "Норма", "ok"),
    ]))

    story.append(analysis_table("Биохимия и прочее", [
        ("Гомоцистеин", "8.96 мкмоль/л", "4.44-13.56", "Норма", "ok"),
        ("Щелочная фосфатаза", "92 Ед/л", "42-98", "Норма", "ok"),
        ("Инсулин", "4.1 мкМЕ/мл", "2.6-24.9", "Норма", "ok"),
    ]))

    story.append(analysis_table("Гормональная панель", [
        ("Эстрадиол", "3.90 пг/мл", "постменоп. 2-21", "Норма", "ok"),
        ("Эстрон", "9.90 пг/мл", "постменоп. 3-32", "Норма", "ok"),
        ("Эстриол", "<0.01 нг/мл", "<0.08", "Норма", "ok"),
        ("Прогестерон", "<0.32 нмоль/л", "<0.64", "Норма", "ok"),
        ("Пролактин", "8.12 нг/мл", "3.00-26.71", "Норма", "ok"),
        ("АКТГ", "9.57 пг/мл", "0-46", "Норма", "ok"),
    ]))

    story.append(PageBreak())

    story.append(analysis_table("Микробиоценоз — ключевые отклонения", [
        ("Candida spp", "1 375", "493", "Повышен x2.8", "alert"),
        ("Aspergillus spp", "483", "188", "Повышен x2.6", "alert"),
        ("Micromycetes (каллистер.)", "2 497", "795", "Повышен x3.1", "alert"),
        ("Micromycetes (тентостер.)", "3 477", "857", "Повышен x4.1", "alert"),
        ("HHV-1,2 (герпес)", "3 730", "800", "Повышен x4.7", "alert"),
        ("Staphylococcus spp", "708", "464", "Повышен", "warn"),
        ("Propionibact. freudenr.", "3 753", "1 868", "Повышен x2", "warn"),
        ("Alcaligenes spp", "103", "60", "Повышен", "warn"),
        ("Eubacterium spp", "3 677", "6 364", "Снижен", "warn"),
        ("Nocardia asteroides", "353", "1 063", "Снижен", "warn"),
        ("Streptococcus spp", "0", "144", "Снижен", "warn"),
        ("Плазмалоген (16а)", "17", "50", "Снижен", "alert"),
        ("Эндотоксин (сумма)", "0.3", "0.5", "Норма", "ok"),
    ]))

    # Summary of microbiome
    micro_summary = [
        [Paragraph('<b>Показатель</b>', ST["table_cell_bold"]),
         Paragraph('<b>Проба</b>', ST["table_cell_bold"]),
         Paragraph('<b>Норма</b>', ST["table_cell_bold"]),
         Paragraph('<b>Статус</b>', ST["table_cell_bold"])],
        [Paragraph('Общая бактериальная нагрузка', ST["table_cell"]),
         Paragraph('18 665', ST["table_cell"]),
         Paragraph('21 257', ST["table_cell"]),
         status_tag("Снижена", "warn")],
        [Paragraph('Микроскопические грибы', ST["table_cell"]),
         Paragraph('7 831', ST["table_cell"]),
         Paragraph('2 332', ST["table_cell"]),
         status_tag("x3.4", "alert")],
        [Paragraph('Вирусы', ST["table_cell"]),
         Paragraph('3 730', ST["table_cell"]),
         Paragraph('1 444', ST["table_cell"]),
         status_tag("x2.6", "alert")],
        [Paragraph('Общая микробная нагрузка', ST["table_cell"]),
         Paragraph('30 227', ST["table_cell"]),
         Paragraph('25 033', ST["table_cell"]),
         status_tag("Повышена", "warn")],
    ]
    ms_t = Table(micro_summary, colWidths=[55*mm, 30*mm, 30*mm, 30*mm])
    ms_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), SAGE), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROUNDEDCORNERS", [3, 3, 3, 3])]))
    story.append(Paragraph("Сводка микробиоценоза", ST["h2"]))
    story.append(ms_t)

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # LIPIDS — MAIN PROBLEM
    # ═══════════════════════════════════════════
    story.extend(section_header("02", "Дислипидемия — снижение холестерина"))

    story.append(alert_box(
        '<b>Холестерин общий 8.74 ммоль/л (норма &lt;5.18) и ЛПНП 6.3 ммоль/л (норма &lt;2.6) '
        'значительно повышены. Это основной фактор сердечно-сосудистого риска, '
        'особенно в постменопаузе. Необходима консультация кардиолога.</b>'
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Почему холестерин так высок", ST["h3"]))
    story.append(Paragraph(
        'В постменопаузе (эстрадиол 3.9 пг/мл) утрачивается защитное действие эстрогенов '
        'на липидный обмен. Эстрогены стимулировали рецепторы ЛПНП в печени, способствуя '
        'выведению «плохого» холестерина. Без этого механизма ЛПНП растёт на 10-20% в первые '
        '2-3 года после менопаузы. У 77% женщин 45-65 лет наблюдается повышенный холестерин. '
        'Однако уровень ЛПНП 6.3 — это выраженное превышение, которое требует активной коррекции.',
        ST["body"]))

    story.append(Paragraph(
        '<b>Хорошие новости:</b> ЛПВП 1.69 — хороший уровень (выше 1.55), триглицериды 0.92 — '
        'в норме, инсулин 4.1 — отлично, гомоцистеин 8.96 — норма, воспаления нет.',
        ST["body"]))

    story.append(Paragraph("Нутрицевтическая коррекция", ST["h3"]))

    story.append(Paragraph('<b>1. Омега-3 жирные кислоты (EPA + DHA):</b>', ST["body_bold"]))
    story.append(dose_box("РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Омега-3 — 2 000-3 000 мг EPA+DHA в день",
        "Снижает триглицериды на 20-30%, улучшает функцию ЛПВП, противовоспалительное действие. С едой."))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph('<b>2. Берберин:</b>', ST["body_bold"]))
    story.append(dose_box("РЕКОМЕНДУЕМЫЙ ПРИЁМ",
        "Берберин — 500 мг, 2-3 раза в день перед едой",
        "Снижает ЛПНП на 20-25% и общий холестерин на 15-20%. Механизм действия схож со статинами "
        "(активация AMPK и увеличение рецепторов ЛПНП). Курс 3 мес. с контролем липидного профиля."))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph('<b>3. Растворимая клетчатка</b> — 10-25 г/день:', ST["body_bold"]))
    story.append(bullet('Псиллиум (подорожник) — 5-10 г/день. Снижает ЛПНП на 5-10%'))
    story.append(bullet('Овсянка (бета-глюканы), яблоки, семена льна, чечевица'))

    story.append(Paragraph('<b>4. Растительные стеролы</b> — 2 г/день. Снижают ЛПНП на 6-15%.', ST["body_bold"]))

    story.append(Paragraph('<b>5. Красный ферментированный рис</b> — рассмотреть с врачом. '
        'Содержит натуральный монаколин К, аналогичный ловастатину.', ST["body_bold"]))

    story.append(Spacer(1, 3*mm))
    story.append(Paragraph("Диетические рекомендации по холестерину", ST["h3"]))
    story.append(bullet('<b>Оливковое масло extra virgin</b> — 2-3 ст. ложки в день (основной жир)'))
    story.append(bullet('<b>Орехи</b> (грецкие, миндаль) — 30 г/день'))
    story.append(bullet('<b>Жирная рыба</b> — 3-4 раза в неделю (лосось, скумбрия, сельдь)'))
    story.append(bullet('<b>Авокадо</b> — половинка в день'))
    story.append(bullet('<b>Бобовые</b> — чечевица, нут, фасоль — 3-4 раза в неделю'))
    story.append(bullet('<b>Исключить:</b> трансжиры, маргарин, фаст-фуд, промышленная выпечка'))
    story.append(bullet('<b>Ограничить:</b> рафинированный сахар, белую муку, избыток насыщенных жиров'))

    story.append(Spacer(1, 6*mm))

    # ═══════════════════════════════════════════
    # MICROBIOME
    # ═══════════════════════════════════════════
    story.extend(section_header("03", "Микробиоценоз — грибковая и вирусная нагрузка"))

    story.append(Paragraph(
        'Анализ по Осипову выявил <b>значительное повышение грибковой нагрузки</b> (в 3.4 раза выше нормы) '
        'и <b>повышение вирусной активности</b> (герпес HHV-1,2 — в 4.7 раза выше нормы). '
        'Также снижен плазмалоген (17 vs 50) — маркер целостности клеточных мембран. '
        'Эндотоксин в норме (0.3 vs 0.5) — кишечная проницаемость сохранена.',
        ST["body"]))

    story.append(Paragraph("Грибковая нагрузка — Candida и Aspergillus", ST["h3"]))
    story.append(Paragraph(
        'Candida spp повышена в 2.8 раза, Aspergillus — в 2.6 раза, Micromycetes — в 3-4 раза. '
        'Грибковый дисбаланс связан со снижением нормофлоры (Eubacterium снижен почти вдвое, '
        'Streptococcus отсутствует, Nocardia снижен в 3 раза). Нормальная микрофлора конкурирует '
        'с грибками за питательные субстраты — при её снижении грибки разрастаются.',
        ST["body"]))

    story.append(Paragraph("Противогрибковый протокол (6-8 недель):", ST["h3"]))
    story.append(bullet('<b>Каприловая кислота</b> — 1 000-2 000 мг/день (натуральный антифунгал из кокоса)'))
    story.append(bullet('<b>Экстракт орегано (карвакрол)</b> — 200-600 мг/день'))
    story.append(bullet('<b>Берберин</b> — 500 мг x 2 р/день (также антигрибковое + снижает холестерин)'))
    story.append(bullet('<b>Чеснок (аллицин)</b> — экстракт 600-1 200 мг/день или 2-3 зубчика сырого'))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        '<b>Антигрибковая диета (6-8 недель):</b> исключить сахар, дрожжевой хлеб, алкоголь, '
        'виноград, сухофрукты, грибы, сыры с плесенью, квас. Увеличить овощи, белок, '
        'кокосовое масло.',
        ST["body"]))

    story.append(Paragraph("Восстановление нормофлоры:", ST["h3"]))
    story.append(bullet('<b>Пробиотики</b> — Lactobacillus rhamnosus GG + Saccharomyces boulardii '
                        '(конкурент Candida). После окончания противогрибковой фазы'))
    story.append(bullet('<b>Пребиотики</b> — инулин, ФОС (фруктоолигосахариды) — 5-10 г/день'))
    story.append(bullet('<b>Ферментированные продукты</b> — квашеная капуста, кимчи, кефир (если нет грибков)'))
    story.append(bullet('<b>Клетчатка</b> — 25-35 г/день из овощей, бобовых, цельнозерновых'))

    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Вирусная нагрузка — герпес HHV-1,2", ST["h3"]))
    story.append(Paragraph(
        'HHV-1,2 повышен в 4.7 раза (3 730 vs 800). Герпесвирусы — хронические инфекции, которые '
        'реактивируются при снижении иммунитета. Снижение нормофлоры и грибковый дисбаланс могут '
        'способствовать реактивации.',
        ST["body"]))
    story.append(bullet('<b>L-лизин</b> — 1 000-1 500 мг/день (подавляет репликацию герпеса)'))
    story.append(bullet('<b>Цинк</b> — 15-25 мг/день (иммуномодулятор)'))
    story.append(bullet('<b>Витамин C</b> — 500-1 000 мг/день'))
    story.append(bullet('Ограничить продукты с аргинином: орехи, шоколад, семечки (аргинин стимулирует герпес)'))

    story.append(Spacer(1, 3*mm))

    story.append(Paragraph("Плазмалоген — снижен", ST["h3"]))
    story.append(Paragraph(
        '<b>Плазмалоген 17</b> (норма 50) — значительно снижен. Плазмалогены — это '
        'фосфолипиды клеточных мембран, важные антиоксиданты. Их снижение ассоциировано '
        'с воспалительными процессами в кишечнике и окислительным стрессом.',
        ST["body"]))
    story.append(bullet('<b>Омега-3</b> (уже в протоколе) — предшественник плазмалогенов'))
    story.append(bullet('<b>Антиоксиданты:</b> витамин E — 200-400 МЕ/день, коэнзим Q10 — 100-200 мг/день'))
    story.append(bullet('Морепродукты (особенно моллюски) — богаты плазмалогенами'))

    story.append(Spacer(1, 6*mm))

    # ═══════════════════════════════════════════
    # POSTMENOPAUSE
    # ═══════════════════════════════════════════
    story.extend(section_header("04", "Постменопауза — нутритивная поддержка"))

    story.append(Paragraph(
        'Гормональный профиль подтверждает устойчивую постменопаузу: эстрадиол 3.9 пг/мл, '
        'прогестерон &lt;0.32, пролактин и АКТГ в норме. Щитовидная железа не исследована в данном пакете — '
        'рекомендуется проверить ТТГ, Т4 св., Т3 св. при отсутствии данных.',
        ST["body"]))

    story.append(Paragraph("Кости и кальций", ST["h3"]))
    story.append(Paragraph(
        'Кальций 2.42, кальций ионизированный 1.26, фосфор 1.31 — всё в норме. '
        'Витамин D 55.30 — отличный уровень. Но в постменопаузе без эстрогенов '
        'ускоряется потеря костной массы — профилактика остеопороза обязательна.',
        ST["body"]))
    story.append(bullet('<b>Кальций из пищи</b> — 1 000-1 200 мг/день (кунжут, тахини, миндаль, '
                        'сардины с костями, брокколи, листовая зелень, твёрдый сыр)'))
    story.append(bullet('<b>Витамин D3</b> — 2 000 МЕ/день (поддержка текущего уровня 55 нг/мл)'))
    story.append(bullet('<b>Витамин K2 (МК-7)</b> — 100-200 мкг/день (направляет кальций в кости)'))
    story.append(bullet('<b>Магний</b> — 300-400 мг/день (участвует в метаболизме кальция и витамина D)'))
    story.append(bullet('<b>Силовые тренировки</b> — 2-3 раза/неделю (стимулируют остеобласты)'))
    story.append(bullet('Рассмотреть денситометрию (DEXA) для оценки плотности костей'))

    story.append(Paragraph("Сердечно-сосудистая система", ST["h3"]))
    story.append(Paragraph(
        'В постменопаузе сердечно-сосудистый риск у женщин возрастает до уровня мужчин. '
        'При ЛПНП 6.3 и коэф. атерогенности 4.17 — это приоритет номер один.',
        ST["body"]))
    story.append(bullet('Обязательна <b>консультация кардиолога</b> — оценка SCORE2, '
                        'решение вопроса о статинах'))
    story.append(bullet('Контроль артериального давления'))
    story.append(bullet('ЭКГ и УЗДГ сосудов шеи (если не делалось)'))

    story.append(Paragraph("Натрий — незначительно снижен", ST["h3"]))
    story.append(Paragraph(
        '<b>134 ммоль/л</b> при норме 136-145. Минимальное отклонение. Может быть связано с '
        'обильным питьём, приёмом диуретиков или разведением крови. Если нет симптомов '
        '(головокружение, слабость, тошнота) — клинически незначимо. '
        'Не требует специальной коррекции, но стоит контролировать повторно.',
        ST["body"]))

    story.append(Paragraph("Витамин B12 — повышен", ST["h3"]))
    story.append(Paragraph(
        '<b>191 пмоль/л</b> при норме 25-165. Незначительное повышение. Чаще всего связано '
        'с приёмом добавок B12. Если принимаете — можно временно приостановить. '
        'Повышенный B12 без приёма добавок иногда требует дообследования (печень), '
        'но при нормальном гомоцистеине (8.96) — обычно не вызывает опасений.',
        ST["body"]))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # NUTRITION
    # ═══════════════════════════════════════════
    story.extend(section_header("05", "Питание и образ жизни"))

    story.append(Paragraph("Принципы питания", ST["h3"]))
    story.append(bullet('<b>Средиземноморский стиль</b> — основа рациона'))
    story.append(bullet('<b>Овощи</b> — 400-500 г в день (половина тарелки). Особенно крестоцветные '
                        '(брокколи, капуста) и листовая зелень'))
    story.append(bullet('<b>Белок</b> — 1.0-1.2 г на 1 кг веса: рыба, птица, яйца, бобовые, тофу'))
    story.append(bullet('<b>Жирная рыба</b> — 3-4 раза в неделю (омега-3)'))
    story.append(bullet('<b>Оливковое масло</b> — основной жир (2-3 ст. ложки)'))
    story.append(bullet('<b>Клетчатка</b> — 25-35 г/день (бобовые, овощи, овсянка, псиллиум)'))
    story.append(bullet('<b>Ферментированные продукты</b> — квашеная капуста, кимчи (после противогрибковой фазы)'))
    story.append(bullet('<b>Ограничить:</b> сахар (это питание для Candida!), белая мука, алкоголь, '
                        'маргарин, фаст-фуд, колбасы'))

    story.append(Paragraph("Антигрибковая фаза (первые 6-8 недель)", ST["h3"]))
    story.append(Paragraph(
        'В этот период дополнительно исключить: дрожжевой хлеб, виноград, бананы, '
        'сухофрукты, мёд, сыры с плесенью, грибы, квас, пиво. Сахар — полностью.',
        ST["body"]))

    story.append(Paragraph("Физическая активность", ST["h3"]))
    story.append(bullet('<b>Силовые тренировки</b> — 2-3 раза/нед (кости, мышцы, метаболизм)'))
    story.append(bullet('<b>Ходьба</b> — 8 000-10 000 шагов/день'))
    story.append(bullet('<b>Кардио</b> — 2-3 раза/нед по 30-40 мин (плавание, велосипед)'))
    story.append(bullet('Избегать чрезмерных нагрузок — повышают кортизол'))

    story.append(Paragraph("Сон и стресс", ST["h3"]))
    story.append(bullet('<b>Сон 7-8 часов</b>. Магний перед сном (300-400 мг) — улучшает сон'))
    story.append(bullet('Стресс активирует герпес и подавляет иммунитет'))
    story.append(bullet('Прогулки на свежем воздухе, дыхательные техники, йога'))

    story.append(PageBreak())

    # ═══════════════════════════════════════════
    # PROTOCOL
    # ═══════════════════════════════════════════
    story.extend(section_header("06", "Сводный протокол приёма"))
    story.append(Paragraph(
        'Что принимать, когда и в какой дозе. Разнесено по времени для оптимального усвоения.',
        ST["body"]))
    story.append(Spacer(1, 2*mm))
    story.append(protocol_table())

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Важные правила приёма", ST["h3"]))
    story.append(bullet('<b>Берберин</b> — за 30 мин до еды, 2-3 раза в день'))
    story.append(bullet('<b>Омега-3</b> — с едой (обед или ужин)'))
    story.append(bullet('<b>Витамин D3 + K2</b> — утром, с жирной пищей'))
    story.append(bullet('<b>Магний</b> — вечером перед сном'))
    story.append(bullet('<b>Псиллиум</b> — перед сном, с большим количеством воды, отдельно от лекарств'))
    story.append(bullet('<b>Противогрибковые</b> — начинать постепенно (1 средство, затем добавлять)'))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Контрольные анализы", ST["h3"]))
    story.append(Paragraph('Через <b>3 месяца</b> повторить:', ST["body"]))
    story.append(bullet('Липидный профиль (холестерин общий, ЛПНП, ЛПВП, триглицериды)'))
    story.append(bullet('Микробиоценоз по Осипову'))
    story.append(bullet('Натрий'))
    story.append(bullet('АЛТ, АСТ (контроль на фоне берберина)'))
    story.append(bullet('ТТГ, Т4 св. (если ещё не сдавала)'))
    story.append(bullet('Денситометрия (DEXA) — при отсутствии данных'))

    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Приоритет действий", ST["h3"]))
    priority_data = [
        [Paragraph('<b>Срочно</b>', ST["table_cell_bold"]),
         Paragraph('Кардиолог (ЛПНП 6.3, коэф. атерогенности 4.17)', ST["table_cell"])],
        [Paragraph('<b>В первую очередь</b>', ST["table_cell_bold"]),
         Paragraph('Берберин + омега-3 + клетчатка (снижение холестерина)', ST["table_cell"])],
        [Paragraph('<b>Параллельно</b>', ST["table_cell_bold"]),
         Paragraph('Противогрибковый протокол (6-8 недель)', ST["table_cell"])],
        [Paragraph('<b>Постоянно</b>', ST["table_cell_bold"]),
         Paragraph('Витамин D + K2 + магний + кальций; антигрибковая диета; физ. активность', ST["table_cell"])],
    ]
    priority_t = Table(priority_data, colWidths=[45*mm, 110*mm])
    priority_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), RED_BG),
        ("BACKGROUND", (0, 1), (0, 1), AMBER_BG),
        ("BACKGROUND", (0, 2), (0, 2), SAGE_BG),
        ("BACKGROUND", (0, 3), (0, 3), LIGHT_GRAY),
        ("GRID", (0, 0), (-1, -1), 0.3, BORDER_LIGHT),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROUNDEDCORNERS", [3, 3, 3, 3])]))
    story.append(priority_t)

    story.append(Spacer(1, 6*mm))
    story.append(colored_line())
    story.append(Spacer(1, 2*mm))

    # Signature
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
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story.append(sig_t)

    doc.build(story, canvasmaker=LetterheadCanvas)
    print(f"PDF saved to: {output_path}")


if __name__ == "__main__":
    output = "/Users/serzhbigulov/Documents/Zalina/Рекомендации_нутрициолога_Филонова_Татьяна.pdf"
    build_pdf(output)
