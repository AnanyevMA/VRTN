#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VRTN Workshop Official Letterhead Generator
Generates official corporate letters, statements, appeals, and claims for IP Ananiev M.A.
Outputs PDF (pixel-perfect with ReportLab) and DOCX (with python-docx).
"""

import os
import sys
import json
import argparse
from datetime import datetime

# Path resolution for assets
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
ASSETS_DIR = os.path.join(SKILL_DIR, "assets")

LOGO_PATH = os.path.join(ASSETS_DIR, "logo.png")
FONT_REGULAR = os.path.join(ASSETS_DIR, "Montserrat-Regular.ttf")
FONT_BOLD = os.path.join(ASSETS_DIR, "Montserrat-Bold.ttf")

DEFAULT_SENDER = {
    "name": "ИП Ананьев Максим Андреевич",
    "inn_ogrn": "ИНН 190207495060 | ОГРНИП 314190222000018",
    "production": "Производство: г. Саяногорск, Интернациональный мкрн, д. 1, оф. 89Н",
    "showroom": "Шоурум: г. Абакан, пр. Дружбы Народов, 20А",
    "contacts": "Тел.: +7 913 057-13-78 | E-mail: AnanyevM@mail.ru",
    "bank_account": "р/с 40802810171000005247 в АБАКАНСКОЕ ОТДЕЛЕНИЕ №8602 ПАО СБЕРБАНК",
    "bik_corr": "БИК 049514608 | к/с 30101810500000000608"
}

DEFAULT_BANK_DETAILS = [
    ("Получатель:", "Индивидуальный предприниматель Ананьев Максим Андреевич"),
    ("ИНН:", "190207495060"),
    ("Расчетный счет:", "40802810171000005247"),
    ("Банк получателя:", "АБАКАНСКОЕ ОТДЕЛЕНИЕ №8602 ПАО СБЕРБАНК"),
    ("БИК:", "049514608"),
    ("Корр. счет:", "30101810500000000608"),
]

def generate_pdf(data, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT

    # Register Montserrat
    pdfmetrics.registerFont(TTFont("Montserrat-Regular", FONT_REGULAR))
    pdfmetrics.registerFont(TTFont("Montserrat-Bold", FONT_BOLD))

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=45.685,
        rightMargin=45.685,
        topMargin=34.0,
        bottomMargin=34.0
    )

    printable_width = 595.2756 - (45.685 * 2)  # 503.9056 pt
    styles = getSampleStyleSheet()

    tagline_style = ParagraphStyle(
        "Tagline",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#4a4a4a"),
        spaceAfter=6
    )

    sender_style = ParagraphStyle(
        "SenderStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=7.2,
        leading=9.5,
        textColor=colors.HexColor("#222222")
    )

    recipient_style = ParagraphStyle(
        "RecipientStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.0,
        leading=10.5,
        textColor=colors.HexColor("#1a1a1a")
    )

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Montserrat-Bold",
        fontSize=11,
        leading=14,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1a1a1a"),
        spaceBefore=8,
        spaceAfter=3
    )

    subtitle_style = ParagraphStyle(
        "DocSubtitle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.5,
        leading=11.5,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#333333"),
        spaceAfter=10
    )

    salutation_style = ParagraphStyle(
        "Salutation",
        parent=styles["Normal"],
        fontName="Montserrat-Bold",
        fontSize=9.5,
        leading=12.5,
        textColor=colors.HexColor("#1a1a1a"),
        spaceAfter=7
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.5,
        leading=11.8,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#1a1a1a"),
        spaceAfter=6
    )

    callout_style = ParagraphStyle(
        "CalloutStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.5,
        leading=11.8,
        textColor=colors.HexColor("#1a1a1a")
    )

    bank_style = ParagraphStyle(
        "BankStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.0,
        leading=11.0,
        textColor=colors.HexColor("#1a1a1a")
    )

    app_style = ParagraphStyle(
        "AppStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=7.8,
        leading=10.5,
        textColor=colors.HexColor("#333333")
    )

    sign_style = ParagraphStyle(
        "SignStyle",
        parent=styles["Normal"],
        fontName="Montserrat-Regular",
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor("#1a1a1a")
    )

    story = []

    # 1. Logo
    logo_img = Image(LOGO_PATH, width=175.75, height=56.49)
    logo_img.hAlign = "LEFT"
    story.append(logo_img)

    # Tagline
    story.append(Paragraph("VRTN Workshop. С благодарностью за ваш выбор и доверие.", tagline_style))

    # Line 1
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC"), spaceBefore=3, spaceAfter=8))

    # Sender & Recipient table
    sender_p = f"""<b>{DEFAULT_SENDER['name']}</b><br/>
{DEFAULT_SENDER['inn_ogrn']}<br/>
{DEFAULT_SENDER['production']}<br/>
{DEFAULT_SENDER['showroom']}<br/>
{DEFAULT_SENDER['contacts']}<br/>
{DEFAULT_SENDER['bank_account']}<br/>
{DEFAULT_SENDER['bik_corr']}"""

    recipient_lines = []
    if data.get("recipient_position"):
        recipient_lines.append(f"<b>{data['recipient_position']}</b>")
    if data.get("recipient_company"):
        recipient_lines.append(f"<b>{data['recipient_company']}</b>")
    if data.get("recipient_name"):
        recipient_lines.append(f"<b>{data['recipient_name']}</b>")
    if data.get("recipient_address"):
        recipient_lines.append(data['recipient_address'].replace('\n', '<br/>'))
    if data.get("recipient_details"):
        recipient_lines.append(data['recipient_details'])

    recipient_p = "<br/>".join(recipient_lines)

    header_table = Table(
        [[Paragraph(sender_p, sender_style), Paragraph(recipient_p, recipient_style)]],
        colWidths=[290, 213.9]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(header_table)

    # Line 2
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#DDDDDD"), spaceBefore=8, spaceAfter=8))

    # Title & Subtitle
    if data.get("title"):
        story.append(Paragraph(data["title"], title_style))
    if data.get("subtitle"):
        story.append(Paragraph(data["subtitle"], subtitle_style))

    # Salutation
    if data.get("salutation"):
        story.append(Paragraph(data["salutation"], salutation_style))

    # Body paragraphs
    for p_text in data.get("paragraphs", []):
        story.append(Paragraph(p_text, body_style))

    # Callout / Highlight box (e.g. reason for refund, important terms)
    if data.get("callout"):
        callout_content = data["callout"]
        callout_box = Table([[Paragraph(callout_content, callout_style)]], colWidths=[printable_width])
        callout_box.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8F8F8')),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E0E0E0')),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(Spacer(1, 2))
        story.append(callout_box)
        story.append(Spacer(1, 6))

    # Bank details table
    if data.get("include_bank_details"):
        if data.get("bank_intro"):
            story.append(Paragraph(data["bank_intro"], body_style))
        else:
            story.append(Paragraph("Денежные средства просим перечислить по следующим банковским реквизитам:", body_style))

        bank_rows = []
        for label, val in DEFAULT_BANK_DETAILS:
            bank_rows.append([Paragraph(f"<b>{label}</b>", bank_style), Paragraph(val, bank_style)])

        bank_table = Table(bank_rows, colWidths=[120, printable_width - 120])
        bank_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 1.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('LINEBELOW', (0,0), (-1,-1), 0.25, colors.HexColor('#EEEEEE')),
        ]))
        story.append(bank_table)
        story.append(Spacer(1, 6))

    # Attachments
    if data.get("attachments"):
        app_html = "<b>Приложение:</b><br/>" + "<br/>".join(data["attachments"])
        story.append(Paragraph(app_html, app_style))
        story.append(Spacer(1, 10))

    # Signature Block
    date_str = data.get("date", f"«___» ____________ {datetime.now().year} г.")
    sig_left = "<b>Индивидуальный предприниматель</b><br/>Ананьев М. А."
    sig_right = f"_____________________ / М. А. Ананьев /<br/><br/>М. П. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{date_str}"

    sig_table = Table(
        [[Paragraph(sig_left, sign_style), Paragraph(sig_right, sign_style)]],
        colWidths=[240, 263.9]
    )
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(sig_table)

    def draw_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Montserrat-Regular", 7)
        canvas.setFillColor(colors.HexColor("#777777"))
        canvas.drawCentredString(595.2756 / 2.0, 28.35, "CREATED IN KHAKASSIA | EST. 2015 | VRTN WORKSHOP")
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_footer)
    print(f"Generated PDF: {output_path}")

def generate_docx(data, output_path):
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn

    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.63)
        section.right_margin = Inches(0.63)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Montserrat'
    font.size = Pt(8.5)
    font.color.rgb = RGBColor(0x1a, 0x1a, 0x1a)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
            node = OxmlElement(m)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def set_cell_shading(cell, color_hex):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
        for b_name, b_val in borders.items():
            if b_val:
                node = OxmlElement(f'w:{b_name}')
                node.set(qn('w:val'), b_val.get('val', 'single'))
                node.set(qn('w:sz'), str(b_val.get('sz', 4)))
                node.set(qn('w:space'), '0')
                node.set(qn('w:color'), b_val.get('color', 'CCCCCC'))
                tcBorders.append(node)
            else:
                node = OxmlElement(f'w:{b_name}')
                node.set(qn('w:val'), 'none')
                tcBorders.append(node)
        tcPr.append(tcBorders)

    # 1. Logo
    p_logo = doc.add_paragraph()
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(2)
    run_logo = p_logo.add_run()
    run_logo.add_picture(LOGO_PATH, width=Inches(2.44))

    # Tagline
    p_tag = doc.add_paragraph()
    p_tag.paragraph_format.space_before = Pt(0)
    p_tag.paragraph_format.space_after = Pt(4)
    run_tag = p_tag.add_run("VRTN Workshop. С благодарностью за ваш выбор и доверие.")
    run_tag.font.size = Pt(8.0)
    run_tag.font.color.rgb = RGBColor(0x4a, 0x4a, 0x4a)

    # Line 1
    t_l1 = doc.add_table(rows=1, cols=1)
    t_l1.alignment = WD_TABLE_ALIGNMENT.CENTER
    c1 = t_l1.rows[0].cells[0]
    c1.width = Inches(7.01)
    set_cell_borders(c1, bottom={'val': 'single', 'sz': 4, 'color': 'CCCCCC'})
    set_cell_margins(c1, top=0, bottom=0, left=0, right=0)
    p_e1 = c1.paragraphs[0]
    p_e1.paragraph_format.space_before = Pt(0)
    p_e1.paragraph_format.space_after = Pt(0)
    p_e1.paragraph_format.line_spacing = Pt(2)

    # Header Two Columns Table
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    cs = t_hdr.rows[0].cells[0]
    cr = t_hdr.rows[0].cells[1]
    cs.width = Inches(4.0)
    cr.width = Inches(3.01)
    set_cell_margins(cs, top=100, bottom=100, left=0, right=100)
    set_cell_margins(cr, top=100, bottom=100, left=100, right=0)
    set_cell_borders(cs)
    set_cell_borders(cr)

    # Sender
    ps = cs.paragraphs[0]
    ps.paragraph_format.space_before = Pt(2)
    ps.paragraph_format.space_after = Pt(0)
    ps.paragraph_format.line_spacing = Pt(10.5)
    r = ps.add_run(f"{DEFAULT_SENDER['name']}\n")
    r.bold = True
    r.font.size = Pt(7.5)
    sender_rest = (
        f"{DEFAULT_SENDER['inn_ogrn']}\n"
        f"{DEFAULT_SENDER['production']}\n"
        f"{DEFAULT_SENDER['showroom']}\n"
        f"{DEFAULT_SENDER['contacts']}\n"
        f"{DEFAULT_SENDER['bank_account']}\n"
        f"{DEFAULT_SENDER['bik_corr']}"
    )
    r2 = ps.add_run(sender_rest)
    r2.font.size = Pt(7.2)
    r2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Recipient
    pr = cr.paragraphs[0]
    pr.paragraph_format.space_before = Pt(2)
    pr.paragraph_format.space_after = Pt(0)
    pr.paragraph_format.line_spacing = Pt(11)

    recip_title = []
    if data.get("recipient_position"):
        recip_title.append(data["recipient_position"])
    if data.get("recipient_company"):
        recip_title.append(data["recipient_company"])
    if data.get("recipient_name"):
        recip_title.append(data["recipient_name"])
    if recip_title:
        rt_run = pr.add_run("\n".join(recip_title) + "\n")
        rt_run.bold = True
        rt_run.font.size = Pt(8.0)

    recip_body = []
    if data.get("recipient_address"):
        recip_body.append(data["recipient_address"])
    if data.get("recipient_details"):
        recip_body.append(data["recipient_details"])
    if recip_body:
        rb_run = pr.add_run("\n".join(recip_body))
        rb_run.font.size = Pt(8.0)

    # Line 2
    t_l2 = doc.add_table(rows=1, cols=1)
    t_l2.alignment = WD_TABLE_ALIGNMENT.CENTER
    c2 = t_l2.rows[0].cells[0]
    c2.width = Inches(7.01)
    set_cell_borders(c2, bottom={'val': 'single', 'sz': 4, 'color': 'DDDDDD'})
    set_cell_margins(c2, top=0, bottom=0, left=0, right=0)
    p_e2 = c2.paragraphs[0]
    p_e2.paragraph_format.space_before = Pt(0)
    p_e2.paragraph_format.space_after = Pt(0)
    p_e2.paragraph_format.line_spacing = Pt(2)

    # Title & Subtitle
    if data.get("title"):
        pt = doc.add_paragraph()
        pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pt.paragraph_format.space_before = Pt(10)
        pt.paragraph_format.space_after = Pt(2)
        r_t = pt.add_run(data["title"])
        r_t.bold = True
        r_t.font.size = Pt(11.0)

    if data.get("subtitle"):
        psub = doc.add_paragraph()
        psub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        psub.paragraph_format.space_before = Pt(0)
        psub.paragraph_format.space_after = Pt(10)
        r_sub = psub.add_run(data["subtitle"])
        r_sub.font.size = Pt(8.5)
        r_sub.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # Salutation
    if data.get("salutation"):
        psal = doc.add_paragraph()
        psal.paragraph_format.space_before = Pt(0)
        psal.paragraph_format.space_after = Pt(7)
        r_sal = psal.add_run(data["salutation"])
        r_sal.bold = True
        r_sal.font.size = Pt(9.5)

    # Body
    for p_text in data.get("paragraphs", []):
        pb = doc.add_paragraph()
        pb.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pb.paragraph_format.space_before = Pt(0)
        pb.paragraph_format.space_after = Pt(7)
        pb.paragraph_format.line_spacing = Pt(12)
        # Parse simple bold tags <b></b>
        parts = p_text.split("<b>")
        for part in parts:
            if "</b>" in part:
                bold_txt, regular_txt = part.split("</b>", 1)
                rb = pb.add_run(bold_txt)
                rb.bold = True
                pb.add_run(regular_txt)
            else:
                pb.add_run(part)

    # Callout
    if data.get("callout"):
        t_call = doc.add_table(rows=1, cols=1)
        t_call.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_call = t_call.rows[0].cells[0]
        c_call.width = Inches(7.01)
        set_cell_shading(c_call, "F8F8F8")
        set_cell_borders(c_call, top={'val': 'single', 'sz': 4, 'color': 'E0E0E0'},
                                  bottom={'val': 'single', 'sz': 4, 'color': 'E0E0E0'},
                                  left={'val': 'single', 'sz': 4, 'color': 'E0E0E0'},
                                  right={'val': 'single', 'sz': 4, 'color': 'E0E0E0'})
        set_cell_margins(c_call, top=80, bottom=80, left=120, right=120)
        pc = c_call.paragraphs[0]
        pc.paragraph_format.space_before = Pt(0)
        pc.paragraph_format.space_after = Pt(0)
        pc.paragraph_format.line_spacing = Pt(12)
        raw_callout = data["callout"]
        parts = raw_callout.split("<b>")
        for part in parts:
            if "</b>" in part:
                b_txt, reg_txt = part.split("</b>", 1)
                rb = pc.add_run(b_txt)
                rb.bold = True
                pc.add_run(reg_txt)
            else:
                pc.add_run(part)

    # Bank details
    if data.get("include_bank_details"):
        p_bintro = doc.add_paragraph()
        p_bintro.paragraph_format.space_before = Pt(8)
        p_bintro.paragraph_format.space_after = Pt(4)
        intro_txt = data.get("bank_intro", "Денежные средства просим перечислить по следующим банковским реквизитам:")
        p_bintro.add_run(intro_txt)

        t_bnk = doc.add_table(rows=len(DEFAULT_BANK_DETAILS), cols=2)
        t_bnk.alignment = WD_TABLE_ALIGNMENT.CENTER
        for idx, (label, val) in enumerate(DEFAULT_BANK_DETAILS):
            row = t_bnk.rows[idx]
            clbl, cval = row.cells[0], row.cells[1]
            clbl.width = Inches(1.8)
            cval.width = Inches(5.21)
            set_cell_margins(clbl, top=20, bottom=20, left=50, right=50)
            set_cell_margins(cval, top=20, bottom=20, left=50, right=50)
            set_cell_borders(clbl, bottom={'val': 'single', 'sz': 2, 'color': 'EEEEEE'})
            set_cell_borders(cval, bottom={'val': 'single', 'sz': 2, 'color': 'EEEEEE'})
            pl = clbl.paragraphs[0]
            pl.paragraph_format.space_before = Pt(1)
            pl.paragraph_format.space_after = Pt(1)
            rl = pl.add_run(label)
            rl.bold = True
            rl.font.size = Pt(8.0)
            pv = cval.paragraphs[0]
            pv.paragraph_format.space_before = Pt(1)
            pv.paragraph_format.space_after = Pt(1)
            rv = pv.add_run(val)
            rv.font.size = Pt(8.0)

    # Attachments
    if data.get("attachments"):
        pa = doc.add_paragraph()
        pa.paragraph_format.space_before = Pt(8)
        pa.paragraph_format.space_after = Pt(10)
        pa.paragraph_format.line_spacing = Pt(11)
        rah = pa.add_run("Приложение:\n")
        rah.bold = True
        rah.font.size = Pt(7.8)
        rab = pa.add_run("\n".join(data["attachments"]))
        rab.font.size = Pt(7.8)
        rab.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # Signature
    t_s = doc.add_table(rows=1, cols=2)
    t_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    csl, csr = t_s.rows[0].cells[0], t_s.rows[0].cells[1]
    csl.width = Inches(3.5)
    csr.width = Inches(3.51)
    set_cell_margins(csl, top=50, bottom=50, left=0, right=0)
    set_cell_margins(csr, top=50, bottom=50, left=0, right=0)
    set_cell_borders(csl)
    set_cell_borders(csr)

    psl = csl.paragraphs[0]
    psl.paragraph_format.space_before = Pt(0)
    psl.paragraph_format.space_after = Pt(0)
    rsl = psl.add_run("Индивидуальный предприниматель\nАнаньев М. А.")
    rsl.bold = True
    rsl.font.size = Pt(8.5)

    date_str = data.get("date", f"«___» ____________ {datetime.now().year} г.")
    psr = csr.paragraphs[0]
    psr.paragraph_format.space_before = Pt(0)
    psr.paragraph_format.space_after = Pt(0)
    psr.paragraph_format.line_spacing = Pt(14)
    rsr = psr.add_run(f"_____________________ / М. А. Ананьев /\n\nМ. П.         {date_str}")
    rsr.font.size = Pt(8.5)

    # Footer
    footer = doc.sections[0].footer
    pf = footer.paragraphs[0]
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = pf.add_run("CREATED IN KHAKASSIA | EST. 2015 | VRTN WORKSHOP")
    rf.font.size = Pt(7.0)
    rf.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

    doc.save(output_path)
    print(f"Generated DOCX: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Generate official letter for VRTN Workshop / IP Ananiev M.A.")
    parser.add_argument("--config", "-c", help="Path to JSON config with letter data")
    parser.add_argument("--output-pdf", help="Output PDF file path")
    parser.add_argument("--output-docx", help="Output DOCX file path")
    args = parser.parse_args()

    if not args.config:
        print("Usage: python generate_letter.py --config letter.json [--output-pdf path.pdf] [--output-docx path.docx]")
        sys.exit(1)

    with open(args.config, "r", encoding="utf-8") as f:
        data = json.load(f)

    base_name = data.get("output_basename", "Письмо_VRTN")
    out_dir = data.get("output_dir", os.getcwd())

    pdf_out = args.output_pdf or os.path.join(out_dir, f"{base_name}.pdf")
    docx_out = args.output_docx or os.path.join(out_dir, f"{base_name}.docx")

    formats = data.get("formats", ["pdf", "docx"])
    if "pdf" in formats:
        generate_pdf(data, pdf_out)
    if "docx" in formats:
        generate_docx(data, docx_out)

if __name__ == "__main__":
    main()
