# -*- coding: utf-8 -*-
"""
build_full_report.py
Generates the comprehensive 30-page academic project report in Word (.docx) format:
"Professional Portfolio Website for B. Kannan, Advocate, Madras High Court, Chennai"
Department of AI & ML, KPR Institute of Engineering and Technology.
"""

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    doc = Document()
    
    # Page setup - 1 inch margins all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
        # Configure Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("KPR INSTITUTE OF ENGINEERING AND TECHNOLOGY | DEPARTMENT OF AI & ML")
        hrun.font.name = "Times New Roman"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(120, 120, 120)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        frun1 = fp.add_run("Assignment II | Client-Based Revenue Generation Project")
        frun1.font.name = "Times New Roman"
        frun1.font.size = Pt(8.5)
        frun1.font.color.rgb = RGBColor(120, 120, 120)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(33, 37, 41)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(4)

    # Color Constants
    NAVY_HEX = "0D1B2A"
    NAVY_RGB = RGBColor(13, 27, 42)
    GOLD_HEX = "C9A84C"
    GOLD_RGB = RGBColor(201, 168, 76)
    DARK_BLUE_RGB = RGBColor(19, 32, 52)
    GRAY_TEXT_RGB = RGBColor(90, 100, 110)

    def add_p(text="", bold=False, italic=False, size=11, color=RGBColor(33, 37, 41), align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=4, line_spacing=1.25):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            r = p.add_run(text)
            r.bold = bold
            r.italic = italic
            r.font.name = 'Times New Roman'
            r.font.size = Pt(size)
            r.font.color.rgb = color
        return p

    def add_h1(text, space_before=16, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.name = 'Times New Roman'
        r.font.size = Pt(15)
        r.font.color.rgb = NAVY_RGB
        return p

    def add_h2(text, space_before=12, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12.5)
        r.font.color.rgb = DARK_BLUE_RGB
        return p

    def add_h3(text, space_before=8, space_after=3):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(120, 80, 20)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.2
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.bold = True
            r1.font.name = 'Times New Roman'
            r1.font.size = Pt(10.5)
            r1.font.color.rgb = RGBColor(20, 30, 45)
        r2 = p.add_run(text)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = RGBColor(40, 45, 55)
        return p

    def set_cell_background(cell, fill_color):
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
        tc_pr.append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tcMar>
        ''')
        tc_pr.append(tc_mar)

    def set_cell_borders(cell, top="D0D7DE", bottom="D0D7DE", left="D0D7DE", right="D0D7DE", sz="4"):
        tc_pr = cell._tc.get_or_add_tcPr()
        borders_xml = f'<w:tcBorders {nsdecls("w")}>'
        borders_xml += f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{top}"/>' if top else '<w:top w:val="none"/>'
        borders_xml += f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{bottom}"/>' if bottom else '<w:bottom w:val="none"/>'
        borders_xml += f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{left}"/>' if left else '<w:left w:val="none"/>'
        borders_xml += f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{right}"/>' if right else '<w:right w:val="none"/>'
        borders_xml += '</w:tcBorders>'
        tc_pr.append(parse_xml(borders_xml))

    def format_row(row, col_widths, bg_color="FFFFFF", is_header=False):
        for idx, cell in enumerate(row.cells):
            cell.width = col_widths[idx]
            if is_header:
                set_cell_background(cell, "0D1B2A")
                set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
                set_cell_borders(cell, top="0D1B2A", bottom="C9A84C", left="0D1B2A", right="0D1B2A", sz="8")
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.bold = True
                        r.font.size = Pt(9.5)
                        r.font.color.rgb = RGBColor(245, 240, 232)
            else:
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
                set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", sz="4")
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(9.5)
                        r.font.color.rgb = RGBColor(30, 41, 59)

    print("Framework initialized. Writing chapters...")
    return doc

print("Script template ready.")
