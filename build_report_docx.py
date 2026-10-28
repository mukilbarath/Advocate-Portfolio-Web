# -*- coding: utf-8 -*-
"""
Script to generate a comprehensive, publication-grade 30-page academic project report in .docx format
for KPR Institute of Engineering and Technology, Department of AI & ML.
Project: Revenue Generation Model - Client-Based Final Year Project Report
Title: Professional Portfolio Website for B. Kannan, Advocate, Madras High Court, Chennai
Authors: Mukil Barath S S (24AM028) & Salini N (24AM045)
"""

import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_color):
    """Set background color of a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tc_pr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner margins (padding) of a table cell in dxa (1 pt = 20 dxa)."""
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

def set_cell_borders(cell, top="CCCCCC", bottom="CCCCCC", left=None, right=None, sz="4"):
    """Set cell borders with XML."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders_xml = f'<w:tcBorders {nsdecls("w")}>'
    if top:
        borders_xml += f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{top}"/>'
    else:
        borders_xml += '<w:top w:val="none"/>'
    if bottom:
        borders_xml += f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{bottom}"/>'
    else:
        borders_xml += '<w:bottom w:val="none"/>'
    if left:
        borders_xml += f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{left}"/>'
    else:
        borders_xml += '<w:left w:val="none"/>'
    if right:
        borders_xml += f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{right}"/>'
    else:
        borders_xml += '<w:right w:val="none"/>'
    borders_xml += '</w:tcBorders>'
    tc_pr.append(parse_xml(borders_xml))

def format_table_header(row, col_widths, bg_color="0D1B2A", text_color="F5F0E8"):
    """Format the header row of a table."""
    for idx, cell in enumerate(row.cells):
        cell.width = col_widths[idx]
        set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=140, bottom=140, left=140, right=140)
        set_cell_borders(cell, top="0D1B2A", bottom="C9A84C", left="0D1B2A", right="0D1B2A", sz="8")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                run.font.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor.from_string(text_color)

def format_table_row(row, col_widths, bg_color="FFFFFF", is_even=False):
    """Format a standard body row in a table."""
    fill = "F9FBFD" if is_even else bg_color
    for idx, cell in enumerate(row.cells):
        cell.width = col_widths[idx]
        if fill:
            set_cell_background(cell, fill)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", sz="4")
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(30, 41, 59)

print("Helper functions initialized.")
