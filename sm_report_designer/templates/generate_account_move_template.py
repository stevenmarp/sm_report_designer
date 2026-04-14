#!/usr/bin/env python3
"""
Generate a DOCX template for account.move (Journal Entry)
compatible with sm_report_designer placeholder syntax.

Usage:
    python3 generate_account_move_template.py

Output:
    account_move_template.docx
"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_cell_shading(cell, color_hex):
    """Set cell background color."""
    shading = cell._element.get_or_add_tcPr()
    shd = shading.makeelement(qn('w:shd'), {
        qn('w:fill'): color_hex,
        qn('w:val'): 'clear',
    })
    shading.append(shd)


def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    """Set cell borders."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.makeelement(qn('w:tcBorders'), {})
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            el = borders.makeelement(qn(f'w:{edge}'), {
                qn('w:val'): val.get('val', 'single'),
                qn('w:sz'): val.get('sz', '4'),
                qn('w:space'): '0',
                qn('w:color'): val.get('color', '000000'),
            })
            borders.append(el)
    tcPr.append(borders)


def add_styled_text(paragraph, text, bold=False, size=10, color=None, font_name='Calibri'):
    """Add a styled run to a paragraph."""
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = font_name
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run


def create_template():
    doc = Document()

    # -- Page margins --
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # ============================
    # HEADER SECTION
    # ============================

    # Company Name
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_styled_text(p, '{{company_id.name}}', bold=True, size=16, color=(44, 62, 80))

    # Document Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_styled_text(p, 'JOURNAL ENTRY', bold=True, size=14, color=(52, 73, 94))

    # Entry Number
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_styled_text(p, '{{name}}', bold=True, size=12, color=(41, 128, 185))

    # Spacer
    doc.add_paragraph()

    # ============================
    # INFO TABLE (2 columns)
    # ============================
    info_table = doc.add_table(rows=5, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    info_data = [
        ('Date', '{{date}}', 'Journal', '{{journal_id.name}}'),
        ('Reference', '{{ref}}', 'Status', '{{state}}'),
        ('Partner', '{{partner_id.name}}', 'Currency', '{{currency_id.name}}'),
        ('Move Type', '{{move_type}}', 'Company', '{{company_id.name}}'),
        ('Amount Total', '{{amount_total}}', 'Amount Due', '{{amount_residual}}'),
    ]

    for row_idx, (label1, val1, label2, val2) in enumerate(info_data):
        row = info_table.rows[row_idx]

        # Left label
        p = row.cells[0].paragraphs[0]
        add_styled_text(p, label1, bold=True, size=9, color=(100, 100, 100))
        set_cell_shading(row.cells[0], 'F5F5F5')

        # Left value
        p = row.cells[1].paragraphs[0]
        add_styled_text(p, val1, size=9)

        # Right label
        p = row.cells[2].paragraphs[0]
        add_styled_text(p, label2, bold=True, size=9, color=(100, 100, 100))
        set_cell_shading(row.cells[2], 'F5F5F5')

        # Right value
        p = row.cells[3].paragraphs[0]
        add_styled_text(p, val2, size=9)

    # Spacer
    doc.add_paragraph()

    # ============================
    # JOURNAL ITEMS TABLE (loop)
    # ============================
    p = doc.add_paragraph()
    add_styled_text(p, 'Journal Items', bold=True, size=11, color=(44, 62, 80))

    # Table: 7 columns
    cols = ['#', 'Account', 'Label', 'Partner', 'Debit', 'Credit', 'Balance']
    line_table = doc.add_table(rows=4, cols=7)
    line_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # -- Header row --
    header_row = line_table.rows[0]
    for i, col_name in enumerate(cols):
        p = header_row.cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_styled_text(p, col_name, bold=True, size=9, color=(255, 255, 255))
        set_cell_shading(header_row.cells[i], '2C3E50')

    # -- Loop start row --
    loop_start_row = line_table.rows[1]
    p = loop_start_row.cells[0].paragraphs[0]
    add_styled_text(p, '{{#line_ids}}', size=8, color=(150, 150, 150))
    # Merge all cells in this row for the loop marker
    for i in range(1, 7):
        p2 = loop_start_row.cells[i].paragraphs[0]
        p2.text = ''

    # -- Data row (template) --
    data_row = line_table.rows[2]
    placeholders = [
        '{{sequence}}',
        '{{account_id.name}}',
        '{{name}}',
        '{{partner_id.name}}',
        '{{debit}}',
        '{{credit}}',
        '{{balance}}',
    ]
    for i, ph in enumerate(placeholders):
        p = data_row.cells[i].paragraphs[0]
        if i >= 4:  # numeric columns right-aligned
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_styled_text(p, ph, size=9)

    # -- Loop end row --
    loop_end_row = line_table.rows[3]
    p = loop_end_row.cells[0].paragraphs[0]
    add_styled_text(p, '{{/line_ids}}', size=8, color=(150, 150, 150))
    for i in range(1, 7):
        p2 = loop_end_row.cells[i].paragraphs[0]
        p2.text = ''

    # Spacer
    doc.add_paragraph()

    # ============================
    # TOTALS SECTION
    # ============================
    totals_table = doc.add_table(rows=2, cols=4)
    totals_table.alignment = WD_TABLE_ALIGNMENT.RIGHT

    # Total Debit
    p = totals_table.rows[0].cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_styled_text(p, 'Total Debit:', bold=True, size=10, color=(44, 62, 80))

    p = totals_table.rows[0].cells[3].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_styled_text(p, '{{amount_total_signed}}', bold=True, size=10)

    # Total Credit
    p = totals_table.rows[1].cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_styled_text(p, 'Amount Due:', bold=True, size=10, color=(44, 62, 80))

    p = totals_table.rows[1].cells[3].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_styled_text(p, '{{amount_residual}}', bold=True, size=10, color=(192, 57, 43))

    # Spacer
    doc.add_paragraph()
    doc.add_paragraph()

    # ============================
    # NARRATION / NOTES
    # ============================
    p = doc.add_paragraph()
    add_styled_text(p, 'Notes:', bold=True, size=10, color=(100, 100, 100))

    p = doc.add_paragraph()
    add_styled_text(p, '{{narration}}', size=9)

    # ============================
    # FOOTER
    # ============================
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_styled_text(p, 'Generated by SM Report Designer', size=8, color=(180, 180, 180))

    # Save
    output_path = 'account_move_template.docx'
    doc.save(output_path)
    print(f'Template saved: {output_path}')
    print()
    print('Placeholders used:')
    print('  Header: name, date, ref, state, partner_id.name, journal_id.name,')
    print('          currency_id.name, company_id.name, move_type,')
    print('          amount_total, amount_residual')
    print('  Loop:   line_ids (account.move.line)')
    print('  Lines:  sequence, account_id.name, name, partner_id.name,')
    print('          debit, credit, balance')
    print('  Footer: amount_total_signed, amount_residual, narration')


if __name__ == '__main__':
    create_template()
