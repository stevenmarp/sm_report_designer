#!/usr/bin/env python3
"""
Generate XLSX templates for account.move and sale.order
compatible with sm_report_designer placeholder syntax.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def style_header_cell(ws, row, col, value, width=18):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=True, size=9, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='2C3E50')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = Border(
        bottom=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
    )
    ws.column_dimensions[get_column_letter(col)].width = width
    return cell


def style_label(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=True, size=9, color='555555')
    cell.fill = PatternFill('solid', fgColor='F2F2F2')
    cell.alignment = Alignment(vertical='center')
    cell.border = Border(
        bottom=Side(style='thin', color='CCCCCC'),
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
    )
    return cell


def style_value(ws, row, col, value, align='left'):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(size=9)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    cell.border = Border(
        bottom=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
    )
    return cell


def style_data_cell(ws, row, col, value, align='left'):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(size=9)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    cell.border = Border(
        bottom=Side(style='thin', color='DDDDDD'),
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
    )
    return cell


def style_loop_marker(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(size=7, color='999999', italic=True)
    return cell


def style_total_label(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=True, size=10, color='2C3E50')
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = Border(top=Side(style='double', color='2C3E50'))
    return cell


def style_total_value(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=True, size=10, color='C0392B')
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = Border(top=Side(style='double', color='2C3E50'))
    return cell


# ==============================================================
#  ACCOUNT MOVE TEMPLATE
# ==============================================================
def create_account_move_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Journal Entry'

    # -- Title --
    r = 1
    title = ws.cell(row=r, column=1, value='{{company_id.name}}')
    title.font = Font(bold=True, size=16, color='2C3E50')
    ws.merge_cells('A1:G1')

    r = 2
    sub = ws.cell(row=r, column=1, value='JOURNAL ENTRY')
    sub.font = Font(bold=True, size=13, color='34495E')
    ws.merge_cells('A2:G2')

    r = 3
    ref = ws.cell(row=r, column=1, value='{{name}}')
    ref.font = Font(bold=True, size=11, color='2980B9')
    ws.merge_cells('A3:G3')

    r = 4  # spacer

    # -- Info section --
    info = [
        ('Date', '{{date}}', 'Journal', '{{journal_id.name}}'),
        ('Reference', '{{ref}}', 'Status', '{{state}}'),
        ('Partner', '{{partner_id.name}}', 'Currency', '{{currency_id.name}}'),
        ('Move Type', '{{move_type}}', 'Company', '{{company_id.name}}'),
        ('Amount Total', '{{amount_total}}', 'Amount Due', '{{amount_residual}}'),
    ]

    for i, (l1, v1, l2, v2) in enumerate(info):
        r = 5 + i
        style_label(ws, r, 1, l1)
        style_value(ws, r, 2, v1)
        # col 3 spacer
        style_label(ws, r, 4, l2)
        style_value(ws, r, 5, v2)

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 3
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 25

    r = 11  # spacer

    # -- Journal Items header --
    r = 12
    headers = ['#', 'Account', 'Label', 'Partner', 'Debit', 'Credit', 'Balance']
    widths = [6, 22, 25, 20, 15, 15, 15]
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        style_header_cell(ws, r, c, h, w)

    # -- Loop start --
    r = 13
    style_loop_marker(ws, r, 1, '{{#line_ids}}')

    # -- Data row --
    r = 14
    placeholders = [
        ('{{sequence}}', 'center'),
        ('{{account_id.name}}', 'left'),
        ('{{name}}', 'left'),
        ('{{partner_id.name}}', 'left'),
        ('{{debit}}', 'right'),
        ('{{credit}}', 'right'),
        ('{{balance}}', 'right'),
    ]
    for c, (ph, align) in enumerate(placeholders, 1):
        style_data_cell(ws, r, c, ph, align)

    # -- Loop end --
    r = 15
    style_loop_marker(ws, r, 1, '{{/line_ids}}')

    # -- Totals --
    r = 17
    style_total_label(ws, r, 5, 'Total Debit:')
    style_total_value(ws, r, 6, '{{amount_total_signed}}')

    r = 18
    style_total_label(ws, r, 5, 'Amount Due:')
    style_total_value(ws, r, 6, '{{amount_residual}}')

    # -- Notes --
    r = 20
    ws.cell(row=r, column=1, value='Notes:').font = Font(bold=True, size=9, color='888888')
    r = 21
    ws.cell(row=r, column=1, value='{{narration}}').font = Font(size=9)
    ws.merge_cells('A21:G21')

    # -- Footer --
    r = 23
    ft = ws.cell(row=r, column=1, value='Generated by SM Report Designer')
    ft.font = Font(size=8, color='BBBBBB', italic=True)
    ws.merge_cells('A23:G23')

    # Print settings
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'

    wb.save('account_move_template.xlsx')
    print('Saved: account_move_template.xlsx')


# ==============================================================
#  SALE ORDER TEMPLATE
# ==============================================================
def create_sale_order_xlsx():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sales Order'

    # -- Title --
    r = 1
    title = ws.cell(row=r, column=1, value='{{company_id.name}}')
    title.font = Font(bold=True, size=16, color='2C3E50')
    ws.merge_cells('A1:G1')

    r = 2
    sub = ws.cell(row=r, column=1, value='SALES ORDER')
    sub.font = Font(bold=True, size=13, color='34495E')
    ws.merge_cells('A2:G2')

    r = 3
    ref = ws.cell(row=r, column=1, value='{{name}}')
    ref.font = Font(bold=True, size=11, color='2980B9')
    ws.merge_cells('A3:G3')

    r = 4  # spacer

    # -- Info section --
    info = [
        ('Customer', '{{partner_id.name}}', 'Date', '{{date_order}}'),
        ('Salesperson', '{{user_id.name}}', 'Payment Terms', '{{payment_term_id.name}}'),
        ('Reference', '{{client_order_ref}}', 'Company', '{{company_id.name}}'),
        ('Currency', '{{currency_id.name}}', 'Status', '{{state}}'),
    ]

    for i, (l1, v1, l2, v2) in enumerate(info):
        r = 5 + i
        style_label(ws, r, 1, l1)
        style_value(ws, r, 2, v1)
        style_label(ws, r, 4, l2)
        style_value(ws, r, 5, v2)

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 3
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 25

    r = 10  # spacer

    # -- Delivery address --
    r = 10
    ws.cell(row=r, column=1, value='Delivery Address:').font = Font(bold=True, size=9, color='555555')
    ws.cell(row=r, column=2, value='{{partner_shipping_id.name}}').font = Font(size=9)

    r = 12  # spacer

    # -- Order Lines header --
    r = 12
    headers = ['#', 'Product', 'Description', 'Quantity', 'Unit Price', 'Discount (%)', 'Subtotal']
    widths = [6, 22, 28, 12, 15, 14, 16]
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        style_header_cell(ws, r, c, h, w)

    # -- Loop start --
    r = 13
    style_loop_marker(ws, r, 1, '{{#order_line}}')

    # -- Data row --
    r = 14
    placeholders = [
        ('{{sequence}}', 'center'),
        ('{{product_id.name}}', 'left'),
        ('{{name}}', 'left'),
        ('{{product_uom_qty}}', 'right'),
        ('{{price_unit}}', 'right'),
        ('{{discount}}', 'right'),
        ('{{price_subtotal}}', 'right'),
    ]
    for c, (ph, align) in enumerate(placeholders, 1):
        style_data_cell(ws, r, c, ph, align)

    # -- Loop end --
    r = 15
    style_loop_marker(ws, r, 1, '{{/order_line}}')

    # -- Totals --
    r = 17
    style_total_label(ws, r, 5, 'Untaxed Amount:')
    ws.cell(row=r, column=5).border = Border(top=Side(style='double', color='2C3E50'))
    style_total_value(ws, r, 6, '{{amount_untaxed}}')
    # merge col 6-7 visually
    style_total_value(ws, r, 7, '')

    r = 18
    style_total_label(ws, r, 5, 'Taxes:')
    style_total_value(ws, r, 6, '{{amount_tax}}')

    r = 19
    style_total_label(ws, r, 5, 'Total:')
    c = ws.cell(row=r, column=6, value='{{amount_total}}')
    c.font = Font(bold=True, size=12, color='C0392B')
    c.alignment = Alignment(horizontal='right', vertical='center')
    c.border = Border(
        top=Side(style='thin', color='2C3E50'),
        bottom=Side(style='double', color='2C3E50'),
    )

    # -- Notes --
    r = 21
    ws.cell(row=r, column=1, value='Terms & Conditions:').font = Font(bold=True, size=9, color='888888')
    r = 22
    ws.cell(row=r, column=1, value='{{note}}').font = Font(size=9)
    ws.merge_cells('A22:G22')

    # -- Footer --
    r = 24
    ft = ws.cell(row=r, column=1, value='Generated by SM Report Designer')
    ft.font = Font(size=8, color='BBBBBB', italic=True)
    ws.merge_cells('A24:G24')

    # Print settings
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'

    wb.save('sale_order_template.xlsx')
    print('Saved: sale_order_template.xlsx')


if __name__ == '__main__':
    create_account_move_xlsx()
    create_sale_order_xlsx()
    print('\nDone! Upload these files to SM Report Designer.')
    print('\naccount_move_template.xlsx -> Model: Journal Entry (account.move)')
    print('sale_order_template.xlsx   -> Model: Sale Order (sale.order)')
