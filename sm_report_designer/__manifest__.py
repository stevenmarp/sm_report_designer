# -*- coding: utf-8 -*-
{
    'name': 'Report Designer - Word & Excel Template Engine',
    'version': '18.0.3.0.0',
    'category': 'Productivity',
    'summary': 'Design Odoo 18 reports from Word (.docx) and Excel (.xlsx) templates — no developer needed',
    'description': """
Report Designer - Word & Excel Template Engine (Odoo 18)
========================================================

Fully compatible with **Odoo 18** Community and Enterprise.

Let functional users create professional reports without coding.
Upload a Word or Excel template, map placeholders to Odoo fields,
and generate reports with one click.

Key Features
------------
* Upload .docx or .xlsx templates
* Auto-detect ``{{placeholder}}`` patterns
* Map placeholders to any Odoo model field via UI
* Table row loops with ``{{#field}}`` / ``{{/field}}`` markers
* Output as original format, PDF, HTML, ODT, or RTF (via LibreOffice)
* One-click print action on any Odoo model
* Multi-record batch generation (ZIP download)
* PDF merge — multiple records into a single PDF
* Domain-based template selection — auto-pick templates by record conditions
* PDF preview in browser — preview before downloading
* Direct URL endpoints for PDF preview and download

Advanced Features (v3.0)
------------------------
* **Python Expressions** — ``{{=amount_total * 0.1}}`` inline code
* **Conditional Blocks** — ``{{#if state == 'sale'}}…{{#else}}…{{/if}}``
* **Format Pipes** — ``{{amount|currency}}``, ``{{date_order|date}}``,
  ``{{name|upper}}``, ``{{total|number}}``
* **Image Support** — ``{{img:company_id.logo}}`` inserts binary images
* **Cover Page** — optional DOCX cover prepended to PDF output

How It Works
------------
1. Create a template in Word or Excel using ``{{field_name}}`` placeholders
2. Upload it to Odoo and select the target model
3. Click **Parse Template** to auto-detect placeholders
4. Map each placeholder to an Odoo field path
5. Click **Activate** to add a print button to the model
6. Users click the action menu to generate the report

Template Syntax
---------------
* ``{{name}}`` — replaced with the record's ``name`` field
* ``{{partner_id.name}}`` — dot notation for relational fields
* ``{{#order_line}}`` / ``{{/order_line}}`` — repeat table rows for One2many
* ``{{amount_total|currency}}`` — format as monetary value
* ``{{date_order|date}}`` — format as locale date
* ``{{=amount_untaxed + amount_tax}}`` — evaluate Python expression
* ``{{#if state == 'sale'}}`` … ``{{/if}}`` — conditional blocks
* ``{{img:company_id.logo}}`` — insert image from binary field
    """,
    'author': 'Steven Marp',
    'website': 'https://apps.odoo.com/apps/browse?repo_maintainer_id=512936',
    'license': 'OPL-1',
    'price': 99.00,
    'currency': 'USD',
    'depends': ['base', 'mail'],
    'data': [
        'security/security.xml',
        'security/ir.model.access.csv',
        'wizard/report_template_wizard_views.xml',
        'views/report_template_views.xml',
        'views/menu.xml',
    ],
    'images': ['static/description/banner.gif'],
    'installable': True,
    'application': True,
    'auto_install': False,
}
