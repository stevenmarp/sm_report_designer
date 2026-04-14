=========================================================
Report Designer — Word & Excel Template Engine for Odoo 18
=========================================================

.. contents:: Table of Contents
   :depth: 3
   :local:

----

1. Getting Started
==================

Prerequisites
-------------

+----------------------------+----------------------------------------------------------------+
| Requirement                | Description                                                    |
+============================+================================================================+
| Odoo Version               | 19.0 (Community or Enterprise)                                 |
+----------------------------+----------------------------------------------------------------+
| LibreOffice                | Required for PDF/HTML/ODT/RTF conversion. Install with         |
|                            | ``sudo apt install libreoffice`` on Ubuntu/Debian.             |
+----------------------------+----------------------------------------------------------------+
| PyPDF2 or pypdf            | Optional. Required for PDF merge feature.                      |
|                            | ``pip install pypdf`` or ``pip install PyPDF2``                 |
+----------------------------+----------------------------------------------------------------+
| Pillow (PIL)               | Optional. Used for image size detection. Falls back to         |
|                            | struct-based detection if not available.                       |
+----------------------------+----------------------------------------------------------------+

Installation
------------

1. Upload the ``sm_report_designer`` module folder to your server's **custom addons directory**.
2. **Restart** the Odoo service to detect the new module.
3. Navigate to **Apps** menu, click **Update Apps List**.
4. Search for **"Report Designer"** and click **Install**.

.. note::
   After installation, a new menu item appears under
   **Settings → Technical → Reporting → Report Designer**.

----

2. Creating Your First Template
===============================

Step 1: Design the Template
----------------------------

Open Microsoft Word, LibreOffice Writer, or Google Docs and create your report layout.
Use ``{{field_name}}`` placeholders wherever you want Odoo data to appear.

Example Word template::

    SALES ORDER: {{name}}
    Customer: {{partner_id}}
    Date: {{date_order}}

    Total: {{amount_total}}

Save the file as ``.docx`` (or ``.xlsx`` for spreadsheet reports).

Step 2: Upload & Configure
---------------------------

1. Go to **Settings → Technical → Reporting → Report Designer**
2. Click **Create**
3. Fill in:

   - **Name**: e.g., "Sales Order Quotation"
   - **Model**: Select the target Odoo model (e.g., ``sale.order``)
   - **Template File**: Upload your ``.docx`` or ``.xlsx`` file
   - **Output Format**: Choose PDF, Original, HTML, ODT, or RTF

Step 3: Parse & Map Fields
---------------------------

1. Go to the **Field Mappings** tab
2. Click the **Parse Template** button
3. All ``{{placeholder}}`` patterns are auto-detected
4. For each placeholder, select the correct Odoo field from the dropdown
5. Use dot notation for relational fields: ``partner_id.name``, ``company_id.logo``

Step 4: Activate
-----------------

1. Click the **Activate** button at the top of the form
2. A new action is registered on the target model
3. Users can now find the report in the model's **Action** menu

Step 5: Generate Report
------------------------

1. Open any record of the target model (e.g., a Sale Order)
2. Click **Action** menu → select your report template
3. The report is generated and downloaded automatically

----

3. Template Syntax
==================

Simple Placeholders
-------------------

Replace with the value of a field on the current record::

    {{name}}                    → "S00042"
    {{partner_id}}              → "Azure Interior"
    {{amount_total}}            → "15000000.5"

Dot Notation (Relational Fields)
---------------------------------

Access fields on related records using dots::

    {{partner_id.name}}         → "Azure Interior"
    {{partner_id.email}}        → "azure@example.com"
    {{company_id.name}}         → "My Company"
    {{partner_id.country_id.name}} → "Indonesia"

Format Pipes
------------

Append ``|format`` to format values inline::

    {{amount_total|currency}}   → "Rp 15,000,000.50" (locale-aware)
    {{date_order|date}}         → "04/14/2026" (locale-aware)
    {{create_date|datetime}}    → "04/14/2026 10:30:00"
    {{amount_untaxed|number}}   → "12,000,000.00"
    {{name|upper}}              → "S00042"
    {{name|lower}}              → "s00042"
    {{partner_id|title}}        → "Azure Interior"

Python Expressions
------------------

Prefix with ``=`` to evaluate Python expressions::

    {{=amount_untaxed + amount_tax}}        → sum of two fields
    {{=round(amount_total * 0.1, 2)}}       → 10% rounded
    {{=amount_untaxed * 0.11}}              → PPN 11%
    {{='PAID' if state == 'sale' else 'DRAFT'}}

Available variables in expressions:

- All record fields (by name)
- ``record`` — the current Odoo record object
- ``user`` — current user (``self.env.user``)
- ``env`` — the Odoo environment
- ``time``, ``datetime`` — Python standard library modules

.. warning::
   Expressions are evaluated using Odoo's ``safe_eval`` for security.
   Dangerous operations (file I/O, imports, etc.) are blocked.

Conditional Blocks
------------------

Show or hide paragraphs/rows based on conditions::

    {{#if amount_tax}}
    Tax Amount: {{amount_tax|currency}}
    This order includes tax.
    {{#else}}
    This order is tax-free.
    {{/if}}

**Important rules:**

- Each ``{{#if ...}}``, ``{{#else}}``, and ``{{/if}}`` must be on its own
  paragraph (DOCX) or row (XLSX).
- Do NOT mix conditional tags with other content on the same line.
- Nesting is supported::

    {{#if partner_id}}
    Customer: {{partner_id}}
        {{#if amount_tax}}
        Tax: {{amount_tax}}
        {{/if}}
    {{/if}}

- Conditions can be simple field names (truthy check) or expressions::

    {{#if amount_tax}}           → true if amount_tax is non-zero/non-empty
    {{#if state == 'sale'}}      → expression evaluation

Table Row Loops
---------------

Repeat table rows for One2many/Many2many fields::

    +-----------+----------------+--------+-----------+
    | No        | Product        | Qty    | Subtotal  |
    +-----------+----------------+--------+-----------+
    | {{#order_line}}                                  |
    +-----------+----------------+--------+-----------+
    | {{sequence}} | {{product_id}} | {{product_uom_qty}} | {{price_subtotal}} |
    +-----------+----------------+--------+-----------+
    | {{/order_line}}                                  |
    +-----------+----------------+--------+-----------+

The rows between ``{{#order_line}}`` and ``{{/order_line}}`` are repeated
for each record in the ``order_line`` One2many field.

**Setup:** In Field Mappings, map ``order_line`` as a **loop field** pointing
to the One2many relation (e.g., ``order_line``). Map sub-fields (``sequence``,
``product_id``, etc.) within the loop context.

Image Insertion
---------------

Insert binary field images with the ``img:`` prefix::

    {{img:company_id.logo}}         → company logo
    {{img:partner_id.image_128}}    → partner avatar
    {{img:product_id.image_256}}    → product image

**DOCX behavior:**

- Images are embedded as inline drawings in the document
- Auto-resized to max 15cm × 10cm with preserved aspect ratio
- Supports PNG, JPEG, GIF, WEBP, BMP formats

**XLSX behavior:**

- Images are inserted using openpyxl's Image class
- Anchored to the cell where the placeholder was

----

4. Advanced Features
====================

PDF Merge
---------

Combine multiple records into a single PDF:

1. Open your template settings
2. Check the **Merge PDF** checkbox
3. Set **Output Format** to **PDF**
4. In a list view, select multiple records
5. Click **Action** → your template
6. One merged PDF is downloaded

Cover Page
----------

Add a professional cover page to your PDF:

1. Set **Output Format** to **PDF** (the Cover Page tab appears)
2. Go to the **Cover Page** tab
3. Upload a ``.docx`` file as the cover template
4. Use the same ``{{placeholder}}`` syntax — the cover is rendered with the same record data
5. The cover page is automatically prepended to the PDF output

Domain-Based Templates
-----------------------

Create multiple templates for the same model with different conditions:

- Template A: Domain ``[('amount_total', '>', 10000000)]`` → for large orders
- Template B: Domain ``[('amount_total', '<=', 10000000)]`` → for small orders

The system auto-selects the matching template when generating a report.

PDF Preview
-----------

Preview reports in the browser before downloading:

- Click the **Preview** button on the template form
- Or access directly: ``/report_designer/preview/<template_id>/<record_id>``
- PDF is displayed inline in the browser (not downloaded)

----

5. Output Formats
=================

+----------+-------------------+--------------------------------------------------+
| Format   | File Extension    | Notes                                            |
+==========+===================+==================================================+
| Original | .docx / .xlsx     | Keeps the template format as-is                  |
+----------+-------------------+--------------------------------------------------+
| PDF      | .pdf              | Requires LibreOffice. Supports merge & cover     |
+----------+-------------------+--------------------------------------------------+
| HTML     | .html             | Requires LibreOffice for conversion              |
+----------+-------------------+--------------------------------------------------+
| ODT      | .odt              | OpenDocument format. Requires LibreOffice        |
+----------+-------------------+--------------------------------------------------+
| RTF      | .rtf              | Rich Text Format. Requires LibreOffice           |
+----------+-------------------+--------------------------------------------------+

----

6. Troubleshooting
===================

LibreOffice not found
----------------------

If PDF conversion fails with "LibreOffice not found":

1. Install LibreOffice: ``sudo apt install libreoffice``
2. Verify: ``libreoffice --version``
3. Ensure the ``libreoffice`` binary is in the system PATH

Placeholders not replaced
--------------------------

- Make sure you clicked **Parse Template** after uploading
- Check that each placeholder is mapped to a valid field
- For relational fields, use dot notation: ``partner_id.name`` (not just ``partner_id``)
- Verify the field exists on the selected model

Images not showing
-------------------

- Ensure the binary field contains actual image data (not empty)
- Check the field path: ``company_id.logo`` (not ``company_id.logo_web``)
- Supported formats: PNG, JPEG, GIF, WEBP, BMP
- Maximum size: 15cm × 10cm (auto-scaled with aspect ratio)

Conditional blocks not working
-------------------------------

- Each ``{{#if}}``, ``{{#else}}``, ``{{/if}}`` must be on its **own paragraph** (DOCX)
  or **own row** (XLSX)
- Do NOT put other text on the same line as a conditional tag
- Check for typos: ``{{#if field}}`` not ``{{# if field}}``

----

7. Changelog
============

Version 19.0.3.0.0
-------------------

- Added Python expressions ``{{=expr}}``
- Added conditional blocks ``{{#if}}/{{#else}}/{{/if}}``
- Added format pipes ``{{field|format}}`` (currency, date, datetime, number, upper, lower, title)
- Added image support ``{{img:field}}`` for DOCX and XLSX
- Added cover page support (PDF only)

Version 19.0.2.0.0
-------------------

- Added PDF merge for multiple records
- Added domain-based template selection
- Added PDF preview in browser
- Added direct URL endpoints for preview and download

Version 19.0.1.0.0
-------------------

- Initial release
- Word (.docx) and Excel (.xlsx) template support
- Auto-detect placeholders
- Field mapping via UI
- Table row loops for One2many fields
- Multi-format output (PDF, HTML, ODT, RTF)
- One-click print action
- Batch generation with ZIP download
