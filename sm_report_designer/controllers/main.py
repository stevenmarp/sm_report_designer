# -*- coding: utf-8 -*-

import base64
import logging

from odoo import http
from odoo.http import request, content_disposition

_logger = logging.getLogger(__name__)


class SmReportDesignerController(http.Controller):

    @http.route(
        '/report_designer/preview/<int:template_id>/<int:record_id>',
        type='http', auth='user',
    )
    def preview_pdf(self, template_id, record_id, **kwargs):
        """Render and return a PDF for inline browser preview."""
        template = request.env['sm.report.template'].browse(template_id)
        if not template.exists():
            return request.not_found()

        record = request.env[template.model_name].browse(record_id)
        if not record.exists():
            return request.not_found()

        file_bytes = base64.b64decode(template.template_file)
        ctx = template._build_render_context(record)

        if template.template_type == 'docx':
            output = template._render_docx(file_bytes, ctx)
        elif template.template_type == 'xlsx':
            output = template._render_xlsx(file_bytes, ctx)
        else:
            return request.not_found()

        pdf_bytes = template._convert_format(
            output, template.template_type, 'pdf',
        )

        filename = '%s - %s.pdf' % (
            template.name, record.display_name or record.id,
        )

        return request.make_response(
            pdf_bytes,
            headers=[
                ('Content-Type', 'application/pdf'),
                ('Content-Disposition', 'inline; filename="%s"' % filename),
                ('Content-Length', len(pdf_bytes)),
            ],
        )

    @http.route(
        '/report_designer/download/<int:template_id>/<int:record_id>',
        type='http', auth='user',
    )
    def download_report(self, template_id, record_id, **kwargs):
        """Generate and download the report in its configured output format."""
        template = request.env['sm.report.template'].browse(template_id)
        if not template.exists():
            return request.not_found()

        record = request.env[template.model_name].browse(record_id)
        if not record.exists():
            return request.not_found()

        file_bytes = base64.b64decode(template.template_file)
        ctx = template._build_render_context(record)

        if template.template_type == 'docx':
            output = template._render_docx(file_bytes, ctx)
        elif template.template_type == 'xlsx':
            output = template._render_xlsx(file_bytes, ctx)
        else:
            return request.not_found()

        if template.output_format not in ('original', False):
            output = template._convert_format(
                output, template.template_type, template.output_format,
            )
            ext = template.output_format
        else:
            ext = template.template_type

        mime_map = {
            'pdf': 'application/pdf',
            'docx': 'application/vnd.openxmlformats-officedocument'
                    '.wordprocessingml.document',
            'xlsx': 'application/vnd.openxmlformats-officedocument'
                    '.spreadsheetml.sheet',
            'odt': 'application/vnd.oasis.opendocument.text',
            'html': 'text/html',
            'rtf': 'application/rtf',
        }
        content_type = mime_map.get(ext, 'application/octet-stream')

        filename = '%s - %s.%s' % (
            template.name, record.display_name or record.id, ext,
        )

        return request.make_response(
            output,
            headers=[
                ('Content-Type', content_type),
                ('Content-Disposition', content_disposition(filename)),
                ('Content-Length', len(output)),
            ],
        )
