# -*- coding: utf-8 -*-
from odoo import fields, models
from odoo.tools import json_default
import io
import json
import xlsxwriter


class EramReport(models.TransientModel):
    _name = 'eram.report'
    _description = 'Eram Report'

    from_date = fields.Date(string="From Date")
    to_date = fields.Date(string="To Date")

    def action_print_report(self):
        domain = [('picking_type_id.code', '=', 'incoming')]

        picking_ids = self.env['stock.picking'].search(domain).ids

        data = {'inwards': picking_ids}

        return {
            'type': 'ir.actions.report',
            'report_name': 'eram.report_xlsx_inward',
            'report_type': 'xlsx',
            'data': {
                'model': 'eram.report',
                'options': json.dumps(data, default=json_default),
                'output_format': 'xlsx',
                'report_name': 'Eram Inward Report',
            },
        }

    def get_xlsx_report(self, data, response):
        picking_ids = data.get('inwards', [])
        pickings = self.env['stock.picking'].sudo().browse(picking_ids)

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Inward Report')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9EAD3',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
        })

        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        center_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'})

        merge_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#FFF2CC',
            'num_format': '#,##0.00',
        })

        grand_label_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#E0E0E0',
        })

        grand_value_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#E0E0E0',
            'num_format': '#,##0.00',
        })

        headers = [
            'SI No', 'GRN NO.', 'Invoice Date', 'Received Date', 'Project Code',
            'PR Number', 'PO Number', 'Invoice Number', 'Description', 'Part Number',
            'PO Qty', 'Received Qty', 'Unit', 'Rate', 'GST',
            'Total Amount', 'Grand Total', 'Supplier'
        ]

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        row = 1
        si_no = 1
        grand_total = 0.0

        for picking in pickings:
            valid_moves = picking.move_ids.filtered(lambda m: m.state != 'cancel')
            if not valid_moves:
                continue

            group_start_row = row

            picking_total = 0.0

            for move in valid_moves:
                project_code = f"{picking.e_project_id.name or ''}-{picking.e_task_id.name or ''}".strip('-')
                po_number = picking.purchase_id.name if picking.purchase_id else (picking.e_po_no or '')
                invoice_number = picking.e_bill_id.name if picking.e_bill_id else ''
                gst_amount = (move.e_price_total or 0.0) - (move.e_total_untaxed or 0.0)
                line_total = move.e_price_total or 0.0

                picking_total += line_total
                grand_total += line_total

                col = 0
                worksheet.write(row, col, si_no, center_format); col += 1
                worksheet.write(row, col, picking.e_grn_id.name or '', cell_format); col += 1
                worksheet.write(row, col, picking.e_invoice_date or '', cell_format); col += 1
                worksheet.write(row, col, picking.date_done.date() if picking.date_done else '', cell_format); col += 1
                worksheet.write(row, col, project_code, cell_format); col += 1
                worksheet.write(row, col, picking.e_pr_id.pr_number if picking.e_pr_id else '', cell_format); col += 1
                worksheet.write(row, col, po_number, cell_format); col += 1
                worksheet.write(row, col, invoice_number, cell_format); col += 1
                worksheet.write(row, col, move.e_description or move.product_id.name or '', cell_format); col += 1
                worksheet.write(row, col, move.e_part_no or '', cell_format); col += 1
                worksheet.write(row, col, move.product_uom_qty, money_format); col += 1
                worksheet.write(row, col, move.quantity, money_format); col += 1
                worksheet.write(row, col, move.product_uom.name or '', center_format); col += 1
                worksheet.write(row, col, move.price_unit or 0.0, money_format); col += 1
                worksheet.write(row, col, gst_amount, money_format); col += 1
                worksheet.write(row, col, line_total, money_format); col += 1
                worksheet.write(row, col, '', cell_format); col += 1
                worksheet.write(row, col, picking.partner_id.name or '', cell_format)

                row += 1
                si_no += 1

            group_end_row = row - 1
            if group_start_row <= group_end_row:
                total_value = picking.e_amount_total or picking_total or 0.0

                worksheet.merge_range(
                    group_start_row, 16,
                    group_end_row, 16,
                    total_value,
                    merge_format
                )

                grand_total += total_value
        if row > 1:
            row += 1
            worksheet.write(row, 14, 'GRAND TOTAL', grand_label_format)
            worksheet.write(row, 15, '', cell_format)
            worksheet.write(row, 16, grand_total, grand_value_format)
            worksheet.write(row, 17, '', cell_format)

        widths = [6, 15, 12, 12, 20, 12, 14, 14, 35, 16, 10, 10, 8, 12, 12, 14, 14, 28]
        for col, width in enumerate(widths):
            worksheet.set_column(col, col, width)

        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()