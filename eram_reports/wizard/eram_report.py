# -*- coding: utf-8 -*-
from odoo import fields, models
from odoo.tools import json_default, html2plaintext
import io
import json
import xlsxwriter


class EramReport(models.TransientModel):
    _name = 'eram.report'
    _description = 'Eram Report'

    from_date = fields.Date(string="From Date")
    to_date = fields.Date(string="To Date")

    def action_print_report(self):
        mrp_orders = self.env['mrp.production'].search([])
        outwards_picking_ids = mrp_orders.picking_ids.ids

        domain = [('picking_type_id.code', '=', 'incoming')]

        inwards_picking_ids = self.env['stock.picking'].search(domain).ids

        data = {
            'inwards': inwards_picking_ids,
            'outwards': outwards_picking_ids,
        }

        return {
            'type': 'ir.actions.report',
            'report_name': 'eram.report_xlsx_inward_outward',
            'report_type': 'xlsx',
            'data': {
                'model': 'eram.report',
                'options': json.dumps(data, default=json_default),
                'output_format': 'xlsx',
                'report_name': 'Eram Inward & Outward Report',
            },
        }

    def get_xlsx_report(self, data, response):
        inwards_ids = data.get('inwards', [])
        outwards_ids = data.get('outwards', [])

        inwards_pickings = self.env['stock.picking'].sudo().browse(inwards_ids)
        outwards_pickings = self.env['stock.picking'].sudo().browse(outwards_ids)

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # ====================== INWARDS SHEET (unchanged) ======================
        ws_in = workbook.add_worksheet('Inwards')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAD3', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        center_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'})

        headers_in = [
            'SI No', 'GRN NO.', 'Invoice Date', 'Received Date', 'Project Code',
            'PR Number', 'PO Number', 'Invoice Number', 'Description', 'Part Number',
            'PO Qty', 'Received Qty', 'Unit', 'Rate', 'GST', 'Total Amount',
            'Grand Total', 'Supplier'
        ]
        for col, header in enumerate(headers_in):
            ws_in.write(0, col, header, header_format)

        row = 1
        si_no = 1
        grand_total_in = 0.0

        for picking in inwards_pickings:
            valid_moves = picking.move_ids.filtered(lambda m: m.state != 'cancel')
            if not valid_moves:
                continue
            for move in valid_moves:
                project_code = f"{picking.e_project_id.name or ''}-{picking.e_task_id.name or ''}".strip('-')
                po_number = picking.purchase_id.name if picking.purchase_id else (picking.e_po_no or '')
                invoice_number = picking.e_bill_id.name if picking.e_bill_id else ''
                description = html2plaintext(move.e_description or move.product_id.name or '').strip()
                part_no = html2plaintext(move.e_part_no or '').strip()

                gst_amount = (move.e_price_total or 0.0) - (move.e_total_untaxed or 0.0)
                line_total = move.e_price_total or 0.0
                grand_total_in += line_total

                col = 0
                ws_in.write(row, col, si_no, center_format);
                col += 1
                ws_in.write(row, col, picking.e_grn_id.name or '', cell_format);
                col += 1
                ws_in.write(row, col, picking.e_invoice_date or '', cell_format);
                col += 1
                ws_in.write(row, col, picking.e_invoice_received_date or '', cell_format);
                col += 1
                ws_in.write(row, col, project_code, cell_format);
                col += 1
                ws_in.write(row, col, picking.e_pr_id.pr_number if picking.e_pr_id else '', cell_format);
                col += 1
                ws_in.write(row, col, po_number, cell_format);
                col += 1
                ws_in.write(row, col, invoice_number, cell_format);
                col += 1
                ws_in.write(row, col, description, cell_format);
                col += 1
                ws_in.write(row, col, part_no, cell_format);
                col += 1
                ws_in.write(row, col, move.product_uom_qty, money_format);
                col += 1
                ws_in.write(row, col, move.quantity, money_format);
                col += 1
                ws_in.write(row, col, move.product_uom.name or '', center_format);
                col += 1
                ws_in.write(row, col, move.price_unit or 0.0, money_format);
                col += 1
                ws_in.write(row, col, gst_amount, money_format);
                col += 1
                ws_in.write(row, col, line_total, money_format);
                col += 1
                ws_in.write(row, col, '', cell_format);
                col += 1
                ws_in.write(row, col, picking.partner_id.name or '', cell_format)

                row += 1
                si_no += 1

        if row > 1:
            row += 1
            ws_in.write(row, 14, 'GRAND TOTAL',
                        workbook.add_format({'bold': True, 'border': 1, 'align': 'right', 'bg_color': '#E0E0E0'}))
            ws_in.write(row, 16, grand_total_in, workbook.add_format(
                {'bold': True, 'border': 1, 'align': 'center', 'num_format': '#,##0.00', 'bg_color': '#E0E0E0'}))

        widths_in = [6, 15, 12, 12, 20, 12, 14, 14, 35, 16, 10, 10, 8, 12, 12, 14, 14, 28]
        for col, w in enumerate(widths_in):
            ws_in.set_column(col, col, w)

        # ────────────────────────────────────────────────
        # SHEET 2: Outwards - Corrected Values + Proper Merging
        # ────────────────────────────────────────────────
        ws_out = workbook.add_worksheet('Outwards')

        # Formats
        out_header_format = workbook.add_format({
            'bold': True, 'bg_color': '#E6F3FF', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })

        out_cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        out_center_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        out_money_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'align': 'center', 'valign': 'vcenter',
        })

        merge_format = workbook.add_format({
            'border': 1, 'valign': 'vcenter', 'align': 'left',
        })

        merge_center_format = workbook.add_format({
            'border': 1, 'valign': 'vcenter', 'align': 'center',
        })

        grand_label_format = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'right', 'valign': 'vcenter', 'bg_color': '#E0E0E0',
        })

        grand_value_format = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#E0E0E0', 'num_format': '#,##0.00',
        })

        headers_out = [
            'SL No', 'MWR NO.', 'Date', 'Project Code',
            'Employee Name', 'Description', 'Part No',
            'Qty', 'Unit', 'Rate', 'GST', 'Total Amount', 'Justification/Remarks'
        ]

        for col, header in enumerate(headers_out):
            ws_out.write(0, col, header, out_header_format)

        row_out = 1
        sl_no = 1
        grand_total_out = 0.0

        for picking in outwards_pickings:
            if not picking.move_ids:
                continue

            project_code = f"{picking.e_project_id.name or ''}-{picking.e_task_id.name or ''}".strip('-')
            employee_name = picking.e_requested_by.name if picking.e_requested_by else ''
            mwr_no = picking.name or ''

            # Proper date handling - convert to Excel date number
            date_val = picking.e_date or picking.date_done or picking.scheduled_date or fields.Date.today()
            if hasattr(date_val, 'strftime'):
                date_str = date_val.strftime('%Y-%m-%d')
            else:
                date_str = str(date_val)

            valid_moves = picking.move_ids.filtered(lambda m: m.state != 'cancel')

            for move in valid_moves:
                description = html2plaintext(move.e_description_out or move.product_id.name or '').strip()
                part_no = html2plaintext(move.e_part_no or '').strip()
                unit_name = move.product_uom.name or ''
                remarks = html2plaintext(move.e_remarks or '').strip()

                valuation_layers = move.eram_valuation_ids.sorted(key=lambda l: (l.create_date or l.id or 0))

                group_start_row = row_out

                if not valuation_layers:
                    # Single row fallback
                    qty = abs(move.product_uom_qty or 0.0)

                    ws_out.write(row_out, 0, sl_no, out_center_format)
                    ws_out.write(row_out, 1, mwr_no, out_cell_format)
                    ws_out.write(row_out, 2, date_str, out_cell_format)
                    ws_out.write(row_out, 3, project_code, out_cell_format)
                    ws_out.write(row_out, 4, employee_name, out_cell_format)
                    ws_out.write(row_out, 5, description, out_cell_format)
                    ws_out.write(row_out, 6, part_no, out_cell_format)
                    ws_out.write(row_out, 7, qty, out_center_format)  # Qty column
                    ws_out.write(row_out, 8, unit_name, out_center_format)  # Unit column
                    ws_out.write(row_out, 9, 0.0, out_money_format)  # Rate
                    ws_out.write(row_out, 10, 0.0, out_money_format)  # GST
                    ws_out.write(row_out, 11, 0.0, out_money_format)  # Total Amount
                    ws_out.write(row_out, 12, remarks, out_cell_format)

                    row_out += 1
                    sl_no += 1
                    continue

                # Multiple valuation layers
                for i, layer in enumerate(valuation_layers):
                    layer_qty = abs(layer.quantity or 0.0)
                    total_taxed = abs(layer.total_taxed or 0.0)
                    value = abs(layer.value or 0.0)
                    gst_amount = total_taxed - value
                    unit_rate = abs(layer.unit_cost or 0.0)

                    grand_total_out += total_taxed

                    if i == 0:
                        # First row of group - write SL No and common fields
                        ws_out.write(row_out, 0, sl_no, out_center_format)
                        ws_out.write(row_out, 1, mwr_no, out_cell_format)
                        ws_out.write(row_out, 2, date_str, out_cell_format)
                        ws_out.write(row_out, 3, project_code, out_cell_format)
                        ws_out.write(row_out, 4, employee_name, out_cell_format)
                        ws_out.write(row_out, 5, description, out_cell_format)
                        ws_out.write(row_out, 6, part_no, out_cell_format)
                    else:
                        # For subsequent rows, leave SL No and common fields blank (will be merged)
                        ws_out.write(row_out, 0, '', out_center_format)
                        ws_out.write(row_out, 1, '', out_cell_format)
                        ws_out.write(row_out, 2, '', out_cell_format)
                        ws_out.write(row_out, 3, '', out_cell_format)
                        ws_out.write(row_out, 4, '', out_cell_format)
                        ws_out.write(row_out, 5, '', out_cell_format)
                        ws_out.write(row_out, 6, '', out_cell_format)

                    # Per-layer fields - Qty, Unit, Rate, GST, Total Amount
                    ws_out.write(row_out, 7, layer_qty, out_center_format)  # Qty column
                    ws_out.write(row_out, 8, unit_name, out_center_format)  # Unit column
                    ws_out.write(row_out, 9, unit_rate, out_money_format)  # Rate
                    ws_out.write(row_out, 10, gst_amount, out_money_format)  # GST
                    ws_out.write(row_out, 11, total_taxed, out_money_format)  # Total Amount

                    # Remarks only on first row
                    if i == 0:
                        ws_out.write(row_out, 12, remarks, out_cell_format)
                    else:
                        ws_out.write(row_out, 12, '', out_cell_format)

                    row_out += 1

                # Merge common columns if multiple layers
                last_row = row_out - 1
                if group_start_row < last_row:
                    # Merge columns 0-6 (SL No through Part No)
                    ws_out.merge_range(group_start_row, 0, last_row, 0, sl_no, merge_center_format)
                    ws_out.merge_range(group_start_row, 1, last_row, 1, mwr_no, merge_format)
                    ws_out.merge_range(group_start_row, 2, last_row, 2, date_str, merge_format)
                    ws_out.merge_range(group_start_row, 3, last_row, 3, project_code, merge_format)
                    ws_out.merge_range(group_start_row, 4, last_row, 4, employee_name, merge_format)
                    ws_out.merge_range(group_start_row, 5, last_row, 5, description, merge_format)
                    ws_out.merge_range(group_start_row, 6, last_row, 6, part_no, merge_format)
                    # Merge remarks column if there are multiple rows
                    if remarks:
                        ws_out.merge_range(group_start_row, 12, last_row, 12, remarks, merge_format)

                sl_no += 1

        # Grand Total
        if row_out > 1:
            row_out += 1
            ws_out.write(row_out, 10, 'GRAND TOTAL', grand_label_format)
            ws_out.write(row_out, 11, grand_total_out, grand_value_format)

        # Column widths
        widths_out = [8, 22, 14, 20, 22, 40, 18, 12, 10, 14, 14, 16, 45]
        for col, width in enumerate(widths_out):
            ws_out.set_column(col, col, width)

        # Finalize
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()