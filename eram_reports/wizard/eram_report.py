# -*- coding: utf-8 -*-
from odoo import api, fields, models
from odoo.tools import json_default, html2plaintext
import io
import json
import xlsxwriter
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar


class EramReport(models.TransientModel):
    _name = 'eram.report'
    _description = 'Eram Report'

    year = fields.Selection(
        string='Year',
        selection='_get_year',
        required=True,
        default=lambda self: str(datetime.now().year)
    )

    month = fields.Selection(
        string='Month',
        selection='_get_month',
        required=True,
        default=lambda self: str((datetime.now().replace(day=1) - relativedelta(months=1)).month)
    )

    @api.model
    def _get_year(self):
        current_year = datetime.now().year
        return [(str(y), str(y)) for y in range(2021, current_year + 1)]

    @api.model
    def _get_month(self):
        return [(str(m), calendar.month_name[m]) for m in range(1, 13)]

    def action_print_report(self):
        year = int(self.year)
        month = int(self.month)

        first_day = datetime(year, month, 1)
        if month == 12:
            last_day = datetime(year + 1, 1, 1) - relativedelta(days=1)
        else:
            last_day = datetime(year, month + 1, 1) - relativedelta(days=1)

        first_day_str = first_day.strftime('%Y-%m-%d')
        last_day_str = last_day.strftime('%Y-%m-%d')

        mrp_orders = self.env['mrp.production'].search([])
        all_outwards_picking_ids = mrp_orders.picking_ids.ids

        domain_in = [
            ('picking_type_id.code', '=', 'incoming'),
            ('e_invoice_date', '>=', first_day_str),
            ('e_invoice_date', '<=', last_day_str),
        ]
        inwards_picking_ids = self.env['stock.picking'].search(domain_in).ids

        domain_out = [
            ('id', 'in', all_outwards_picking_ids),
            ('e_date', '>=', first_day_str),
            ('e_date', '<=', last_day_str),
        ]
        filtered_outwards_picking_ids = self.env['stock.picking'].search(domain_out).ids

        data = {
            'inwards': inwards_picking_ids,
            'outwards': filtered_outwards_picking_ids,
            'year': self.year,
            'month': self.month,
        }

        return {
            'type': 'ir.actions.report',
            'report_name': 'eram.report_xlsx_inward_outward',
            'report_type': 'xlsx',
            'data': {
                'model': 'eram.report',
                'options': json.dumps(data, default=json_default),
                'output_format': 'xlsx',
                'report_name': f'Eram Inward & Outward Report - {calendar.month_name[month]} {year}',
            },
        }

    # ─────────────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _get_project_key(self, picking):
        """Return 'ProjectName_TaskName' (strips trailing/leading dashes)."""
        proj = picking.e_project_id.name or ''
        task = picking.e_task_id.name or ''
        if proj and task:
            return f"{proj}_{task}"
        return proj or task or 'UNKNOWN'

    def _compute_opening_balance(self, year, month):
        """
        Opening balance = total inward amount received BEFORE the selected month
                        - total outward amount issued BEFORE the selected month.
        Returns (opening_balance_total, {project_key: opening_balance}).
        """
        cutoff_str = datetime(year, month, 1).strftime('%Y-%m-%d')

        # --- Inwards before cutoff ---
        domain_in = [
            ('picking_type_id.code', '=', 'incoming'),
            ('e_invoice_date', '<', cutoff_str),
        ]
        prev_inwards = self.env['stock.picking'].sudo().search(domain_in)

        total_in = 0.0
        proj_in = {}
        for picking in prev_inwards:
            for move in picking.move_ids.filtered(lambda m: m.state != 'cancel'):
                amt = move.e_price_total or 0.0
                total_in += amt
                key = self._get_project_key(picking)
                proj_in[key] = proj_in.get(key, 0.0) + amt

        # --- Outwards before cutoff ---
        mrp_orders = self.env['mrp.production'].search([])
        all_out_ids = mrp_orders.picking_ids.ids

        domain_out = [
            ('id', 'in', all_out_ids),
            ('e_date', '<', cutoff_str),
        ]
        prev_outwards = self.env['stock.picking'].sudo().search(domain_out)

        total_out = 0.0
        proj_out = {}
        for picking in prev_outwards:
            for move in picking.move_ids.filtered(lambda m: m.state != 'cancel'):
                for layer in move.eram_valuation_ids:
                    amt = abs(layer.total_taxed or 0.0)
                    total_out += amt
                    key = self._get_project_key(picking)
                    proj_out[key] = proj_out.get(key, 0.0) + amt

        opening_total = total_in - total_out

        all_projects = set(list(proj_in.keys()) + list(proj_out.keys()))
        proj_opening = {k: proj_in.get(k, 0.0) - proj_out.get(k, 0.0) for k in all_projects}

        return opening_total, proj_opening

    def _compute_period_inward(self, inwards_pickings):
        """Return (total, {project_key: total}) for the selected-month inwards."""
        total = 0.0
        by_proj = {}
        for picking in inwards_pickings:
            for move in picking.move_ids.filtered(lambda m: m.state != 'cancel'):
                amt = move.e_price_total or 0.0
                total += amt
                key = self._get_project_key(picking)
                by_proj[key] = by_proj.get(key, 0.0) + amt
        return total, by_proj

    def _compute_period_outward(self, outwards_pickings):
        """Return (total, {project_key: total}) for the selected-month outwards."""
        total = 0.0
        by_proj = {}
        for picking in outwards_pickings:
            for move in picking.move_ids.filtered(lambda m: m.state != 'cancel'):
                for layer in move.eram_valuation_ids:
                    amt = abs(layer.total_taxed or 0.0)
                    total += amt
                    key = self._get_project_key(picking)
                    by_proj[key] = by_proj.get(key, 0.0) + amt
        return total, by_proj

    # ─────────────────────────────────────────────────────────────────────────
    # Main report generator
    # ─────────────────────────────────────────────────────────────────────────

    def get_xlsx_report(self, data, response):
        inwards_ids = data.get('inwards', [])
        outwards_ids = data.get('outwards', [])
        year = int(data.get('year', datetime.now().year))
        month = int(data.get('month', datetime.now().month))

        month_name = calendar.month_name[month]
        period_title = f"{month_name} {year}"

        # last day of selected month
        if month == 12:
            last_day_of_month = datetime(year + 1, 1, 1) - relativedelta(days=1)
        else:
            last_day_of_month = datetime(year, month + 1, 1) - relativedelta(days=1)
        last_day_str = last_day_of_month.strftime('%B %dth %Y')

        # previous month label  (e.g. "December 31st 2025")
        prev_month_last = datetime(year, month, 1) - relativedelta(days=1)
        prev_month_label = prev_month_last.strftime('%B %dst %Y') if prev_month_last.day == 1 else \
                           prev_month_last.strftime('%B %dth %Y')

        inwards_pickings = self.env['stock.picking'].sudo().browse(inwards_ids)
        outwards_pickings = self.env['stock.picking'].sudo().browse(outwards_ids)

        # Compute aggregated figures
        opening_total, proj_opening = self._compute_opening_balance(year, month)
        inward_total, proj_inward = self._compute_period_inward(inwards_pickings)
        outward_total, proj_outward = self._compute_period_outward(outwards_pickings)
        closing_total = opening_total + inward_total - outward_total

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # ══════════════════════════════════════════════════════════════════════
        # COMMON FORMATS
        # ══════════════════════════════════════════════════════════════════════
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
        })
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAD3', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        center_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'})
        bold_border_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'valign': 'vcenter', 'bg_color': '#E0E0E0',
        })
        bold_money_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'num_format': '#,##0.00',
            'valign': 'vcenter', 'bg_color': '#E0E0E0',
        })
        sign_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
        })
        label_right_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'right', 'valign': 'vcenter', 'bg_color': '#E0E0E0',
        })

        # ══════════════════════════════════════════════════════════════════════
        # SHEET 1 – INDEX
        # ══════════════════════════════════════════════════════════════════════
        ws_idx = workbook.add_worksheet('Index')

        # Row 0: company heading
        ws_idx.merge_range(0, 5, 0, 7, 'EPEC COE (BANGALORE) - ' + str(year),
                           workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center'}))

        # Row 1: period heading
        period_heading = f'Bangalore Inventory Stock  for {month_name} - {year}'
        ws_idx.merge_range(1, 4, 1, 8, period_heading,
                           workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center'}))

        # Row 3: table headers
        idx_headers = ['Sl No.', 'Description', 'Total Amount']
        for col, h in enumerate(idx_headers):
            ws_idx.write(3, col + 5, h, header_fmt)

        # Data rows
        rows = [
            (1, f'Opening Balance Until {prev_month_label}', opening_total),
            (2, f'{month_name} {year}  Inward (+)',          inward_total),
            (3, f'{month_name} {year}  Outward (-)',         outward_total),
        ]
        for i, (sl, desc, amt) in enumerate(rows):
            ws_idx.write(4 + i, 5, sl,   center_fmt)
            ws_idx.write(4 + i, 6, desc, cell_fmt)
            ws_idx.write(4 + i, 7, amt,  money_fmt)

        # Closing balance row (spans Sl+Desc, value in Total Amount col)
        closing_label = f'Closing Balance Until  {last_day_str}'
        ws_idx.merge_range(7, 5, 7, 6, closing_label, bold_border_fmt)
        ws_idx.write(7, 7, closing_total, bold_money_fmt)

        # Signature block
        sig_row = 10
        for col, name in [(3, 'Prepared By'), (6, 'Checked & Verified By'), (8, 'Approved By')]:
            ws_idx.write(sig_row, col, name, sign_fmt)
        for col, name in [(3, 'Kanniammal N'), (6, 'Sneha V John'), (8, 'Basil Issac')]:
            ws_idx.write(sig_row + 2, col, name, sign_fmt)
        for col, role in [(3, 'Inventory Lead'), (6, 'Project Manager'), (8, 'Vice President R&D')]:
            ws_idx.write(sig_row + 3, col, role, sign_fmt)

        ws_idx.set_row(3, 20)
        for col, w in enumerate([6, 6, 6, 15, 10, 8, 40, 18]):
            ws_idx.set_column(col, col, w)

        # ══════════════════════════════════════════════════════════════════════
        # SHEET 2 – PROJECT WISE
        # ══════════════════════════════════════════════════════════════════════
        ws_pw = workbook.add_worksheet('Project wise ')

        pw_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
        })
        pw_header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAD3', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        pw_cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        pw_center_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        pw_money_fmt = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter', 'align': 'right',
        })
        pw_total_label_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'right', 'valign': 'vcenter',
            'bg_color': '#E0E0E0',
        })
        pw_total_money_fmt = workbook.add_format({
            'bold': True, 'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
            'align': 'right', 'bg_color': '#E0E0E0',
        })

        # Row 0: sheet title
        ws_pw.merge_range(0, 2, 0, 8,
                          f'Bangalore Inventory Stock Details {year}-{str(year + 1)[-2:]}',
                          pw_title_fmt)

        # Headers – row 1
        # Col:  2=SL.No  3=Projects  4=Total Stock Until [prev month]
        #        5=Inward [month]  6=Total Stock [month]
        #        7=Outward [month]  8=Closing Stock [month]
        col_headers = [
            'SL.No',
            'Projects',
            f'Total Stock Until\n{prev_month_label}[INR]',
            f'{month_name} Stock-{year} [INR]',
            f'Total Stock {month_name}\n{year}[INR]',
            f'{month_name}  Month\nOutward {year}[INR]',
            f'Closing Stock In\n{last_day_str}  {year}[INR]',
        ]
        for i, h in enumerate(col_headers):
            ws_pw.write(1, 2 + i, h, pw_header_fmt)

        # Collect all projects (union of opening, inward, outward)
        all_projects = sorted(
            set(list(proj_opening.keys()) + list(proj_inward.keys()) + list(proj_outward.keys()))
        )

        row_pw = 2
        sl = 1
        sum_opening = sum_inward = sum_outward = sum_total = sum_closing = 0.0

        for proj in all_projects:
            ob = proj_opening.get(proj, 0.0)
            iw = proj_inward.get(proj, 0.0)
            ow = proj_outward.get(proj, 0.0)
            total_stock = ob + iw       # col 6 = col4 + col5
            closing = total_stock - ow  # col 8 = col6 - col7

            sum_opening += ob
            sum_inward  += iw
            sum_outward += ow
            sum_total   += total_stock
            sum_closing += closing

            ws_pw.write(row_pw, 2, sl,          pw_center_fmt)
            ws_pw.write(row_pw, 3, proj,         pw_cell_fmt)
            ws_pw.write(row_pw, 4, ob,           pw_money_fmt)
            ws_pw.write(row_pw, 5, iw,           pw_money_fmt)
            ws_pw.write(row_pw, 6, total_stock,  pw_money_fmt)
            ws_pw.write(row_pw, 7, ow,           pw_money_fmt)
            ws_pw.write(row_pw, 8, closing,      pw_money_fmt)

            row_pw += 1
            sl += 1

        # Total row
        ws_pw.merge_range(row_pw, 2, row_pw, 3, 'Total Amount:', pw_total_label_fmt)
        ws_pw.write(row_pw, 4, sum_opening,  pw_total_money_fmt)
        ws_pw.write(row_pw, 5, sum_inward,   pw_total_money_fmt)
        ws_pw.write(row_pw, 6, sum_total,    pw_total_money_fmt)
        ws_pw.write(row_pw, 7, sum_outward,  pw_total_money_fmt)
        ws_pw.write(row_pw, 8, sum_closing,  pw_total_money_fmt)
        row_pw += 2

        # Signature block
        for col, label in [(2, 'Prepared By'), (5, 'Checked & Verified By'), (8, 'Approved By')]:
            ws_pw.write(row_pw, col, label, sign_fmt)
        for col, name in [(2, 'Kanniammal N'), (5, 'Sneha V John'), (8, 'Basil Issac')]:
            ws_pw.write(row_pw + 2, col, name, sign_fmt)
        for col, role in [(2, 'Inventory Lead'), (5, 'Project Manager'), (8, 'Vice President R&D')]:
            ws_pw.write(row_pw + 3, col, role, sign_fmt)

        ws_pw.set_row(1, 45)
        for col, w in enumerate([5, 5, 8, 35, 22, 20, 22, 22, 22]):
            ws_pw.set_column(col, col, w)

        # ══════════════════════════════════════════════════════════════════════
        # SHEET 3 – INWARDS  (existing logic, unchanged structure)
        # ══════════════════════════════════════════════════════════════════════
        ws_in = workbook.add_worksheet('Inwards')

        # Reuse formats from workbook (already defined above)
        ws_in.merge_range(0, 0, 0, 17, f'INWARD REPORT - {period_title}', title_fmt)

        headers_in = [
            'SI No', 'GRN NO.', 'Invoice Date', 'Received Date', 'Project Code',
            'PR Number', 'PO Number', 'Invoice Number', 'Description', 'Part Number',
            'PO Qty', 'Received Qty', 'Unit', 'Rate', 'GST', 'Total Amount',
            'Grand Total', 'Supplier',
        ]
        for col, h in enumerate(headers_in):
            ws_in.write(1, col, h, header_fmt)

        row = 2
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

                ws_in.write(row, 0,  si_no,                                            center_fmt)
                ws_in.write(row, 1,  picking.e_grn_id.name or '',                     cell_fmt)
                ws_in.write(row, 2,  picking.e_invoice_date or '',                    cell_fmt)
                ws_in.write(row, 3,  picking.e_invoice_received_date or '',           cell_fmt)
                ws_in.write(row, 4,  project_code,                                     cell_fmt)
                ws_in.write(row, 5,  picking.e_pr_id.pr_number if picking.e_pr_id else '', cell_fmt)
                ws_in.write(row, 6,  po_number,                                        cell_fmt)
                ws_in.write(row, 7,  invoice_number,                                   cell_fmt)
                ws_in.write(row, 8,  description,                                      cell_fmt)
                ws_in.write(row, 9,  part_no,                                          cell_fmt)
                ws_in.write(row, 10, move.product_uom_qty,                             money_fmt)
                ws_in.write(row, 11, move.quantity,                                    money_fmt)
                ws_in.write(row, 12, move.product_uom.name or '',                     center_fmt)
                ws_in.write(row, 13, move.price_unit or 0.0,                          money_fmt)
                ws_in.write(row, 14, gst_amount,                                       money_fmt)
                ws_in.write(row, 15, line_total,                                       money_fmt)
                ws_in.write(row, 16, '',                                               cell_fmt)
                ws_in.write(row, 17, picking.partner_id.name or '',                   cell_fmt)

                row += 1
                si_no += 1

        if row > 2:
            row += 1
            ws_in.write(row, 14, 'GRAND TOTAL', workbook.add_format({
                'bold': True, 'border': 1, 'align': 'right', 'bg_color': '#E0E0E0',
            }))
            ws_in.write(row, 16, grand_total_in, workbook.add_format({
                'bold': True, 'border': 1, 'align': 'center',
                'num_format': '#,##0.00', 'bg_color': '#E0E0E0',
            }))

        widths_in = [6, 15, 12, 12, 20, 12, 14, 14, 35, 16, 10, 10, 8, 12, 12, 14, 14, 28]
        for col, w in enumerate(widths_in):
            ws_in.set_column(col, col, w)

        # ══════════════════════════════════════════════════════════════════════
        # SHEET 4 – OUTWARDS  (existing logic, unchanged structure)
        # ══════════════════════════════════════════════════════════════════════
        ws_out = workbook.add_worksheet('Outwards')

        ws_out.merge_range(0, 0, 0, 12, f'OUTWARD REPORT - {period_title}', title_fmt)

        out_header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#E6F3FF', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        out_cell_fmt    = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        out_center_fmt  = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        out_money_fmt   = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        merge_fmt       = workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'left'})
        merge_center_fmt= workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'center'})
        grand_label_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'right', 'valign': 'vcenter', 'bg_color': '#E0E0E0'})
        grand_value_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'num_format': '#,##0.00'})

        headers_out = [
            'SL No', 'MWR NO.', 'Date', 'Project Code',
            'Employee Name', 'Description', 'Part No',
            'Qty', 'Unit', 'Rate', 'GST', 'Total Amount', 'Justification/Remarks',
        ]
        for col, h in enumerate(headers_out):
            ws_out.write(1, col, h, out_header_fmt)

        row_out = 2
        sl_no = 1
        grand_total_out = 0.0

        for picking in outwards_pickings:
            if not picking.move_ids:
                continue

            project_code = f"{picking.e_project_id.name or ''}-{picking.e_task_id.name or ''}".strip('-')
            employee_name = picking.e_requested_by.name if picking.e_requested_by else ''
            mwr_no = picking.name or ''

            date_val = picking.e_date or picking.date_done or picking.scheduled_date or fields.Date.today()
            date_str = date_val.strftime('%Y-%m-%d') if hasattr(date_val, 'strftime') else str(date_val)

            valid_moves = picking.move_ids.filtered(lambda m: m.state != 'cancel')

            for move in valid_moves:
                description = html2plaintext(move.e_description_out or move.product_id.name or '').strip()
                part_no     = html2plaintext(move.e_part_no or '').strip()
                unit_name   = move.product_uom.name or ''
                remarks     = html2plaintext(move.e_remarks or '').strip()

                valuation_layers = move.eram_valuation_ids.sorted(key=lambda l: (l.create_date or l.id or 0))
                group_start_row  = row_out

                if not valuation_layers:
                    qty = abs(move.product_uom_qty or 0.0)
                    ws_out.write(row_out, 0,  sl_no,        out_center_fmt)
                    ws_out.write(row_out, 1,  mwr_no,       out_cell_fmt)
                    ws_out.write(row_out, 2,  date_str,     out_cell_fmt)
                    ws_out.write(row_out, 3,  project_code, out_cell_fmt)
                    ws_out.write(row_out, 4,  employee_name,out_cell_fmt)
                    ws_out.write(row_out, 5,  description,  out_cell_fmt)
                    ws_out.write(row_out, 6,  part_no,      out_cell_fmt)
                    ws_out.write(row_out, 7,  qty,          out_center_fmt)
                    ws_out.write(row_out, 8,  unit_name,    out_center_fmt)
                    ws_out.write(row_out, 9,  0.0,          out_money_fmt)
                    ws_out.write(row_out, 10, 0.0,          out_money_fmt)
                    ws_out.write(row_out, 11, 0.0,          out_money_fmt)
                    ws_out.write(row_out, 12, remarks,      out_cell_fmt)
                    row_out += 1
                    sl_no   += 1
                    continue

                for i, layer in enumerate(valuation_layers):
                    layer_qty   = abs(layer.quantity  or 0.0)
                    total_taxed = abs(layer.total_taxed or 0.0)
                    value       = abs(layer.value     or 0.0)
                    gst_amount  = total_taxed - value
                    unit_rate   = abs(layer.unit_cost  or 0.0)

                    grand_total_out += total_taxed

                    if i == 0:
                        ws_out.write(row_out, 0, sl_no,        out_center_fmt)
                        ws_out.write(row_out, 1, mwr_no,       out_cell_fmt)
                        ws_out.write(row_out, 2, date_str,     out_cell_fmt)
                        ws_out.write(row_out, 3, project_code, out_cell_fmt)
                        ws_out.write(row_out, 4, employee_name,out_cell_fmt)
                        ws_out.write(row_out, 5, description,  out_cell_fmt)
                        ws_out.write(row_out, 6, part_no,      out_cell_fmt)
                    else:
                        for c in range(7):
                            ws_out.write(row_out, c, '', out_cell_fmt if c != 0 else out_center_fmt)

                    ws_out.write(row_out, 7,  layer_qty,   out_center_fmt)
                    ws_out.write(row_out, 8,  unit_name,   out_center_fmt)
                    ws_out.write(row_out, 9,  unit_rate,   out_money_fmt)
                    ws_out.write(row_out, 10, gst_amount,  out_money_fmt)
                    ws_out.write(row_out, 11, total_taxed, out_money_fmt)
                    ws_out.write(row_out, 12, remarks if i == 0 else '', out_cell_fmt)

                    row_out += 1

                last_row = row_out - 1
                if group_start_row < last_row:
                    ws_out.merge_range(group_start_row, 0, last_row, 0, sl_no,         merge_center_fmt)
                    ws_out.merge_range(group_start_row, 1, last_row, 1, mwr_no,        merge_fmt)
                    ws_out.merge_range(group_start_row, 2, last_row, 2, date_str,      merge_fmt)
                    ws_out.merge_range(group_start_row, 3, last_row, 3, project_code,  merge_fmt)
                    ws_out.merge_range(group_start_row, 4, last_row, 4, employee_name, merge_fmt)
                    ws_out.merge_range(group_start_row, 5, last_row, 5, description,   merge_fmt)
                    ws_out.merge_range(group_start_row, 6, last_row, 6, part_no,       merge_fmt)
                    if remarks:
                        ws_out.merge_range(group_start_row, 12, last_row, 12, remarks, merge_fmt)

                sl_no += 1

        if row_out > 2:
            row_out += 1
            ws_out.write(row_out, 10, 'GRAND TOTAL',   grand_label_fmt)
            ws_out.write(row_out, 11, grand_total_out, grand_value_fmt)

        widths_out = [8, 22, 14, 20, 22, 40, 18, 12, 10, 14, 14, 16, 45]
        for col, w in enumerate(widths_out):
            ws_out.set_column(col, col, w)

        # ── finalise ──────────────────────────────────────────────────────────
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()