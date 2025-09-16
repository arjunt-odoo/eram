# -*- coding: utf-8 -*-
from odoo import fields, models
from odoo.tools import json_default
from datetime import date
import json
import io
import os
from PIL import Image

try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter


class EramSaleOrderReport(models.TransientModel):
    _name = 'eram.sale.order.report'
    _description = 'Eram Sale Order Report'

    from_date = fields.Date(string="From Date")
    to_date = fields.Date(string="To Date")
    type = fields.Selection([('xlsx', "XLSX")], string="Report Type",
                            default='xlsx', required=True)

    def action_print_report(self):
        if self.type == 'xlsx':
            domain = [('state', '=', 'sale')]
            if self.from_date and self.to_date:
                domain.append(('date_order', '>=', self.from_date))
                domain.append(('date_order', '<=', self.to_date))
            elif self.from_date and not self.to_date:
                domain.append(('date_order', '>=', self.from_date))
            elif not self.from_date and self.to_date:
                domain.append(('date_order', '<=', self.to_date))
            records = self.env['sale.order'].search_read(domain, ['id'])
            record_ids = [item.get('id', False) for item in records if item]
            data = {
                'orders': record_ids
            }
            return {
                'type': 'ir.actions.report',
                'data': {'model': 'eram.sale.order.report',
                         'options': json.dumps(data, default=json_default),
                         'output_format': 'xlsx',
                         'report_name': 'Sales Excel Report',
                         },
                'report_type': 'xlsx',
            }

    def get_xlsx_report(self, data, response):
        order_id_list = data.get('orders', [])
        order_id_list.sort()
        order_ids = self.env['sale.order'].browse(order_id_list)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        light_green = '#C6EFCE'
        red = '#FF0000'
        heading_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 36,
            'bg_color': light_green,
            'text_wrap': True
        })
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True
        })
        date_format = workbook.add_format(
            {'num_format': 'dd-mm-yyyy', 'align': 'center', 'text_wrap': True})
        currency_format = workbook.add_format(
            {'num_format': '#,##0.00', 'align': 'center', 'text_wrap': True})
        center_format = workbook.add_format({'align': 'center', 'text_wrap': True})
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        merge_date_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'dd-mm-yyyy',
            'text_wrap': True
        })
        red_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_color': red
        })
        row_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        col_widths = [8] + [15] * 22

        module_path = os.path.dirname(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        logo_path = os.path.join(
            module_path,
            'eram_report_templates',
            'static',
            'src',
            'img',
            'logo.png'
        )

        footer_path = os.path.join(
            module_path,
            'eram_report_templates',
            'static',
            'src',
            'img',
            'footer.png'
        )

        if os.path.exists(logo_path):
            with Image.open(logo_path) as img:
                logo_width, logo_height = img.size

            scale_factor = 0.7
            sheet.insert_image('A1', logo_path, {
                'x_offset': 10,
                'y_offset': 21,
                'x_scale': scale_factor,
                'y_scale': scale_factor,
                'object_position': 3
            })

        report_date = fields.Date.context_today(self)
        sheet.merge_range('A2:W2', f'SALES ORDER REPORT-{report_date}', heading_format)

        sheet.set_row(1, 45)

        headers = [
            'SI No', 'Full Name', 'Account Name', 'Product Name', 'Description',
            'Quantity', 'Quote No', 'Quote Date', 'Purchase Order No',
            'Invoice No', 'Invoice Date', 'Invoice Value', 'Payment Terms',
            'Payment Status', 'Advance Payment Received/Not Received',
            'Advance Payment Amount', 'Advance Payment Received Date',
            'Balance Payment', 'Balance Payment Received Date', 'Payment Due',
            'Payment Due Date', 'Shipment Status', 'Remarks'
        ]

        for col, header in enumerate(headers):
            sheet.write(3, col, header, header_format)
            col_widths[col] = max(col_widths[col], len(header) + 2)

        row = 4
        grand_total_qty = 0
        grand_total_value = 0
        si_no = 1

        def write_center(sheet, row, col, value, format=None, is_overdue=False):
            if format is None:
                format = red_format if is_overdue else row_format

            sheet.write(row, col, value, format)
            if value is not None:
                col_widths[col] = max(col_widths[col], len(str(value)) + 2)

        # Track merged ranges to avoid overlaps
        merged_ranges = set()

        def safe_merge(sheet, first_row, first_col, last_row, last_col, data, cell_format=None):
            """Safely merge cells only if the range hasn't been merged before"""
            merge_key = f"{first_row}:{first_col}:{last_row}:{last_col}"
            if merge_key in merged_ranges:
                return

            if first_row == last_row and first_col == last_col:
                # Single cell, no need to merge
                if cell_format:
                    sheet.write(first_row, first_col, data, cell_format)
                else:
                    sheet.write(first_row, first_col, data)
            else:
                # Multi-cell merge
                if cell_format:
                    sheet.merge_range(first_row, first_col, last_row, last_col, data, cell_format)
                else:
                    sheet.merge_range(first_row, first_col, last_row, last_col, data)

            merged_ranges.add(merge_key)

        for order in order_ids:
            order_start_row = row
            partner = order.partner_invoice_id or order.partner_id
            full_name = partner.name
            account_name = order.partner_id.name
            order_line_quantities = {}
            current_si_no = si_no

            # First pass: collect all invoice data for each line and count total rows for this order
            line_invoice_data = {}
            total_order_rows = 0

            for line in order.order_line:
                invoice_lines = self.env['account.move.line'].search([
                    ('sale_line_ids', 'in', line.ids),
                    ('parent_state', '!=', 'cancel')
                ])
                order_line_quantities[line.id] = line.product_uom_qty

                invoices = invoice_lines.mapped('move_id')
                line_invoice_data[line.id] = []
                invoice_count = 0

                if not invoices:
                    line_invoice_data[line.id].append({
                        'invoice_no': '-',
                        'invoice_date': '-',
                        'invoice_value': line.price_total,
                        'payment_status': 'NOT RECEIVED',
                        'advance_payment': 'Not Received',
                        'advance_amount': 0.0,
                        'advance_date': '-',
                        'balance_payment': line.price_total,
                        'balance_date': '-',
                        'payment_due': '-',
                        'payment_due_date': '-',
                        'buyer_order_no': '-',
                        'is_overdue': False,
                        'payment_terms': '-'
                    })
                    invoice_count = 1
                else:
                    for invoice in invoices:
                        invoice_amount = invoice.amount_total
                        amount_residual = invoice.amount_residual

                        advance_amount = sum(invoice.matched_payment_ids.filtered(
                            lambda p: p.state in ('in_process', 'paid')).mapped('amount'))

                        days_overdue = 0
                        payment_due_display = '-'
                        is_overdue = False

                        if invoice.invoice_date_due:
                            today = date.today()
                            due_date = fields.Date.from_string(invoice.invoice_date_due)
                            if today > due_date:
                                days_overdue = (today - due_date).days
                                payment_due_display = f"{days_overdue} days overdue"
                                is_overdue = True
                            else:
                                payment_due_display = '-'
                        else:
                            payment_due_display = '-'

                        if invoice.payment_state == 'paid':
                            payment_status = 'RECEIVED'
                            advance_payment = 'Received'
                            advance_date = invoice.invoice_date
                            balance_payment = 0.0
                            balance_date = invoice.invoice_date
                            payment_due_display = '-'
                        elif invoice.payment_state == 'partial':
                            payment_status = 'PARTIALLY RECEIVED'
                            advance_payment = 'Received'
                            advance_date = invoice.invoice_date
                            balance_payment = amount_residual
                            balance_date = invoice.invoice_date_due
                            if days_overdue > 0:
                                payment_due_display = f"{days_overdue} days overdue"
                                is_overdue = True
                            else:
                                payment_due_display = '-'
                        else:
                            payment_status = 'NOT RECEIVED'
                            advance_payment = 'Not Received'
                            advance_date = '-'
                            balance_payment = amount_residual
                            balance_date = invoice.invoice_date_due
                            if days_overdue > 0:
                                payment_due_display = f"{days_overdue} days overdue"
                                is_overdue = True
                            else:
                                payment_due_display = '-'

                        buyer_order_no = invoice.e_buyer_order_no or '-'

                        line_invoice_data[line.id].append({
                            'invoice_no': invoice.name or '-',
                            'invoice_date': invoice.invoice_date or '-',
                            'invoice_value': invoice_amount,
                            'payment_status': payment_status,
                            'advance_payment': advance_payment,
                            'advance_amount': advance_amount,
                            'advance_date': advance_date,
                            'balance_payment': balance_payment,
                            'balance_date': balance_date,
                            'payment_due': payment_due_display,
                            'payment_due_date': invoice.invoice_date_due or '-',
                            'buyer_order_no': buyer_order_no,
                            'is_overdue': is_overdue,
                            'payment_terms': invoice.invoice_payment_term_id.name
                        })
                        invoice_count += 1

                total_order_rows += invoice_count if invoice_count > 0 else 1

            # Second pass: write data to spreadsheet
            for line in order.order_line:
                invoice_data = line_invoice_data.get(line.id, [])

                picking_ids = line.move_ids.picking_id.filtered(lambda p: p.state != 'cancel')
                if picking_ids:
                    if all(picking.state == 'done' for picking in picking_ids):
                        shipment_status = 'Shipped'
                    elif all(picking.state != 'done' for picking in picking_ids):
                        shipment_status = 'Not Shipped'
                    else:
                        shipment_status = 'Partially Shipped'
                else:
                    shipment_status = 'Nothing to Ship'

                line_start_row = row

                for i, inv in enumerate(invoice_data):
                    # For the first row of each order, write the order-level data
                    if row == order_start_row:
                        write_center(sheet, row, 0, current_si_no, center_format)
                        write_center(sheet, row, 1, full_name)
                        write_center(sheet, row, 2, account_name)
                        write_center(sheet, row, 6, order.name)
                        write_center(sheet, row, 7, order.date_order, date_format)
                        write_center(sheet, row, 12, order.payment_term_id.name or '-')
                        write_center(sheet, row, 22, order.note or '-')
                    else:
                        # For subsequent rows of the same order, leave these cells empty
                        write_center(sheet, row, 0, '', center_format)
                        write_center(sheet, row, 1, '')
                        write_center(sheet, row, 2, '')
                        write_center(sheet, row, 6, '')
                        write_center(sheet, row, 7, '')
                        write_center(sheet, row, 12, '')
                        write_center(sheet, row, 22, '')

                    # Write line-specific data for all rows
                    write_center(sheet, row, 3, line.product_id.name)
                    write_center(sheet, row, 4, line.e_description or "-")
                    line_quantity = order_line_quantities.get(line.id, line.product_uom_qty)
                    write_center(sheet, row, 5, line_quantity, center_format)
                    write_center(sheet, row, 8, inv['buyer_order_no'])
                    write_center(sheet, row, 9, inv['invoice_no'])

                    if inv['invoice_date'] != '-':
                        write_center(sheet, row, 10, inv['invoice_date'], date_format)
                    else:
                        write_center(sheet, row, 10, inv['invoice_date'])

                    write_center(sheet, row, 11, inv['invoice_value'], currency_format)
                    write_center(sheet, row, 13, inv['payment_status'])
                    write_center(sheet, row, 14, inv['advance_payment'])
                    write_center(sheet, row, 15, inv['advance_amount'], currency_format)

                    if isinstance(inv['advance_date'], date) or (
                            isinstance(inv['advance_date'], str) and inv['advance_date'] != '-'):
                        write_center(sheet, row, 16, inv['advance_date'], date_format)
                    else:
                        write_center(sheet, row, 16, inv['advance_date'])

                    write_center(sheet, row, 17, inv['balance_payment'], currency_format)

                    if isinstance(inv['balance_date'], date) or (
                            isinstance(inv['balance_date'], str) and inv['balance_date'] != '-'):
                        write_center(sheet, row, 18, inv['balance_date'], date_format)
                    else:
                        write_center(sheet, row, 18, inv['balance_date'])

                    write_center(sheet, row, 19, inv['payment_due'], None, inv['is_overdue'])

                    if inv['payment_due_date'] != '-':
                        write_center(sheet, row, 20, inv['payment_due_date'], date_format)
                    else:
                        write_center(sheet, row, 20, inv['payment_due_date'])

                    write_center(sheet, row, 21, shipment_status)

                    grand_total_value += inv['invoice_value']
                    row += 1

            order_end_row = row - 1

            # Merge order-level data across all rows of this order
            if order_end_row > order_start_row:
                safe_merge(sheet, order_start_row, 0, order_end_row, 0, current_si_no, merge_format)
                safe_merge(sheet, order_start_row, 1, order_end_row, 1, full_name, merge_format)
                safe_merge(sheet, order_start_row, 2, order_end_row, 2, account_name, merge_format)
                safe_merge(sheet, order_start_row, 6, order_end_row, 6, order.name, merge_format)

                if order.date_order:
                    safe_merge(sheet, order_start_row, 7, order_end_row, 7, order.date_order, merge_date_format)
                else:
                    safe_merge(sheet, order_start_row, 7, order_end_row, 7, '', merge_format)

                safe_merge(sheet, order_start_row, 12, order_end_row, 12, order.payment_term_id.name or '-',
                           merge_format)
                safe_merge(sheet, order_start_row, 22, order_end_row, 22, order.note or '-', merge_format)

            grand_total_qty += sum(order_line_quantities.values())
            si_no += 1

        total_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True
        })

        write_center(sheet, row, 0, 'Grand Total', total_format)
        sheet.write(row, 5, grand_total_qty, total_format)
        sheet.write(row, 11, grand_total_value, total_format)

        for col, width in enumerate(col_widths):
            sheet.set_column(col, col, min(width, 50))

        if os.path.exists(footer_path):
            total_width_pixels = sum([min(w, 50) * 8.1 for w in col_widths])

            with Image.open(footer_path) as img:
                footer_width, footer_height = img.size

            scale_factor = total_width_pixels / 1000

            footer_row = row + 3
            sheet.insert_image(
                f'A{footer_row}',
                footer_path,
                {
                    'x_offset': 0,
                    'y_offset': 0,
                    'x_scale': scale_factor / 4,
                    'y_scale': scale_factor / 4,
                    'object_position': 3
                }
            )
            sheet.set_row(footer_row - 1, footer_height * (scale_factor / 2) / 1.33)

        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()