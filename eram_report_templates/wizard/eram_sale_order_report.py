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
            domain = [('state', '!=', 'cancel')]
            if self.from_date and self.to_date:
                domain.append(('date_order', '>=', self.from_date))
                domain.append(('date_order', '<=', self.to_date))
            elif self.from_date and not self.to_date:
                domain.append(('date_order', '>=', self.from_date))
            elif not self.from_date and self.to_date:
                domain.append(('date_order', '<=', self.to_date))
            records = self.env['sale.order'].search_read(domain, ['id', 'state'])
            record_ids = [item.get('id', False) for item in records if item]

            sale_orders = []
            other_orders = []
            for item in records:
                if item.get('state') == 'sale':
                    sale_orders.append(item.get('id'))
                else:
                    other_orders.append(item.get('id'))

            sorted_record_ids = sale_orders + other_orders
            data = {
                'orders': sorted_record_ids
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
        order_ids = self.env['sale.order'].browse(order_id_list)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        light_green = '#C6EFCE'
        light_blue = '#D9E1F2'
        light_yellow = '#FFEB9C'
        red = '#FF0000'

        heading_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 36,
            'bg_color': light_green,
            'text_wrap': True,
            'border': 1
        })
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True
        })
        date_format = workbook.add_format({
            'num_format': 'dd-mm-yyyy',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        base_currency_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        merge_date_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'dd-mm-yyyy',
            'text_wrap': True,
            'border': 1
        })
        red_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_color': red,
            'border': 1
        })
        row_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        base_total_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True,
            'num_format': '#,##0.00'
        })
        total_qty_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True
        })

        even_row_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': light_blue,
            'border': 1
        })
        odd_row_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })

        even_currency_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': light_blue,
            'border': 1
        })
        odd_currency_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })

        even_date_format = workbook.add_format({
            'num_format': 'dd-mm-yyyy',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': light_blue,
            'border': 1
        })
        odd_date_format = workbook.add_format({
            'num_format': 'dd-mm-yyyy',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })

        even_red_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_color': red,
            'bg_color': light_blue,
            'border': 1
        })
        odd_red_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_color': red,
            'border': 1
        })

        # Increase column widths for better number display
        col_widths = [8, 20, 20, 20, 25, 12, 15, 15, 20, 15, 15, 18, 20, 20, 25, 20, 20, 18, 20, 18, 18, 18, 20]

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
        formatted_date = report_date.strftime('%d-%m-%Y') if report_date else ''
        sheet.merge_range('A2:W2', f'SALES ORDER REPORT-{formatted_date}', heading_format)

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
        grand_total_advance = 0
        grand_total_balance = 0
        si_no = 1

        def write_center(sheet, row, col, value, format=None, is_overdue=False):
            is_even_order = ((row - 4) // max_rows_per_order) % 2 == 0

            if format is None:
                if is_even_order:
                    format = even_row_format
                    if is_overdue:
                        format = even_red_format
                else:
                    format = odd_row_format
                    if is_overdue:
                        format = odd_red_format

            sheet.write(row, col, value, format)

        def get_currency_format(workbook, currency, is_total=False, is_even=False):
            symbol = currency.symbol
            position = currency.position

            if position == 'after':
                num_format = f'#,##0.00 "{symbol}"'
            else:
                num_format = f'"{symbol}" #,##0.00'

            if is_even:
                bg_color = light_blue
            else:
                bg_color = None

            format_props = {
                'num_format': num_format,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'border': 1
            }

            if bg_color:
                format_props['bg_color'] = bg_color

            if is_total:
                format_props.update({
                    'bold': True,
                    'border': 1,
                    'bg_color': light_green,
                    'font_size': 14
                })

            return workbook.add_format(format_props)

        merged_ranges = set()

        def safe_merge(sheet, first_row, first_col, last_row, last_col, data, cell_format=None):
            merge_key = f"{first_row}:{first_col}:{last_row}:{last_col}"
            if merge_key in merged_ranges:
                return

            is_even_order = ((first_row - 4) // max_rows_per_order) % 2 == 0

            if cell_format is None:
                if is_even_order:
                    cell_format = even_row_format
                else:
                    cell_format = odd_row_format

            if first_row == last_row and first_col == last_col:
                sheet.write(first_row, first_col, data, cell_format)
            else:
                sheet.merge_range(first_row, first_col, last_row, last_col, data, cell_format)

            merged_ranges.add(merge_key)

        for order_idx, order in enumerate(order_ids):
            order_start_row = row
            partner = order.partner_invoice_id or order.partner_id
            full_name = partner.name
            account_name = order.partner_id.name
            current_si_no = si_no

            is_even_order = (order_idx % 2 == 0)

            order_invoices = self.env['account.move'].search([
                ('invoice_origin', '=', order.name),
                ('state', '!=', 'cancel')
            ])

            order_lines = order.order_line

            # Get shipment status from sale order
            picking_ids = order.picking_ids.filtered(lambda p: p.state != 'cancel')
            if picking_ids:
                if all(picking.state == 'done' for picking in picking_ids):
                    shipment_status = 'Shipped'
                elif all(picking.state != 'done' for picking in picking_ids):
                    shipment_status = 'Not Shipped'
                else:
                    shipment_status = 'Partially Shipped'
            else:
                shipment_status = 'Nothing to Ship'

            invoice_count = len(order_invoices)
            product_count = len(order_lines)
            max_rows_per_order = max(invoice_count, product_count, 1)

            product_distribution = []
            invoice_distribution = []

            if product_count > 0 and invoice_count > 0:
                if product_count >= invoice_count:
                    base_rows_per_invoice = product_count // invoice_count
                    extra_rows = product_count % invoice_count

                    for i in range(invoice_count):
                        rows_for_this_invoice = base_rows_per_invoice
                        if i < extra_rows:
                            rows_for_this_invoice += 1
                        invoice_distribution.append(rows_for_this_invoice)

                    product_distribution = [1] * product_count

                else:
                    base_rows_per_product = invoice_count // product_count
                    extra_rows = invoice_count % product_count

                    for i in range(product_count):
                        rows_for_this_product = base_rows_per_product
                        if i < extra_rows:
                            rows_for_this_product += 1
                        product_distribution.append(rows_for_this_product)

                    invoice_distribution = [1] * invoice_count
            else:
                product_distribution = [1] * product_count if product_count > 0 else []
                invoice_distribution = [1] * invoice_count if invoice_count > 0 else []

            invoice_data = []
            for invoice in order_invoices:
                invoice_amount = invoice.amount_total
                amount_residual = invoice.amount_residual

                advance_amount = sum(invoice.matched_payment_ids.filtered(
                    lambda p: p.state in ('in_process', 'paid')).mapped('amount'))

                days_overdue = 0
                payment_due_display = 'N/A'
                is_overdue = False

                if invoice.invoice_date_due:
                    today = date.today()
                    due_date = fields.Date.from_string(invoice.invoice_date_due)
                    if today > due_date:
                        days_overdue = (today - due_date).days
                        payment_due_display = f"{days_overdue} days overdue"
                        is_overdue = True
                    else:
                        payment_due_display = 'N/A'
                else:
                    payment_due_display = 'N/A'

                if invoice.payment_state == 'paid':
                    payment_status = 'RECEIVED'
                    advance_payment = 'Received'
                    advance_date = invoice.invoice_date
                    balance_payment = 0.0
                    balance_date = invoice.invoice_date
                    payment_due_display = 'N/A'
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
                        payment_due_display = 'N/A'
                else:
                    payment_status = 'NOT RECEIVED'
                    advance_payment = 'Not Received'
                    advance_date = 'N/A'
                    balance_payment = amount_residual
                    balance_date = invoice.invoice_date_due
                    if days_overdue > 0:
                        payment_due_display = f"{days_overdue} days overdue"
                        is_overdue = True
                    else:
                        payment_due_display = 'N/A'

                buyer_order_no = invoice.e_buyer_order_no or 'N/A'
                payment_terms = invoice.invoice_payment_term_id.name or 'N/A'

                invoice_data.append({
                    'invoice_no': invoice.name or 'N/A',
                    'invoice_date': invoice.invoice_date or 'N/A',
                    'invoice_value': invoice_amount,
                    'payment_status': payment_status,
                    'advance_payment': advance_payment,
                    'advance_amount': advance_amount,
                    'advance_date': advance_date,
                    'balance_payment': balance_payment,
                    'balance_date': balance_date,
                    'payment_due': payment_due_display,
                    'payment_due_date': invoice.invoice_date_due or 'N/A',
                    'buyer_order_no': buyer_order_no,
                    'is_overdue': is_overdue,
                    'payment_terms': payment_terms
                })

            if is_even_order:
                row_fmt = even_row_format
                currency_fmt = even_currency_format
                date_fmt = even_date_format
                red_fmt = even_red_format
            else:
                row_fmt = odd_row_format
                currency_fmt = odd_currency_format
                date_fmt = odd_date_format
                red_fmt = odd_red_format

            order_currency = order.currency_id
            order_currency_format = get_currency_format(workbook, order_currency, is_even=is_even_order)

            for i in range(max_rows_per_order):
                if i == 0:
                    write_center(sheet, row + i, 0, current_si_no, row_fmt)
                    write_center(sheet, row + i, 1, full_name, row_fmt)
                    write_center(sheet, row + i, 2, account_name, row_fmt)
                    write_center(sheet, row + i, 6, order.name, row_fmt)
                    write_center(sheet, row + i, 7, order.date_order, date_fmt)
                    write_center(sheet, row + i, 21, shipment_status, row_fmt)
                    write_center(sheet, row + i, 22, 'N/A', row_fmt)
                else:
                    write_center(sheet, row + i, 0, '', row_fmt)
                    write_center(sheet, row + i, 1, '', row_fmt)
                    write_center(sheet, row + i, 2, '', row_fmt)
                    write_center(sheet, row + i, 6, '', row_fmt)
                    write_center(sheet, row + i, 7, '', row_fmt)
                    write_center(sheet, row + i, 21, '', row_fmt)
                    write_center(sheet, row + i, 22, 'N/A', row_fmt)

            current_row = 0
            for product_idx, line in enumerate(order_lines):
                rows_for_this_product = product_distribution[product_idx] if product_distribution else 1

                for i in range(rows_for_this_product):
                    if current_row + i < max_rows_per_order:
                        write_center(sheet, row + current_row + i, 3, line.product_id.name, row_fmt)
                        write_center(sheet, row + current_row + i, 4, line.e_description or "N/A", row_fmt)
                        write_center(sheet, row + current_row + i, 5, line.product_uom_qty, row_fmt)
                        grand_total_qty += line.product_uom_qty

                if rows_for_this_product > 1:
                    safe_merge(sheet, row + current_row, 3, row + current_row + rows_for_this_product - 1, 3,
                               line.product_id.name, row_fmt)
                    safe_merge(sheet, row + current_row, 4, row + current_row + rows_for_this_product - 1, 4,
                               line.e_description or "N/A", row_fmt)
                    safe_merge(sheet, row + current_row, 5, row + current_row + rows_for_this_product - 1, 5,
                               line.product_uom_qty, row_fmt)

                current_row += rows_for_this_product

            for i in range(current_row, max_rows_per_order):
                write_center(sheet, row + i, 3, '', row_fmt)
                write_center(sheet, row + i, 4, '', row_fmt)
                write_center(sheet, row + i, 5, '', row_fmt)

            current_row = 0
            if invoice_count > 0:
                for invoice_idx, inv in enumerate(invoice_data):
                    rows_for_this_invoice = invoice_distribution[invoice_idx] if invoice_distribution else 1

                    for i in range(rows_for_this_invoice):
                        if current_row + i < max_rows_per_order:
                            write_center(sheet, row + current_row + i, 8, inv['buyer_order_no'], row_fmt)
                            write_center(sheet, row + current_row + i, 9, inv['invoice_no'], row_fmt)

                            if inv['invoice_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 10, inv['invoice_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 10, inv['invoice_date'], row_fmt)

                            write_center(sheet, row + current_row + i, 11, inv['invoice_value'], order_currency_format)
                            write_center(sheet, row + current_row + i, 12, inv['payment_terms'], row_fmt)
                            write_center(sheet, row + current_row + i, 13, inv['payment_status'], row_fmt)
                            write_center(sheet, row + current_row + i, 14, inv['advance_payment'], row_fmt)
                            write_center(sheet, row + current_row + i, 15, inv['advance_amount'], order_currency_format)

                            if isinstance(inv['advance_date'], date) or (
                                    isinstance(inv['advance_date'], str) and inv['advance_date'] != 'N/A'):
                                write_center(sheet, row + current_row + i, 16, inv['advance_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 16, inv['advance_date'], row_fmt)

                            write_center(sheet, row + current_row + i, 17, inv['balance_payment'],
                                         order_currency_format)

                            if isinstance(inv['balance_date'], date) or (
                                    isinstance(inv['balance_date'], str) and inv['balance_date'] != 'N/A'):
                                write_center(sheet, row + current_row + i, 18, inv['balance_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 18, inv['balance_date'], row_fmt)

                            if inv['is_overdue']:
                                write_center(sheet, row + current_row + i, 19, inv['payment_due'], red_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 19, inv['payment_due'], row_fmt)

                            if inv['payment_due_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 20, inv['payment_due_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 20, inv['payment_due_date'], row_fmt)

                            grand_total_value += inv['invoice_value'] if i == 0 else 0
                            grand_total_advance += inv['advance_amount'] if i == 0 else 0
                            grand_total_balance += inv['balance_payment'] if i == 0 else 0

                    if rows_for_this_invoice > 1:
                        safe_merge(sheet, row + current_row, 8, row + current_row + rows_for_this_invoice - 1, 8,
                                   inv['buyer_order_no'], row_fmt)
                        safe_merge(sheet, row + current_row, 9, row + current_row + rows_for_this_invoice - 1, 9,
                                   inv['invoice_no'], row_fmt)
                        safe_merge(sheet, row + current_row, 10, row + current_row + rows_for_this_invoice - 1, 10,
                                   inv['invoice_date'], date_fmt if inv['invoice_date'] != 'N/A' else row_fmt)
                        safe_merge(sheet, row + current_row, 11, row + current_row + rows_for_this_invoice - 1, 11,
                                   inv['invoice_value'], order_currency_format)
                        safe_merge(sheet, row + current_row, 12, row + current_row + rows_for_this_invoice - 1, 12,
                                   inv['payment_terms'], row_fmt)
                        safe_merge(sheet, row + current_row, 13, row + current_row + rows_for_this_invoice - 1, 13,
                                   inv['payment_status'], row_fmt)
                        safe_merge(sheet, row + current_row, 14, row + current_row + rows_for_this_invoice - 1, 14,
                                   inv['advance_payment'], row_fmt)
                        safe_merge(sheet, row + current_row, 15, row + current_row + rows_for_this_invoice - 1, 15,
                                   inv['advance_amount'], order_currency_format)
                        safe_merge(sheet, row + current_row, 16, row + current_row + rows_for_this_invoice - 1, 16,
                                   inv['advance_date'], date_fmt if isinstance(inv['advance_date'], date) or (
                                    isinstance(inv['advance_date'], str) and inv['advance_date'] != 'N/A') else row_fmt)
                        safe_merge(sheet, row + current_row, 17, row + current_row + rows_for_this_invoice - 1, 17,
                                   inv['balance_payment'], order_currency_format)
                        safe_merge(sheet, row + current_row, 18, row + current_row + rows_for_this_invoice - 1, 18,
                                   inv['balance_date'], date_fmt if isinstance(inv['balance_date'], date) or (
                                    isinstance(inv['balance_date'], str) and inv['balance_date'] != 'N/A') else row_fmt)
                        safe_merge(sheet, row + current_row, 19, row + current_row + rows_for_this_invoice - 1, 19,
                                   inv['payment_due'], red_fmt if inv['is_overdue'] else row_fmt)
                        safe_merge(sheet, row + current_row, 20, row + current_row + rows_for_this_invoice - 1, 20,
                                   inv['payment_due_date'], date_fmt if inv['payment_due_date'] != 'N/A' else row_fmt)

                    current_row += rows_for_this_invoice
            else:
                # If no invoices, merge all invoice-related columns into a single blank cell
                safe_merge(sheet, row, 8, row + max_rows_per_order - 1, 20, '', row_fmt)

            for i in range(current_row, max_rows_per_order):
                for col in range(8, 21):
                    write_center(sheet, row + i, col, '', row_fmt)

            order_end_row = row + max_rows_per_order - 1
            if order_end_row > order_start_row:
                safe_merge(sheet, order_start_row, 0, order_end_row, 0, current_si_no, row_fmt)
                safe_merge(sheet, order_start_row, 1, order_end_row, 1, full_name, row_fmt)
                safe_merge(sheet, order_start_row, 2, order_end_row, 2, account_name, row_fmt)
                safe_merge(sheet, order_start_row, 6, order_end_row, 6, order.name, row_fmt)

                if order.date_order:
                    safe_merge(sheet, order_start_row, 7, order_end_row, 7, order.date_order, date_fmt)
                else:
                    safe_merge(sheet, order_start_row, 7, order_end_row, 7, '', row_fmt)

                safe_merge(sheet, order_start_row, 21, order_end_row, 21, shipment_status, row_fmt)
                safe_merge(sheet, order_start_row, 22, order_end_row, 22, 'N/A', row_fmt)

            row += max_rows_per_order
            si_no += 1

        company_currency = self.env.company.currency_id
        company_currency_format = get_currency_format(workbook, company_currency, is_total=True)
        company_total_qty_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': light_green,
            'text_wrap': True,
            'font_size': 14
        })

        write_center(sheet, row, 0, 'Grand Total', company_total_qty_format)
        for col in range(1, 5):
            write_center(sheet, row, col, '', company_total_qty_format)
        write_center(sheet, row, 5, grand_total_qty, company_total_qty_format)
        for col in range(6, 11):
            write_center(sheet, row, col, '', company_total_qty_format)
        write_center(sheet, row, 11, grand_total_value, company_currency_format)
        write_center(sheet, row, 12, '', company_total_qty_format)
        write_center(sheet, row, 13, '', company_total_qty_format)
        write_center(sheet, row, 14, '', company_total_qty_format)
        write_center(sheet, row, 15, grand_total_advance, company_currency_format)
        write_center(sheet, row, 16, '', company_total_qty_format)
        write_center(sheet, row, 17, grand_total_balance, company_currency_format)
        for col in range(18, 23):
            write_center(sheet, row, col, '', company_total_qty_format)

        for col, width in enumerate(col_widths):
            sheet.set_column(col, col, width)

        if os.path.exists(footer_path):
            total_width_pixels = sum([w * 8.1 for w in col_widths])

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