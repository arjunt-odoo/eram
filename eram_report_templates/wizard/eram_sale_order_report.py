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
        light_green = '#daeef3'
        purple = '#daeef3'
        light_blue = '#D9E1F2'
        light_yellow = '#FFEB9C'
        red = '#FF0000'

        # Set column widths based on the sample Excel
        col_widths = [
            8,  # A: SI No
            30,  # B: Full Name
            30,  # C: Account Name
            20,  # D: Product Name
            40,  # E: Description
            12,  # F: Quantity
            30,  # G: Quote No
            15,  # H: Quote Date
            25,  # I: Purchase Order No
            22,  # J: Invoice No
            18,  # K: Invoice Date
            22,  # L: Invoice Value
            30,  # M: Payment Terms
            25,  # N: Payment Status
            22,  # O: Advance Payment Received/Not Received
            25,  # P: Advance Payment Amount
            22,  # Q: Advance Payment Received Date
            25,  # R: Balance Payment
            22,  # S: Balance Payment Received Date
            22,  # T: Payment Due
            18,  # U: Payment Due Date
            25,  # V: Shipment Status
            25,  # W: Remarks
        ]

        # Set column widths (starting from column B since A is empty)
        for col, width in enumerate(col_widths, 1):  # Start from column B (index 1)
            sheet.set_column(col, col, width)

        # Set first column (A) as empty with small width
        sheet.set_column(0, 0, 2)

        heading_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 28,
            'bg_color': light_green,
            'text_wrap': True,
            'border': 1
        })
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': purple,
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

            scale_factor = 0.4
            # Adjust logo position to account for empty first column
            sheet.insert_image('B2', logo_path, {
                'x_offset': 15,
                'y_offset': 35,
                'x_scale': scale_factor,
                'y_scale': scale_factor,
                'object_position': 3
            })

        report_date = fields.Date.context_today(self)
        formatted_date = report_date.strftime('%d-%m-%Y') if report_date else ''
        # Adjust range to account for empty first column
        sheet.merge_range('B3:V3', f'SALES ORDER REPORT-{formatted_date}', heading_format)
        sheet.merge_range('W3:X3', '', header_format)

        sheet.set_row(2, 45)  # Heading row

        headers = [
            'SI No', 'Customer', 'Account Name', 'Product Details', 'Product Description',
            'Quantity', 'Quote No', 'Quote Date', 'Purchase Order No',
            'Invoice No', 'Invoice Date', 'Invoice Value', 'Payment Terms',
            'Payment Status', 'Advance Payment Received/Not Received',
            'Advance Payment', 'Advance Payment Received Date',
            'Balance Payment', 'Balance Payment Received Date', 'Payment Due',
            'Payment Due Date', 'Shipment Status', 'Remarks'
        ]

        # Write headers starting from column B (index 1)
        for col, header in enumerate(headers, 1):
            sheet.write(4, col, header, header_format)

        row = 5  # Start from row 5 (accounting for empty row and header rows)
        grand_total_qty = 0
        grand_total_value = 0
        grand_total_advance = 0
        grand_total_balance = 0
        si_no = 1

        def write_center(sheet, row, col, value, format=None, is_overdue=False):
            is_even_order = ((row - 5) // max_rows_per_order) % 2 == 0

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

            order_invoices = order.invoice_ids.filtered(lambda i: i.state == 'posted')

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

                matched_payments = invoice.matched_payment_ids.filtered(
                    lambda p: p.state in ('in_process', 'paid')
                )

                advance_amount = sum(matched_payments.mapped('amount'))

                if matched_payments:
                    sorted_payments = matched_payments.sorted(key=lambda p: p.date, reverse=True)
                    latest_payment_date = sorted_payments[0].date
                else:
                    latest_payment_date = None

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
                    advance_date = latest_payment_date or invoice.invoice_date
                    balance_payment = 0.0
                    balance_date = 'N/A'
                    payment_due_display = 'N/A'
                elif invoice.payment_state == 'partial':
                    payment_status = 'PARTIALLY RECEIVED'
                    advance_payment = 'Received'
                    advance_date = latest_payment_date or 'N/A'
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
                    # Write to column B (index 1) instead of A (index 0)
                    write_center(sheet, row + i, 1, current_si_no, row_fmt)
                    write_center(sheet, row + i, 2, full_name, row_fmt)
                    write_center(sheet, row + i, 3, account_name, row_fmt)
                    write_center(sheet, row + i, 7, order.name, row_fmt)
                    write_center(sheet, row + i, 8, order.date_order, date_fmt)
                    write_center(sheet, row + i, 22, shipment_status, row_fmt)
                    write_center(sheet, row + i, 23, 'N/A', row_fmt)
                else:
                    write_center(sheet, row + i, 1, '', row_fmt)
                    write_center(sheet, row + i, 2, '', row_fmt)
                    write_center(sheet, row + i, 3, '', row_fmt)
                    write_center(sheet, row + i, 7, '', row_fmt)
                    write_center(sheet, row + i, 8, '', row_fmt)
                    write_center(sheet, row + i, 22, '', row_fmt)
                    write_center(sheet, row + i, 23, 'N/A', row_fmt)

            current_row = 0
            for product_idx, line in enumerate(order_lines):
                rows_for_this_product = product_distribution[product_idx] if product_distribution else 1

                for i in range(rows_for_this_product):
                    if current_row + i < max_rows_per_order:
                        write_center(sheet, row + current_row + i, 4, line.product_id.name, row_fmt)
                        write_center(sheet, row + current_row + i, 5, line.e_description or "N/A", row_fmt)
                        write_center(sheet, row + current_row + i, 6, line.product_uom_qty, row_fmt)
                        grand_total_qty += line.product_uom_qty

                if rows_for_this_product > 1:
                    safe_merge(sheet, row + current_row, 4, row + current_row + rows_for_this_product - 1, 4,
                               line.product_id.name, row_fmt)
                    safe_merge(sheet, row + current_row, 5, row + current_row + rows_for_this_product - 1, 5,
                               line.e_description or "N/A", row_fmt)
                    safe_merge(sheet, row + current_row, 6, row + current_row + rows_for_this_product - 1, 6,
                               line.product_uom_qty, row_fmt)

                current_row += rows_for_this_product

            for i in range(current_row, max_rows_per_order):
                write_center(sheet, row + i, 4, '', row_fmt)
                write_center(sheet, row + i, 5, '', row_fmt)
                write_center(sheet, row + i, 6, '', row_fmt)

            current_row = 0
            if invoice_count > 0:
                for invoice_idx, inv in enumerate(invoice_data):
                    rows_for_this_invoice = invoice_distribution[invoice_idx] if invoice_distribution else 1

                    for i in range(rows_for_this_invoice):
                        if current_row + i < max_rows_per_order:
                            write_center(sheet, row + current_row + i, 9, inv['buyer_order_no'], row_fmt)
                            write_center(sheet, row + current_row + i, 10, inv['invoice_no'], row_fmt)

                            if inv['invoice_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 11, inv['invoice_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 11, inv['invoice_date'], row_fmt)

                            write_center(sheet, row + current_row + i, 12, inv['invoice_value'], order_currency_format)
                            write_center(sheet, row + current_row + i, 13, inv['payment_terms'], row_fmt)
                            write_center(sheet, row + current_row + i, 14, inv['payment_status'], row_fmt)
                            write_center(sheet, row + current_row + i, 15, inv['advance_payment'], row_fmt)
                            write_center(sheet, row + current_row + i, 16, inv['advance_amount'], order_currency_format)

                            if inv['advance_date'] and inv['advance_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 17, inv['advance_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 17, inv['advance_date'], row_fmt)

                            write_center(sheet, row + current_row + i, 18, inv['balance_payment'],
                                         order_currency_format)

                            if inv['balance_date'] and inv['balance_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 19, inv['balance_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 19, inv['balance_date'], row_fmt)

                            if inv['is_overdue'] and inv['payment_due'] != 'N/A':
                                write_center(sheet, row + current_row + i, 20, inv['payment_due'], red_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 20, inv['payment_due'], row_fmt)

                            if inv['payment_due_date'] != 'N/A':
                                write_center(sheet, row + current_row + i, 21, inv['payment_due_date'], date_fmt)
                            else:
                                write_center(sheet, row + current_row + i, 21, inv['payment_due_date'], row_fmt)

                            grand_total_value += inv['invoice_value'] if i == 0 else 0
                            grand_total_advance += inv['advance_amount'] if i == 0 else 0
                            grand_total_balance += inv['balance_payment'] if i == 0 else 0

                    if rows_for_this_invoice > 1:
                        safe_merge(sheet, row + current_row, 9, row + current_row + rows_for_this_invoice - 1, 9,
                                   inv['buyer_order_no'], row_fmt)
                        safe_merge(sheet, row + current_row, 10, row + current_row + rows_for_this_invoice - 1, 10,
                                   inv['invoice_no'], row_fmt)
                        safe_merge(sheet, row + current_row, 11, row + current_row + rows_for_this_invoice - 1, 11,
                                   inv['invoice_date'], date_fmt if inv['invoice_date'] != 'N/A' else row_fmt)
                        safe_merge(sheet, row + current_row, 12, row + current_row + rows_for_this_invoice - 1, 12,
                                   inv['invoice_value'], order_currency_format)
                        safe_merge(sheet, row + current_row, 13, row + current_row + rows_for_this_invoice - 1, 13,
                                   inv['payment_terms'], row_fmt)
                        safe_merge(sheet, row + current_row, 14, row + current_row + rows_for_this_invoice - 1, 14,
                                   inv['payment_status'], row_fmt)
                        safe_merge(sheet, row + current_row, 15, row + current_row + rows_for_this_invoice - 1, 15,
                                   inv['advance_payment'], row_fmt)
                        safe_merge(sheet, row + current_row, 16, row + current_row + rows_for_this_invoice - 1, 16,
                                   inv['advance_amount'], order_currency_format)
                        safe_merge(sheet, row + current_row, 17, row + current_row + rows_for_this_invoice - 1, 17,
                                   inv['advance_date'],
                                   date_fmt if inv['advance_date'] and inv['advance_date'] != 'N/A' else row_fmt)
                        safe_merge(sheet, row + current_row, 18, row + current_row + rows_for_this_invoice - 1, 18,
                                   inv['balance_payment'], order_currency_format)
                        safe_merge(sheet, row + current_row, 19, row + current_row + rows_for_this_invoice - 1, 19,
                                   inv['balance_date'],
                                   date_fmt if inv['balance_date'] and inv['balance_date'] != 'N/A' else row_fmt)
                        safe_merge(sheet, row + current_row, 20, row + current_row + rows_for_this_invoice - 1, 20,
                                   inv['payment_due'],
                                   red_fmt if inv['is_overdue'] and inv['payment_due'] != 'N/A' else row_fmt)
                        safe_merge(sheet, row + current_row, 21, row + current_row + rows_for_this_invoice - 1, 21,
                                   inv['payment_due_date'], date_fmt if inv['payment_due_date'] != 'N/A' else row_fmt)

                    current_row += rows_for_this_invoice
            else:
                # If no invoices, merge all invoice-related columns into a single blank cell
                safe_merge(sheet, row, 9, row + max_rows_per_order - 1, 21, '', row_fmt)

            for i in range(current_row, max_rows_per_order):
                for col in range(9, 22):
                    write_center(sheet, row + i, col, '', row_fmt)

            order_end_row = row + max_rows_per_order - 1
            if order_end_row > order_start_row:
                safe_merge(sheet, order_start_row, 1, order_end_row, 1, current_si_no, row_fmt)
                safe_merge(sheet, order_start_row, 2, order_end_row, 2, full_name, row_fmt)
                safe_merge(sheet, order_start_row, 3, order_end_row, 3, account_name, row_fmt)
                safe_merge(sheet, order_start_row, 7, order_end_row, 7, order.name, row_fmt)

                if order.date_order:
                    safe_merge(sheet, order_start_row, 8, order_end_row, 8, order.date_order, date_fmt)
                else:
                    safe_merge(sheet, order_start_row, 8, order_end_row, 8, '', row_fmt)

                safe_merge(sheet, order_start_row, 22, order_end_row, 22, shipment_status, row_fmt)
                safe_merge(sheet, order_start_row, 23, order_end_row, 23, 'N/A', row_fmt)

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

        # Add empty row before grand total
        # Write grand total starting from column B (index 1)
        # write_center(sheet, row, 1, 'Grand Total', company_total_qty_format)
        sheet.merge_range(f'B{row+1}:F{row+1}', 'Grand Total:', company_total_qty_format)
        # for col in range(2, 6):  # Columns C to F
        #     write_center(sheet, row, col, '', company_total_qty_format)
        write_center(sheet, row, 6, grand_total_qty, company_total_qty_format)  # Quantity column
        for col in range(7, 12):  # Columns H to L
            write_center(sheet, row, col, '', company_total_qty_format)
        write_center(sheet, row, 12, grand_total_value, company_currency_format)  # Invoice Value column
        write_center(sheet, row, 13, '', company_total_qty_format)  # Payment Terms
        write_center(sheet, row, 14, '', company_total_qty_format)  # Payment Status
        write_center(sheet, row, 15, '', company_total_qty_format)  # Advance Payment Received/Not Received
        write_center(sheet, row, 16, grand_total_advance, company_currency_format)  # Advance Payment Amount
        write_center(sheet, row, 17, '', company_total_qty_format)  # Advance Payment Received Date
        write_center(sheet, row, 18, grand_total_balance, company_currency_format)  # Balance Payment
        for col in range(19, 24):  # Columns T to X
            write_center(sheet, row, col, '', company_total_qty_format)

        # Add additional space above footer with text
        row += 3
        note_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'left',
            'valign': 'top',
            'font_size': 14,
            'font_color': red,
            'bg_color': purple
        })
        sheet.merge_range(f'B{row}:X{row}',
                          'Note: This is a system generated report. For any discrepancies, please contact the accounts department.',
                          note_format)

        row += 1
        purple_format = workbook.add_format({
            'bold': True,
            'bg_color': purple,
            'text_wrap': True,
            'align': 'center',
            'font_size': 14,
        })
        sheet.merge_range(f'B{row}:X{row}', '', purple_format)

        # Add Prepared By and Reviewed By headers
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Prepared By', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Reviewed By', purple_format)
        sheet.merge_range(f'O{row}:X{row}', 'Approved By', purple_format)

        # Add empty row with purple background color
        row += 1
        sheet.merge_range(f'B{row}:X{row}', '', purple_format)

        # Add names
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Priya Singh', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Sneha John', purple_format)
        sheet.merge_range(f'O{row}:X{row}', 'Basil Issac', purple_format)

        # Add designations
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Data Analyst - Sourcing', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Project Manager', purple_format)
        sheet.merge_range(f'O{row}:X{row}', 'Vice President', purple_format)

        # Add empty row with purple background color
        row += 1
        sheet.merge_range(f'B{row}:X{row}', '', purple_format)

        # Add more space before footer
        row += 2

        if os.path.exists(footer_path):
            total_width_pixels = sum([w * 8.1 for w in col_widths])

            with Image.open(footer_path) as img:
                footer_width, footer_height = img.size

            scale_factor = total_width_pixels / 1000

            footer_row = row + 1
            # Adjust footer position to account for empty first column
            sheet.insert_image(
                f'B{footer_row}',
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