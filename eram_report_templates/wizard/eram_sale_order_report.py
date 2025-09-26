# -*- coding: utf-8 -*-
from odoo import fields, models
from odoo.tools import json_default
from datetime import datetime
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
    state = fields.Selection(selection=[('sale', 'Sale Order'),
                                        ('draft', 'Quotation')], required=True, default='sale')

    def action_print_report(self):
        if self.type == 'xlsx':
            domain = [('state', '=', self.state)]
            if self.from_date and self.to_date:
                domain.append(('date_order', '>=', self.from_date))
                domain.append(('date_order', '<=', self.to_date))
            elif self.from_date and not self.to_date:
                domain.append(('date_order', '>=', self.from_date))
            elif not self.from_date and self.to_date:
                domain.append(('date_order', '<=', self.to_date))
            records = self.env['sale.order'].search_read(domain, ['id'])
            record_ids = [item.get('id', False) for item in records if item]
            doc_no = self.env['ir.sequence'].next_by_code('eram.sale.order.report',)

            data = {
                'orders': record_ids,
                'doc_no': doc_no,
                'state': self.state
            }
            return {
                'type': 'ir.actions.report',
                'data': {
                    'model': 'eram.sale.order.report',
                    'options': json.dumps(data, default=json_default),
                    'output_format': 'xlsx',
                    'report_name': 'Sales Excel Report',
                },
                'report_type': 'xlsx',
            }

    def get_xlsx_report(self, data, response):
        order_id_list = data.get('orders', [])
        state = data.get('state', '')
        doc_no = data.get('doc_no', '')
        current_datetime = fields.Datetime.context_timestamp(self, fields.Datetime.now())
        formatted_datetime = current_datetime.strftime('%d-%m-%Y %H:%M')
        order_ids = self.env['sale.order'].browse(order_id_list)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        light_green = '#daeef3'
        purple = '#daeef3'
        light_blue = '#f2f2f2'
        red = '#FF0000'

        # Column widths
        col_widths = [
            8,   # A: SI No
            30,  # B: Customer
            30,  # C: Account Name
            20,  # D: Product Details
            40,  # E: Product Description
            12,  # F: Quantity
            15,  # G: Unit Price
            30,  # H: Quote No
            15,  # I: Quote Date
            15,  # J: Quote Value
            25,  # K: Purchase Order No
            15,  # L: PO Date
            15,  # M: PO Value
            22,  # N: Invoice No
            18,  # O: Invoice Date
            22,  # P: Invoice Value
            30,  # Q: Payment Terms
            25,  # R: Payment Status
            22,  # S: Advance Payment Received/Not Received
            25,  # T: Advance Payment Amount
            22,  # U: Advance Payment Received Date
            25,  # V: Balance Payment
            22,  # W: Balance Payment Received Date
            22,  # X: Payment Due
            18,  # Y: Payment Due Date
            25,  # Z: Shipment Status
            18,  # AA: Delivery Date
            25,  # AB: Remarks
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

        info_format = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'text_wrap': True
        })

        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': purple,
            'text_wrap': True
        })
        header_doc_format = workbook.add_format({
            'bold': True,
            'align': 'left',
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
            sheet.insert_image('B2', logo_path, {
                'x_offset': 15,
                'y_offset': 35,
                'x_scale': scale_factor,
                'y_scale': scale_factor,
                'object_position': 3
            })

        report_date = fields.Date.context_today(self)
        formatted_date = report_date.strftime('%d-%m-%Y') if report_date else ''

        sheet.merge_range('B3:AA3', f'SALES ORDER REPORT-{formatted_date}', heading_format)
        sheet.merge_range('AB3:AC3', f'Generated by Odoo    \nDate and Time: {formatted_datetime}    \nDoc No: {doc_no}',
                          header_doc_format)

        sheet.set_row(2, 45)  # Heading row (row 3)

        headers = [
            'SI No', 'Customer', 'Account Name', 'Product Details', 'Product Description',
            'Quantity', 'Unit Price', 'Quote No', 'Quote Date', 'Quote Value',
            'Purchase Order No', 'PO Date', 'PO Value', 'Invoice No', 'Invoice Date',
            'Invoice Value', 'Payment Terms', 'Payment Status', 'Advance Payment Received/Not Received',
            'Advance Payment', 'Advance Payment Received Date', 'Balance Payment',
            'Balance Payment Received Date', 'Payment Due', 'Payment Due Date',
            'Shipment Status', 'Delivery Date', 'Remarks'
        ]

        for col, header in enumerate(headers, 1):
            sheet.write(4, col, header, header_format)

        row = 5
        grand_total_qty = 0
        grand_total_value = 0
        grand_total_advance = 0
        grand_total_balance = 0
        grand_total_unit_price = 0
        grand_total_po_value = 0
        grand_total_quote_value = 0
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
            format_props = {
                'num_format': num_format,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'border': 1
            }
            if is_even:
                format_props['bg_color'] = light_blue
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
            is_even_order = ((first_row - 5) // max_rows_per_order) % 2 == 0
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

        def calculate_row_distribution(product_count, po_count, invoice_count):
            if product_count == 0 and po_count == 0 and invoice_count == 0:
                return [], [], [], 1
            max_count = max(product_count, 1 if po_count == 0 else po_count, invoice_count, 1)
            if product_count == 0:
                product_row_spans = []
            else:
                base_span = max_count // product_count
                remainder = max_count % product_count
                product_row_spans = [base_span + 1 if i < remainder else base_span for i in range(product_count)]
            if po_count == 0:
                po_row_spans = [max_count]  # Merge PO columns into a single row spanning all rows
            else:
                base_span = max_count // po_count
                remainder = max_count % po_count
                po_row_spans = [base_span + 1 if i < remainder else base_span for i in range(po_count)]
            if invoice_count == 0:
                invoice_row_spans = []
            else:
                base_span = max_count // invoice_count
                remainder = max_count % invoice_count
                invoice_row_spans = [base_span + 1 if i < remainder else base_span for i in range(invoice_count)]
            return product_row_spans, po_row_spans, invoice_row_spans, max_count

        for order_idx, order in enumerate(order_ids):
            order_start_row = row
            partner = order.partner_invoice_id or order.partner_id
            full_name = partner.name
            account_name = order.partner_id.name
            current_si_no = si_no
            is_even_order = (order_idx % 2 == 0)

            if order.state == 'sale':
                order_invoices = order.invoice_ids.filtered(lambda i: i.state == 'posted')
            else:
                order_invoices = self.env['account.move']
            order_lines = order.order_line
            purchase_orders = order.e_customer_po_ids

            delivery_date = 'N/A'
            if order.picking_ids and order.state == 'sale':
                first_picking = order.picking_ids[0]
                if first_picking.scheduled_date:
                    if isinstance(first_picking.scheduled_date, str):
                        delivery_date = fields.Date.from_string(first_picking.scheduled_date).strftime('%d-%m-%Y')
                    else:
                        delivery_date = first_picking.scheduled_date.strftime('%d-%m-%Y')

            picking_ids = order.picking_ids.filtered(lambda p: p.state != 'cancel')
            if order.state == 'draft':
                shipment_status = 'Under Production'
            elif picking_ids:
                if all(picking.state == 'done' for picking in picking_ids):
                    shipment_status = 'Shipped'
                elif all(picking.state != 'done' for picking in picking_ids):
                    shipment_status = 'Not Shipped'
                else:
                    shipment_status = 'Partially Shipped'
            else:
                shipment_status = 'Nothing to Ship'

            product_count = len(order_lines)
            po_count = len(purchase_orders)
            invoice_count = len(order_invoices)

            product_row_spans, po_row_spans, invoice_row_spans, max_rows_per_order = calculate_row_distribution(
                product_count, po_count, invoice_count
            )

            merged_ranges = set()

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

            current_row = row
            product_row_index = 0
            po_row_index = 0
            invoice_row_index = 0
            merge_invoice_columns = order.state != 'sale' or invoice_count == 0

            for i in range(max_rows_per_order):
                write_center(sheet, current_row + i, 1, current_si_no if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 2, full_name if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 3, account_name if i == 0 else '', row_fmt)

                write_center(sheet, current_row + i, 8, order.name if i == 0 else '', row_fmt)
                if i == 0 and order.date_order:
                    if isinstance(order.date_order, str):
                        quote_date = fields.Date.from_string(order.date_order).strftime('%d-%m-%Y')
                    else:
                        quote_date = order.date_order.strftime('%d-%m-%Y')
                    write_center(sheet, current_row + i, 9, quote_date, date_fmt)
                else:
                    write_center(sheet, current_row + i, 9, '', row_fmt)
                write_center(sheet, current_row + i, 10, order.amount_total if i == 0 else 0.0, order_currency_format)

                # Write PO information
                if po_count == 0:
                    if i == 0:
                        safe_merge(sheet, current_row, 11, current_row + max_rows_per_order - 1, 11, 'N/A', row_fmt)
                        safe_merge(sheet, current_row, 12, current_row + max_rows_per_order - 1, 12, 'N/A', row_fmt)
                        safe_merge(sheet, current_row, 13, current_row + max_rows_per_order - 1, 13, 0.0, order_currency_format)
                else:
                    if po_row_index < len(purchase_orders):
                        po = purchase_orders[po_row_index]
                        row_span = po_row_spans[po_row_index]
                        po_currency = po.currency_id or order.currency_id
                        po_currency_format = get_currency_format(workbook, po_currency, is_even=is_even_order)
                        if i < sum(po_row_spans[:po_row_index + 1]):
                            po_name = po.name or 'N/A'
                            po_date = 'N/A'
                            if po.date:
                                if isinstance(po.date, str):
                                    po_date = fields.Date.from_string(po.date).strftime('%d-%m-%Y')
                                else:
                                    po_date = po.date.strftime('%d-%m-%Y')
                            po_value = po.amount_total or 0.0
                            if row_span > 1 and i == sum(po_row_spans[:po_row_index]):
                                safe_merge(sheet, current_row + i, 11, current_row + i + row_span - 1, 11, po_name, row_fmt)
                                safe_merge(sheet, current_row + i, 12, current_row + i + row_span - 1, 12, po_date, date_fmt if po_date != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 13, current_row + i + row_span - 1, 13, po_value, po_currency_format)
                            elif row_span == 1:
                                write_center(sheet, current_row + i, 11, po_name, row_fmt)
                                write_center(sheet, current_row + i, 12, po_date, date_fmt if po_date != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 13, po_value, po_currency_format)
                            grand_total_po_value += po_value if i == sum(po_row_spans[:po_row_index]) else 0
                            if i == sum(po_row_spans[:po_row_index + 1]) - 1:
                                po_row_index += 1
                        else:
                            write_center(sheet, current_row + i, 11, '', row_fmt)
                            write_center(sheet, current_row + i, 12, '', date_fmt)
                            write_center(sheet, current_row + i, 13, '', po_currency_format)
                    else:
                        write_center(sheet, current_row + i, 11, '', row_fmt)
                        write_center(sheet, current_row + i, 12, '', date_fmt)
                        write_center(sheet, current_row + i, 13, '', order_currency_format)

                if product_row_index < len(order_lines):
                    line = order_lines[product_row_index]
                    row_span = product_row_spans[product_row_index]
                    if i < sum(product_row_spans[:product_row_index + 1]):
                        if row_span > 1 and i == sum(product_row_spans[:product_row_index]):
                            safe_merge(sheet, current_row + i, 4, current_row + i + row_span - 1, 4,
                                       line.product_id.name, row_fmt)
                            safe_merge(sheet, current_row + i, 5, current_row + i + row_span - 1, 5,
                                       line.e_description or "N/A", row_fmt)
                            safe_merge(sheet, current_row + i, 6, current_row + i + row_span - 1, 6,
                                       line.product_uom_qty, row_fmt)
                            safe_merge(sheet, current_row + i, 7, current_row + i + row_span - 1, 7,
                                       line.price_unit, order_currency_format)
                        elif row_span == 1:
                            write_center(sheet, current_row + i, 4, line.product_id.name, row_fmt)
                            write_center(sheet, current_row + i, 5, line.e_description or "N/A", row_fmt)
                            write_center(sheet, current_row + i, 6, line.product_uom_qty, row_fmt)
                            write_center(sheet, current_row + i, 7, line.price_unit, order_currency_format)
                        grand_total_qty += line.product_uom_qty if i == sum(
                            product_row_spans[:product_row_index]) else 0
                        grand_total_unit_price += line.price_unit if i == sum(
                            product_row_spans[:product_row_index]) else 0
                        if i == sum(product_row_spans[:product_row_index + 1]) - 1:
                            product_row_index += 1
                    else:
                        write_center(sheet, current_row + i, 4, '', row_fmt)
                        write_center(sheet, current_row + i, 5, '', row_fmt)
                        write_center(sheet, current_row + i, 6, '', row_fmt)
                        write_center(sheet, current_row + i, 7, '', row_fmt)

                if merge_invoice_columns:
                    for col in range(14, 26):
                        format_to_use = row_fmt
                        if col in (15, 21, 23, 25):
                            format_to_use = date_fmt
                        elif col in (16, 20, 22):
                            format_to_use = order_currency_format
                        write_center(sheet,current_row + i, col, '', format_to_use)
                else:
                    if invoice_row_index < len(order_invoices):
                        invoice = order_invoices[invoice_row_index]
                        row_span = invoice_row_spans[invoice_row_index]
                        if i < sum(invoice_row_spans[:invoice_row_index + 1]):
                            invoice_amount = invoice.amount_total
                            amount_residual = invoice.amount_residual
                            matched_payments = invoice.matched_payment_ids.filtered(
                                lambda p: p.state in ('in_process', 'paid')
                            )
                            advance_amount = sum(matched_payments.mapped('amount'))
                            latest_payment_date = matched_payments.sorted(key=lambda p: p.date, reverse=True)[
                                0].date if matched_payments else None

                            days_overdue = 0
                            payment_due_display = 'N/A'
                            is_overdue = False
                            if invoice.invoice_date_due:
                                today = fields.Date.today()
                                due_date = fields.Date.from_string(invoice.invoice_date_due)
                                if today > due_date:
                                    days_overdue = (today - due_date).days
                                    payment_due_display = f"{days_overdue} days overdue"
                                    is_overdue = True

                            if invoice.payment_state == 'paid':
                                payment_status = 'Received'
                                advance_payment = 'Received'
                                advance_date = latest_payment_date or invoice.invoice_date
                                balance_payment = 0.0
                                balance_date = 'N/A'
                                payment_due_display = 'N/A'
                            elif invoice.payment_state == 'partial':
                                payment_status = 'Partially Received'
                                advance_payment = 'Received'
                                advance_date = latest_payment_date or 'N/A'
                                balance_payment = amount_residual
                                balance_date = invoice.invoice_date_due
                            else:
                                payment_status = 'Not Received'
                                advance_payment = 'Not Received'
                                advance_date = 'N/A'
                                balance_payment = amount_residual
                                balance_date = invoice.invoice_date_due

                            payment_terms = invoice.invoice_payment_term_id.name or 'N/A'

                            invoice_date = 'N/A'
                            if invoice.invoice_date:
                                if isinstance(invoice.invoice_date, str):
                                    invoice_date = fields.Date.from_string(invoice.invoice_date).strftime('%d-%m-%Y')
                                else:
                                    invoice_date = invoice.invoice_date.strftime('%d-%m-%Y')

                            due_date = 'N/A'
                            if invoice.invoice_date_due:
                                if isinstance(invoice.invoice_date_due, str):
                                    due_date = fields.Date.from_string(invoice.invoice_date_due).strftime('%d-%m-%Y')
                                else:
                                    due_date = invoice.invoice_date_due.strftime('%d-%m-%Y')

                            advance_date_formatted = 'N/A'
                            if advance_date and advance_date != 'N/A':
                                if isinstance(advance_date, str):
                                    advance_date_formatted = fields.Date.from_string(advance_date).strftime('%d-%m-%Y')
                                else:
                                    advance_date_formatted = advance_date.strftime('%d-%m-%Y')

                            balance_date_formatted = 'N/A'
                            if balance_date and balance_date != 'N/A':
                                if isinstance(balance_date, str):
                                    balance_date_formatted = fields.Date.from_string(balance_date).strftime('%d-%m-%Y')
                                else:
                                    balance_date_formatted = balance_date.strftime('%d-%m-%Y')

                            if row_span > 1 and i == sum(invoice_row_spans[:invoice_row_index]):
                                safe_merge(sheet, current_row + i, 14, current_row + i + row_span - 1, 14,
                                           invoice.name or 'N/A', row_fmt)
                                safe_merge(sheet, current_row + i, 15, current_row + i + row_span - 1, 15,
                                           invoice_date, date_fmt if invoice_date != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 16, current_row + i + row_span - 1, 16,
                                           invoice_amount, order_currency_format)
                                safe_merge(sheet, current_row + i, 17, current_row + i + row_span - 1, 17,
                                           payment_terms, row_fmt)
                                safe_merge(sheet, current_row + i, 18, current_row + i + row_span - 1, 18,
                                           payment_status, row_fmt)
                                safe_merge(sheet, current_row + i, 19, current_row + i + row_span - 1, 19,
                                           advance_payment, row_fmt)
                                safe_merge(sheet, current_row + i, 20, current_row + i + row_span - 1, 20,
                                           advance_amount, order_currency_format)
                                safe_merge(sheet, current_row + i, 21, current_row + i + row_span - 1, 21,
                                           advance_date_formatted,
                                           date_fmt if advance_date_formatted != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 22, current_row + i + row_span - 1, 22,
                                           balance_payment, order_currency_format)
                                safe_merge(sheet, current_row + i, 23, current_row + i + row_span - 1, 23,
                                           balance_date_formatted,
                                           date_fmt if balance_date_formatted != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 24, current_row + i + row_span - 1, 24,
                                           payment_due_display,
                                           red_fmt if is_overdue and payment_due_display != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 25, current_row + i + row_span - 1, 25,
                                           due_date, date_fmt if due_date != 'N/A' else row_fmt)
                            elif row_span == 1:
                                write_center(sheet, current_row + i, 14, invoice.name or 'N/A', row_fmt)
                                write_center(sheet, current_row + i, 15, invoice_date,
                                             date_fmt if invoice_date != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 16, invoice_amount, order_currency_format)
                                write_center(sheet, current_row + i, 17, payment_terms, row_fmt)
                                write_center(sheet, current_row + i, 18, payment_status, row_fmt)
                                write_center(sheet, current_row + i, 19, advance_payment, row_fmt)
                                write_center(sheet, current_row + i, 20, advance_amount, order_currency_format)
                                write_center(sheet, current_row + i, 21, advance_date_formatted,
                                             date_fmt if advance_date_formatted != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 22, balance_payment, order_currency_format)
                                write_center(sheet, current_row + i, 23, balance_date_formatted,
                                             date_fmt if balance_date_formatted != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 24, payment_due_display,
                                             red_fmt if is_overdue and payment_due_display != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 25, due_date,
                                             date_fmt if due_date != 'N/A' else row_fmt)
                            grand_total_value += invoice_amount if i == sum(
                                invoice_row_spans[:invoice_row_index]) else 0
                            grand_total_advance += advance_amount if i == sum(
                                invoice_row_spans[:invoice_row_index]) else 0
                            grand_total_balance += balance_payment if i == sum(
                                invoice_row_spans[:invoice_row_index]) else 0
                            if i == sum(invoice_row_spans[:invoice_row_index + 1]) - 1:
                                invoice_row_index += 1
                        else:
                            write_center(sheet, current_row + i, 14, '', row_fmt)
                            write_center(sheet, current_row + i, 15, '', date_fmt)
                            write_center(sheet, current_row + i, 16, '', order_currency_format)
                            write_center(sheet, current_row + i, 17, '', row_fmt)
                            write_center(sheet, current_row + i, 18, '', row_fmt)
                            write_center(sheet, current_row + i, 19, '', order_currency_format)
                            write_center(sheet, current_row + i, 20, '', date_fmt)
                            write_center(sheet, current_row + i, 21, '', order_currency_format)
                            write_center(sheet, current_row + i, 22, '', date_fmt)
                            write_center(sheet, current_row + i, 23, '', row_fmt)
                            write_center(sheet, current_row + i, 24, '', date_fmt)
                            write_center(sheet, current_row + i, 25, '', date_fmt)

                write_center(sheet, current_row + i, 26, shipment_status if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 27, delivery_date if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 28, 'N/A' if i == 0 else '', row_fmt)

            grand_total_quote_value += order.amount_total

            if max_rows_per_order > 1:
                safe_merge(sheet, row, 1, row + max_rows_per_order - 1, 1, current_si_no, row_fmt)
                safe_merge(sheet, row, 2, row + max_rows_per_order - 1, 2, full_name, row_fmt)
                safe_merge(sheet, row, 3, row + max_rows_per_order - 1, 3, account_name, row_fmt)
                safe_merge(sheet, row, 8, row + max_rows_per_order - 1, 8, order.name, row_fmt)
                if order.date_order:
                    if isinstance(order.date_order, str):
                        quote_date = fields.Date.from_string(order.date_order).strftime('%d-%m-%Y')
                    else:
                        quote_date = order.date_order.strftime('%d-%m-%Y')
                    safe_merge(sheet, row, 9, row + max_rows_per_order - 1, 9, quote_date, date_fmt)
                else:
                    safe_merge(sheet, row, 9, row + max_rows_per_order - 1, 9, '', row_fmt)
                safe_merge(sheet, row, 10, row + max_rows_per_order - 1, 10, order.amount_total, order_currency_format)
                safe_merge(sheet, row, 26, row + max_rows_per_order - 1, 26, shipment_status, row_fmt)
                safe_merge(sheet, row, 27, row + max_rows_per_order - 1, 27, delivery_date, row_fmt)
                safe_merge(sheet, row, 28, row + max_rows_per_order - 1, 28, 'N/A', row_fmt)

                if merge_invoice_columns:
                    for col in range(14, 26):
                        format_to_use = row_fmt
                        if col in (15, 21, 23, 25):
                            format_to_use = date_fmt
                        elif col in (16, 20, 22):
                            format_to_use = order_currency_format
                        safe_merge(sheet, row, col, row + max_rows_per_order - 1, col, '', format_to_use)

                # Merge PO columns if no POs
                if po_count == 0:
                    safe_merge(sheet, row, 11, row + max_rows_per_order - 1, 11, 'N/A', row_fmt)
                    safe_merge(sheet, row, 12, row + max_rows_per_order - 1, 12, 'N/A', row_fmt)
                    safe_merge(sheet, row, 13, row + max_rows_per_order - 1, 13, 0.0, order_currency_format)

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

        sheet.merge_range(f'B{row + 1}:F{row + 1}', 'Grand Total:', company_total_qty_format)
        write_center(sheet, row, 6, grand_total_qty, company_total_qty_format)
        write_center(sheet, row, 7, '', company_currency_format)
        write_center(sheet, row, 8, '', company_total_qty_format)
        write_center(sheet, row, 9, '', company_total_qty_format)
        write_center(sheet, row, 10, grand_total_quote_value, company_currency_format)
        write_center(sheet, row, 11, '', company_total_qty_format)
        write_center(sheet, row, 12, '', company_total_qty_format)
        write_center(sheet, row, 13, grand_total_po_value, company_currency_format)
        write_center(sheet, row, 14, '', company_total_qty_format)
        write_center(sheet, row, 15, '', company_total_qty_format)
        write_center(sheet, row, 16, grand_total_value, company_currency_format)
        write_center(sheet, row, 17, '', company_total_qty_format)
        write_center(sheet, row, 18, '', company_total_qty_format)
        write_center(sheet, row, 19, '', company_total_qty_format)
        write_center(sheet, row, 20, grand_total_advance, company_currency_format)
        write_center(sheet, row, 21, '', company_total_qty_format)
        write_center(sheet, row, 22, grand_total_balance, company_currency_format)
        write_center(sheet, row, 23, '', company_total_qty_format)
        write_center(sheet, row, 24, '', company_total_qty_format)
        write_center(sheet, row, 25, '', company_total_qty_format)
        write_center(sheet, row, 26, '', company_total_qty_format)
        write_center(sheet, row, 27, '', company_total_qty_format)
        write_center(sheet, row, 28, '', company_total_qty_format)

        row += 3
        note_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'left',
            'valign': 'top',
            'font_size': 14,
            'font_color': red,
        })
        sheet.merge_range(f'B{row}:AD{row}',
                          'Note: This report has been generated using the Odoo tool by the Eram Power engineers with Python code. The report has been converted into Excel format.',
                          note_format)

        row += 1
        purple_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'font_size': 14,
        })
        sheet.merge_range(f'B{row}:AD{row}', '', purple_format)

        row += 1
        sheet.merge_range(f'B{row}:I{row}', 'Prepared By', purple_format)
        sheet.merge_range(f'J{row}:Q{row}', 'Reviewed By', purple_format)
        sheet.merge_range(f'R{row}:AD{row}', 'Approved By', purple_format)

        row += 1
        sheet.merge_range(f'B{row}:AD{row}', '', purple_format)

        row += 1
        sheet.merge_range(f'B{row}:I{row}', 'Priya Singh', purple_format)
        sheet.merge_range(f'J{row}:Q{row}', 'Sneha John', purple_format)
        sheet.merge_range(f'R{row}:AD{row}', 'Basil Issac', purple_format)

        row += 1
        sheet.merge_range(f'B{row}:I{row}', 'Data Analyst - Sourcing', purple_format)
        sheet.merge_range(f'J{row}:Q{row}', 'Project Manager', purple_format)
        sheet.merge_range(f'R{row}:AD{row}', 'Vice President', purple_format)

        row += 1
        sheet.merge_range(f'B{row}:AD{row}', '', purple_format)

        row += 2

        if os.path.exists(footer_path):
            total_width_pixels = sum([w * 8.1 for w in col_widths])
            with Image.open(footer_path) as img:
                footer_width, footer_height = img.size
            scale_factor = total_width_pixels / 1000
            footer_row = row + 1
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