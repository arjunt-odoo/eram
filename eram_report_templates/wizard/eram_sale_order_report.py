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
        current_datetime = fields.Datetime.context_timestamp(self, fields.Datetime().now())
        formatted_datetime = current_datetime.strftime('%d-%m-%Y %H:%M')
        order_ids = self.env['sale.order'].browse(order_id_list)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        light_green = '#daeef3'
        purple = '#daeef3'
        light_blue = '#D9E1F2'
        red = '#FF0000'

        # Set column widths based on the sample Excel - updated with new columns
        col_widths = [
            8,  # A: SI No
            30,  # B: Full Name
            30,  # C: Account Name
            20,  # D: Product Name
            40,  # E: Description
            12,  # F: Quantity
            15,  # G: Unit Price (NEW COLUMN)
            30,  # H: Quote No
            15,  # I: Quote Date
            25,  # J: Purchase Order No
            15,  # K: PO Date (NEW COLUMN)
            20,  # L: PO Value (NEW COLUMN)
            22,  # M: Invoice No
            18,  # N: Invoice Date
            22,  # O: Invoice Value
            30,  # P: Payment Terms
            25,  # Q: Payment Status
            22,  # R: Advance Payment Received/Not Received
            25,  # S: Advance Payment Amount
            22,  # T: Advance Payment Received Date
            25,  # U: Balance Payment
            22,  # V: Balance Payment Received Date
            22,  # W: Payment Due
            18,  # X: Payment Due Date
            25,  # Y: Shipment Status
            18,  # Z: Delivery Date (NEW COLUMN)
            25,  # AA: Remarks
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

        # Format for the generation info row
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
            # Adjust logo position to account for empty first column
            sheet.insert_image('B2', logo_path, {
                'x_offset': 15,
                'y_offset': 35,
                'x_scale': scale_factor,
                'y_scale': scale_factor,
                'object_position': 3
            })

        report_date = fields.Date().context_today(self)
        formatted_date = report_date.strftime('%d-%m-%Y') if report_date else ''

        # Start from row 3 instead of row 4 to remove one empty row
        sheet.merge_range('B3:Z3', f'SALES ORDER REPORT-{formatted_date}', heading_format)
        sheet.merge_range('AA3:AB3', f'Generated by Odoo    \nDate and Time: {formatted_datetime}    \nDoc No: {doc_no}',
                          header_doc_format)

        sheet.set_row(2, 45)  # Heading row (row 3)

        # Updated headers with new columns
        headers = [
            'SI No', 'Customer', 'Account Name', 'Product Details', 'Product Description',
            'Quantity', 'Unit Price', 'Quote No', 'Quote Date', 'Purchase Order No',
            'PO Date', 'PO Value', 'Invoice No', 'Invoice Date', 'Invoice Value',
            'Payment Terms', 'Payment Status', 'Advance Payment Received/Not Received',
            'Advance Payment', 'Advance Payment Received Date', 'Balance Payment',
            'Balance Payment Received Date', 'Payment Due', 'Payment Due Date',
            'Shipment Status', 'Delivery Date', 'Remarks'
        ]

        # Write headers starting from column B (index 1) in row 4
        for col, header in enumerate(headers, 1):
            sheet.write(4, col, header, header_format)

        row = 5  # Start from row 5
        grand_total_qty = 0
        grand_total_value = 0
        grand_total_advance = 0
        grand_total_balance = 0
        grand_total_unit_price = 0
        grand_total_po_value = 0
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

        def calculate_row_distribution(product_count, invoice_count):
            """Calculate how many rows each product and invoice should span"""
            if product_count == 0 and invoice_count == 0:
                return [], [], 1  # Allocate one row for orders with no products or invoices

            elif product_count == 0:
                return [], [1] * invoice_count, invoice_count

            elif invoice_count == 0:
                return [1] * product_count, [], product_count

            else:
                max_rows = max(product_count, invoice_count)
                base_span = max_rows // min(product_count, invoice_count)
                remainder = max_rows % min(product_count, invoice_count)

                if product_count <= invoice_count:
                    product_row_spans = [base_span + 1 if i < remainder else base_span
                                         for i in range(product_count)]
                    invoice_row_spans = [1] * invoice_count
                else:
                    product_row_spans = [1] * product_count
                    invoice_row_spans = [base_span + 1 if i < remainder else base_span
                                         for i in range(invoice_count)]

                return product_row_spans, invoice_row_spans, max_rows

        # Inside the get_xlsx_report method, replace the loop for writing order data with this:

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

            # Get purchase order details
            purchase_order = order.e_customer_po_id
            po_name = purchase_order.name if purchase_order else 'N/A'

            # Format PO date properly
            po_date = 'N/A'
            if purchase_order and purchase_order.date:
                if isinstance(purchase_order.date, str):
                    po_date = fields.Date.from_string(purchase_order.date).strftime('%d-%m-%Y')
                else:
                    po_date = purchase_order.date.strftime('%d-%m-%Y')

            po_value = purchase_order.amount if purchase_order else 0.0
            po_currency = purchase_order.currency_id if purchase_order else order.currency_id

            # Get delivery date from first delivery's scheduled_date
            delivery_date = 'N/A'
            if order.picking_ids and order.state == 'sale':
                first_picking = order.picking_ids[0]
                if first_picking.scheduled_date:
                    if isinstance(first_picking.scheduled_date, str):
                        delivery_date = fields.Date.from_string(first_picking.scheduled_date).strftime('%d-%m-%Y')
                    else:
                        delivery_date = first_picking.scheduled_date.strftime('%d-%m-%Y')

            # Get shipment status from sale order
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
            invoice_count = len(order_invoices)

            # Calculate row distribution
            product_row_spans, invoice_row_spans, max_rows_per_order = calculate_row_distribution(
                product_count, invoice_count
            )

            # Reset merged ranges for each order
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
            po_currency_format = get_currency_format(workbook, po_currency, is_even=is_even_order)

            # Write all rows for the order
            current_row = row
            product_row_index = 0
            invoice_row_index = 0

            # Check if invoice columns should be merged (state != 'sale' or no invoices)
            merge_invoice_columns = order.state != 'sale' or invoice_count == 0

            for i in range(max_rows_per_order):
                # Write common fields (SI No, Customer, Account Name, Quote, PO, Shipment, Delivery, Remarks)
                write_center(sheet, current_row + i, 1, current_si_no if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 2, full_name if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 3, account_name if i == 0 else '', row_fmt)

                # Write Quote information
                write_center(sheet, current_row + i, 8, order.name if i == 0 else '', row_fmt)  # Quote No
                if i == 0 and order.date_order:
                    if isinstance(order.date_order, str):
                        quote_date = fields.Date.from_string(order.date_order).strftime('%d-%m-%Y')
                    else:
                        quote_date = order.date_order.strftime('%d-%m-%Y')
                    write_center(sheet, current_row + i, 9, quote_date, date_fmt)
                else:
                    write_center(sheet, current_row + i, 9, '', row_fmt)

                # Write PO information
                write_center(sheet, current_row + i, 10, po_name if i == 0 else '', row_fmt)
                if i == 0:
                    if po_date != 'N/A':
                        write_center(sheet, current_row + i, 11, po_date, date_fmt)
                    else:
                        write_center(sheet, current_row + i, 11, po_date, row_fmt)
                else:
                    write_center(sheet, current_row + i, 11, '', row_fmt)
                write_center(sheet, current_row + i, 12, po_value if i == 0 else 0.0,
                             po_currency_format if i == 0 else row_fmt)

                # Write shipment and delivery info
                write_center(sheet, current_row + i, 25, shipment_status if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 26, delivery_date if i == 0 else '', row_fmt)
                write_center(sheet, current_row + i, 27, 'N/A' if i == 0 else '', row_fmt)

                # Write product information for this row
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

                # Write invoice information or merge invoice columns
                if merge_invoice_columns:
                    # If state != 'sale' or no invoices, merge invoice columns for this row
                    for col in range(13, 25):  # Columns M to X
                        format_to_use = row_fmt
                        if col in (14, 20, 22, 24):  # Date columns (Invoice Date, Advance Date, Balance Date, Due Date)
                            format_to_use = date_fmt
                        elif col in (15, 19, 21):  # Currency columns (Invoice Value, Advance Amount, Balance Payment)
                            format_to_use = order_currency_format
                        write_center(sheet, current_row + i, col, '', format_to_use)
                else:
                    # Write invoice information for this row
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
                            else:
                                payment_status = 'NOT RECEIVED'
                                advance_payment = 'Not Received'
                                advance_date = 'N/A'
                                balance_payment = amount_residual
                                balance_date = invoice.invoice_date_due

                            payment_terms = invoice.invoice_payment_term_id.name or 'N/A'

                            # Format invoice dates properly
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
                                safe_merge(sheet, current_row + i, 13, current_row + i + row_span - 1, 13,
                                           invoice.name or 'N/A', row_fmt)
                                safe_merge(sheet, current_row + i, 14, current_row + i + row_span - 1, 14,
                                           invoice_date, date_fmt if invoice_date != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 15, current_row + i + row_span - 1, 15,
                                           invoice_amount, order_currency_format)
                                safe_merge(sheet, current_row + i, 16, current_row + i + row_span - 1, 16,
                                           payment_terms, row_fmt)
                                safe_merge(sheet, current_row + i, 17, current_row + i + row_span - 1, 17,
                                           payment_status, row_fmt)
                                safe_merge(sheet, current_row + i, 18, current_row + i + row_span - 1, 18,
                                           advance_payment, row_fmt)
                                safe_merge(sheet, current_row + i, 19, current_row + i + row_span - 1, 19,
                                           advance_amount, order_currency_format)
                                safe_merge(sheet, current_row + i, 20, current_row + i + row_span - 1, 20,
                                           advance_date_formatted,
                                           date_fmt if advance_date_formatted != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 21, current_row + i + row_span - 1, 21,
                                           balance_payment, order_currency_format)
                                safe_merge(sheet, current_row + i, 22, current_row + i + row_span - 1, 22,
                                           balance_date_formatted,
                                           date_fmt if balance_date_formatted != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 23, current_row + i + row_span - 1, 23,
                                           payment_due_display,
                                           red_fmt if is_overdue and payment_due_display != 'N/A' else row_fmt)
                                safe_merge(sheet, current_row + i, 24, current_row + i + row_span - 1, 24,
                                           due_date, date_fmt if due_date != 'N/A' else row_fmt)
                            elif row_span == 1:
                                write_center(sheet, current_row + i, 13, invoice.name or 'N/A', row_fmt)
                                write_center(sheet, current_row + i, 14, invoice_date,
                                             date_fmt if invoice_date != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 15, invoice_amount, order_currency_format)
                                write_center(sheet, current_row + i, 16, payment_terms, row_fmt)
                                write_center(sheet, current_row + i, 17, payment_status, row_fmt)
                                write_center(sheet, current_row + i, 18, advance_payment, row_fmt)
                                write_center(sheet, current_row + i, 19, advance_amount, order_currency_format)
                                write_center(sheet, current_row + i, 20, advance_date_formatted,
                                             date_fmt if advance_date_formatted != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 21, balance_payment, order_currency_format)
                                write_center(sheet, current_row + i, 22, balance_date_formatted,
                                             date_fmt if balance_date_formatted != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 23, payment_due_display,
                                             red_fmt if is_overdue and payment_due_display != 'N/A' else row_fmt)
                                write_center(sheet, current_row + i, 24, due_date,
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
                            # Write empty invoice columns with appropriate background color
                            write_center(sheet, current_row + i, 13, '', row_fmt)  # Invoice No
                            write_center(sheet, current_row + i, 14, '', date_fmt)  # Invoice Date
                            write_center(sheet, current_row + i, 15, '', order_currency_format)  # Invoice Value
                            write_center(sheet, current_row + i, 16, '', row_fmt)  # Payment Terms
                            write_center(sheet, current_row + i, 17, '', row_fmt)  # Payment Status
                            write_center(sheet, current_row + i, 18, '',
                                         row_fmt)  # Advance Payment Received/Not Received
                            write_center(sheet, current_row + i, 19, '',
                                         order_currency_format)  # Advance Payment Amount
                            write_center(sheet, current_row + i, 20, '', date_fmt)  # Advance Payment Received Date
                            write_center(sheet, current_row + i, 21, '', order_currency_format)  # Balance Payment
                            write_center(sheet, current_row + i, 22, '', date_fmt)  # Balance Payment Received Date
                            write_center(sheet, current_row + i, 23, '', row_fmt)  # Payment Due
                            write_center(sheet, current_row + i, 24, '', date_fmt)  # Payment Due Date

            grand_total_po_value += po_value

            # Merge common cells for orders with multiple rows
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
                safe_merge(sheet, row, 10, row + max_rows_per_order - 1, 10, po_name, row_fmt)
                if po_date != 'N/A':
                    safe_merge(sheet, row, 11, row + max_rows_per_order - 1, 11, po_date, date_fmt)
                else:
                    safe_merge(sheet, row, 11, row + max_rows_per_order - 1, 11, po_date, row_fmt)
                safe_merge(sheet, row, 12, row + max_rows_per_order - 1, 12, po_value, po_currency_format)
                safe_merge(sheet, row, 25, row + max_rows_per_order - 1, 25, shipment_status, row_fmt)
                safe_merge(sheet, row, 26, row + max_rows_per_order - 1, 26, delivery_date, row_fmt)
                safe_merge(sheet, row, 27, row + max_rows_per_order - 1, 27, 'N/A', row_fmt)

                # Merge invoice columns if state != 'sale' or no invoices
                if merge_invoice_columns:
                    for col in range(13, 25):  # Columns M to X
                        format_to_use = row_fmt
                        if col in (14, 20, 22, 24):  # Date columns
                            format_to_use = date_fmt
                        elif col in (15, 19, 21):  # Currency columns
                            format_to_use = order_currency_format
                        safe_merge(sheet, row, col, row + max_rows_per_order - 1, col, '', format_to_use)

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

        # Add grand total row
        sheet.merge_range(f'B{row + 1}:F{row + 1}', 'Grand Total:', company_total_qty_format)
        write_center(sheet, row, 6, grand_total_qty, company_total_qty_format)  # Quantity column
        # write_center(sheet, row, 7, grand_total_unit_price, company_currency_format)  # Unit Price column
        write_center(sheet, row, 7, '', company_currency_format)  # Unit Price column
        for col in range(8, 12):  # Columns H to K
            write_center(sheet, row, col, '', company_total_qty_format)
        write_center(sheet, row, 12, grand_total_po_value, company_currency_format)  # PO Value column
        write_center(sheet, row, 13, '', company_total_qty_format)  # Invoice No
        write_center(sheet, row, 14, '', company_total_qty_format)  # Invoice Date
        write_center(sheet, row, 15, grand_total_value, company_currency_format)  # Invoice Value column
        write_center(sheet, row, 16, '', company_total_qty_format)  # Payment Terms
        write_center(sheet, row, 17, '', company_total_qty_format)  # Payment Status
        write_center(sheet, row, 18, '', company_total_qty_format)  # Advance Payment Received/Not Received
        write_center(sheet, row, 19, grand_total_advance, company_currency_format)  # Advance Payment Amount
        write_center(sheet, row, 20, '', company_total_qty_format)  # Advance Payment Received Date
        write_center(sheet, row, 21, grand_total_balance, company_currency_format)  # Balance Payment
        for col in range(22, 28):  # Columns V to AB
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
            # 'bg_color': purple
        })
        sheet.merge_range(f'B{row}:AB{row}',
                          'Note: This report has been generated using the Odoo tool by the Eram Power engineers with Python code. The report has been converted into Excel format.',
                          note_format)

        row += 1
        purple_format = workbook.add_format({
            'bold': True,
            # 'bg_color': purple,
            'text_wrap': True,
            'align': 'center',
            'font_size': 14,
        })
        sheet.merge_range(f'B{row}:AB{row}', '', purple_format)

        # Add Prepared By and Reviewed By headers
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Prepared By', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Reviewed By', purple_format)
        sheet.merge_range(f'O{row}:AB{row}', 'Approved By', purple_format)

        # Add empty row with purple background color
        row += 1
        sheet.merge_range(f'B{row}:AB{row}', '', purple_format)

        # Add names
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Priya Singh', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Sneha John', purple_format)
        sheet.merge_range(f'O{row}:AB{row}', 'Basil Issac', purple_format)

        # Add designations
        row += 1
        sheet.merge_range(f'B{row}:G{row}', 'Data Analyst - Sourcing', purple_format)
        sheet.merge_range(f'H{row}:N{row}', 'Project Manager', purple_format)
        sheet.merge_range(f'O{row}:AB{row}', 'Vice President', purple_format)

        # Add empty row with purple background color
        row += 1
        sheet.merge_range(f'B{row}:AB{row}', '', purple_format)

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