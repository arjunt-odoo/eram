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

            # Generate unique document number
            current_datetime = datetime.now()
            doc_no = f"DOC_SALE_REPORT_{current_datetime.strftime('%d_%m_%y_%H_%M')}"
            formatted_datetime = current_datetime.strftime('%d-%m-%Y %H:%M')

            data = {
                'orders': sorted_record_ids,
                'doc_no': doc_no,
                'generated_datetime': formatted_datetime
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
        doc_no = data.get('doc_no', '')
        generated_datetime = data.get('generated_datetime', '')
        order_ids = self.env['sale.order'].browse(order_id_list)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        light_green = '#daeef3'
        purple = '#daeef3'
        light_blue = '#D9E1F2'
        light_yellow = '#FFEB9C'
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

        # Adjust range to account for new columns and empty first column
        sheet.merge_range('B3:Z3', f'SALES ORDER REPORT-{formatted_date}', heading_format)
        sheet.merge_range('AA3:AB3', f'Generated by Odoo    Date and Time: {generated_datetime}    Doc No: {doc_no}', header_format)

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

        row = 5  # Start from row 5 (accounting for empty row and header rows)
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

        for order_idx, order in enumerate(order_ids):
            order_start_row = row
            partner = order.partner_invoice_id or order.partner_id
            full_name = partner.name
            account_name = order.partner_id.name
            current_si_no = si_no

            is_even_order = (order_idx % 2 == 0)

            order_invoices = order.invoice_ids.filtered(lambda i: i.state == 'posted')

            order_lines = order.order_line

            # Get purchase order details
            purchase_order = order.e_purchase_order_id
            po_name = purchase_order.name if purchase_order else 'N/A'

            # Format PO date properly
            po_date = 'N/A'
            if purchase_order and purchase_order.e_ref_date:
                if isinstance(purchase_order.e_ref_date, str):
                    po_date = fields.Date.from_string(purchase_order.e_ref_date).strftime('%d-%m-%Y')
                else:
                    po_date = purchase_order.e_ref_date.strftime('%d-%m-%Y')

            po_value = purchase_order.amount_total if purchase_order else 0.0
            po_currency = purchase_order.currency_id if purchase_order else order.currency_id

            # Get delivery date from first delivery's scheduled_date
            delivery_date = 'N/A'
            if order.picking_ids:
                first_picking = order.picking_ids[0]
                if first_picking.scheduled_date:
                    # Convert datetime to date and format as dd-mm-yyyy
                    if isinstance(first_picking.scheduled_date, str):
                        delivery_date = fields.Date.from_string(first_picking.scheduled_date).strftime('%d-%m-%Y')
                    else:
                        delivery_date = first_picking.scheduled_date.strftime('%d-%m-%Y')

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

            # PO currency format
            po_currency_format = get_currency_format(workbook, po_currency, is_even=is_even_order)

            # First, write all the individual cells without merging
            for i in range(max_rows_per_order):
                # Write SI No, Customer, Account Name for each row
                write_center(sheet, row + i, 1, current_si_no if i == 0 else '', row_fmt)
                write_center(sheet, row + i, 2, full_name if i == 0 else '', row_fmt)
                write_center(sheet, row + i, 3, account_name if i == 0 else '', row_fmt)

                # Write Quote information for each row
                write_center(sheet, row + i, 8, order.name if i == 0 else '', row_fmt)  # Quote No
                if i == 0 and order.date_order:
                    # Format quote date properly
                    if isinstance(order.date_order, str):
                        quote_date = fields.Date.from_string(order.date_order).strftime('%d-%m-%Y')
                    else:
                        quote_date = order.date_order.strftime('%d-%m-%Y')
                    write_center(sheet, row + i, 9, quote_date, date_fmt)  # Quote Date
                else:
                    write_center(sheet, row + i, 9, '', row_fmt)

                # Write PO information for each row
                write_center(sheet, row + i, 10, po_name if i == 0 else '', row_fmt)  # Purchase Order No

                # Write PO Date with proper date format
                if i == 0:
                    if po_date != 'N/A':
                        write_center(sheet, row + i, 11, po_date, date_fmt)  # PO Date
                    else:
                        write_center(sheet, row + i, 11, po_date, row_fmt)  # PO Date
                else:
                    write_center(sheet, row + i, 11, '', row_fmt)

                write_center(sheet, row + i, 12, po_value if i == 0 else 0.0,
                             po_currency_format if i == 0 else row_fmt)  # PO Value

                # Write shipment and delivery info for each row
                write_center(sheet, row + i, 25, shipment_status if i == 0 else '', row_fmt)  # Shipment Status
                write_center(sheet, row + i, 26, delivery_date if i == 0 else '', row_fmt)  # Delivery Date
                write_center(sheet, row + i, 27, 'N/A' if i == 0 else '', row_fmt)  # Remarks

            # Write product information
            for product_idx, line in enumerate(order_lines):
                if product_idx < max_rows_per_order:
                    write_center(sheet, row + product_idx, 4, line.product_id.name, row_fmt)
                    write_center(sheet, row + product_idx, 5, line.e_description or "N/A", row_fmt)
                    write_center(sheet, row + product_idx, 6, line.product_uom_qty, row_fmt)
                    write_center(sheet, row + product_idx, 7, line.price_unit, order_currency_format)  # Unit Price
                    grand_total_qty += line.product_uom_qty
                    grand_total_unit_price += line.price_unit

            # Write invoice information
            for invoice_idx, invoice in enumerate(order_invoices):
                if invoice_idx < max_rows_per_order:
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
                        today = fields.Date().today()
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
                    else:
                        advance_date_formatted = advance_date

                    balance_date_formatted = 'N/A'
                    if balance_date and balance_date != 'N/A':
                        if isinstance(balance_date, str):
                            balance_date_formatted = fields.Date.from_string(balance_date).strftime('%d-%m-%Y')
                        else:
                            balance_date_formatted = balance_date.strftime('%d-%m-%Y')
                    else:
                        balance_date_formatted = balance_date

                    # Write invoice data
                    write_center(sheet, row + invoice_idx, 13, buyer_order_no, row_fmt)
                    write_center(sheet, row + invoice_idx, 14, invoice.name or 'N/A', row_fmt)

                    if invoice_date != 'N/A':
                        write_center(sheet, row + invoice_idx, 15, invoice_date, date_fmt)
                    else:
                        write_center(sheet, row + invoice_idx, 15, invoice_date, row_fmt)

                    write_center(sheet, row + invoice_idx, 16, invoice_amount, order_currency_format)
                    write_center(sheet, row + invoice_idx, 17, payment_terms, row_fmt)
                    write_center(sheet, row + invoice_idx, 18, payment_status, row_fmt)
                    write_center(sheet, row + invoice_idx, 19, advance_payment, row_fmt)
                    write_center(sheet, row + invoice_idx, 20, advance_amount, order_currency_format)

                    if advance_date_formatted and advance_date_formatted != 'N/A':
                        write_center(sheet, row + invoice_idx, 21, advance_date_formatted, date_fmt)
                    else:
                        write_center(sheet, row + invoice_idx, 21, advance_date_formatted, row_fmt)

                    write_center(sheet, row + invoice_idx, 22, balance_payment, order_currency_format)

                    if balance_date_formatted and balance_date_formatted != 'N/A':
                        write_center(sheet, row + invoice_idx, 23, balance_date_formatted, date_fmt)
                    else:
                        write_center(sheet, row + invoice_idx, 23, balance_date_formatted, row_fmt)

                    if is_overdue and payment_due_display != 'N/A':
                        write_center(sheet, row + invoice_idx, 24, payment_due_display, red_fmt)
                    else:
                        write_center(sheet, row + invoice_idx, 24, payment_due_display, row_fmt)

                    if due_date != 'N/A':
                        write_center(sheet, row + invoice_idx, 25, due_date, date_fmt)
                    else:
                        write_center(sheet, row + invoice_idx, 25, due_date, row_fmt)

                    grand_total_value += invoice_amount
                    grand_total_advance += advance_amount
                    grand_total_balance += balance_payment
                    grand_total_po_value += po_value if invoice_idx == 0 else 0

            # Now merge the cells that should be merged for this order
            if max_rows_per_order > 1:
                # Merge SI No, Customer, Account Name
                safe_merge(sheet, row, 1, row + max_rows_per_order - 1, 1, current_si_no, row_fmt)
                safe_merge(sheet, row, 2, row + max_rows_per_order - 1, 2, full_name, row_fmt)
                safe_merge(sheet, row, 3, row + max_rows_per_order - 1, 3, account_name, row_fmt)

                # Merge Quote information
                safe_merge(sheet, row, 8, row + max_rows_per_order - 1, 8, order.name, row_fmt)
                if order.date_order:
                    # Format quote date for merging
                    if isinstance(order.date_order, str):
                        quote_date = fields.Date.from_string(order.date_order).strftime('%d-%m-%Y')
                    else:
                        quote_date = order.date_order.strftime('%d-%m-%Y')
                    safe_merge(sheet, row, 9, row + max_rows_per_order - 1, 9, quote_date, date_fmt)
                else:
                    safe_merge(sheet, row, 9, row + max_rows_per_order - 1, 9, '', row_fmt)

                # Merge PO information - PO Date should be merged same as Purchase Order Number
                safe_merge(sheet, row, 10, row + max_rows_per_order - 1, 10, po_name, row_fmt)
                # Always merge PO Date column with proper date format
                if po_date != 'N/A':
                    safe_merge(sheet, row, 11, row + max_rows_per_order - 1, 11, po_date, date_fmt)
                else:
                    safe_merge(sheet, row, 11, row + max_rows_per_order - 1, 11, po_date, row_fmt)
                safe_merge(sheet, row, 12, row + max_rows_per_order - 1, 12, po_value, po_currency_format)

                # Merge shipment and delivery info
                safe_merge(sheet, row, 25, row + max_rows_per_order - 1, 25, shipment_status, row_fmt)
                safe_merge(sheet, row, 26, row + max_rows_per_order - 1, 26, delivery_date, row_fmt)
                safe_merge(sheet, row, 27, row + max_rows_per_order - 1, 27, 'N/A', row_fmt)

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
        write_center(sheet, row, 7, grand_total_unit_price, company_currency_format)  # Unit Price column
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
            'bg_color': purple
        })
        sheet.merge_range(f'B{row}:AB{row}',
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