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
        order_ids = self.env['sale.order'].browse(data.get('orders', []))
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()

        heading_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 36
        })

        bold = workbook.add_format({'bold': True, 'align': 'center'})
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': '#D3D3D3'
        })
        date_format = workbook.add_format(
            {'num_format': 'yyyy-mm-dd', 'align': 'center'})
        currency_format = workbook.add_format(
            {'num_format': '#,##0.00', 'align': 'center'})
        center_format = workbook.add_format({'align': 'center'})
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter'
        })
        merge_date_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': 'yyyy-mm-dd'
        })

        col_widths = [15] * 23

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
        total_width = sum(col_widths)

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
        sheet.merge_range('A2:W2', f'SALES ORDER REPORT-{report_date}',
                          heading_format)

        sheet.set_row(1, 45)

        headers = [
            'Full Name', 'Account Name', 'Product Name', 'Description',
            'Quantity', 'Quote No', 'Quote Date', 'Purchase Order No',
            'Purchase Order Date', 'Invoice No', 'Invoice Date',
            'Invoice Value', 'Payment Terms', 'Payment Status',
            'Advance Payment Received/Not Received', 'Advance Payment Amount',
            'Advance Payment Received Date', 'Balance Payment',
            'Balance Payment Received Date', 'Payment Due',
            'Payment Due Date', 'Shipment Status', 'Remarks'
        ]

        for col, header in enumerate(headers):
            sheet.write(3, col, header, header_format)
            col_widths[col] = max(col_widths[col],
                                  len(header) + 2)

        row = 4
        grand_total_qty = 0
        grand_total_value = 0

        def write_center(sheet, row, col, value, format=center_format):
            sheet.write(row, col, value, format)
            if value is not None:
                col_widths[col] = max(col_widths[col],
                                      len(str(value)) + 2)

        for order in order_ids:
            order_start_row = row
            order_line_count = 0
            partner = order.partner_invoice_id or order.partner_id
            full_name = partner.name
            account_name = order.partner_id.name

            # Track quantities for this order to avoid double-counting
            order_line_quantities = {}

            for line in order.order_line:
                invoice_lines = self.env['account.move.line'].search([
                    ('sale_line_ids', 'in', line.ids),
                    ('parent_state', '!=', 'cancel')
                ])
                invoices = invoice_lines.mapped('move_id')
                order_line_count += len(invoices) if invoices else 1
                # Store the quantity for this line (will be used only once)
                order_line_quantities[line.id] = line.product_uom_qty

            for line in order.order_line:
                invoice_lines = self.env['account.move.line'].search([
                    ('sale_line_ids', 'in', line.ids),
                    ('parent_state', '!=', 'cancel')
                ])

                invoices = invoice_lines.mapped('move_id')

                if not invoices:
                    invoice_data = [{
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
                        'payment_due_date': '-'
                    }]
                else:
                    invoice_data = []
                    for invoice in invoices:
                        invoice_amount = invoice.amount_total
                        amount_residual = invoice.amount_residual

                        advance_amount = sum(invoice.matched_payment_ids.filtered(
                            lambda p: p.state in ('in_process', 'paid')).mapped('amount'))

                        days_overdue = 0
                        payment_due_display = '-'
                        if invoice.invoice_date_due:
                            today = date.today()
                            due_date = fields.Date.from_string(
                                invoice.invoice_date_due)
                            if today > due_date:
                                days_overdue = (today - due_date).days
                                payment_due_display = f"{days_overdue} days overdue"
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
                            else:
                                payment_due_display = '-'

                        invoice_data.append({
                            'invoice_no': invoice.name,
                            'invoice_date': invoice.invoice_date or '-',
                            'invoice_value': invoice_amount,
                            'payment_status': payment_status,
                            'advance_payment': advance_payment,
                            'advance_amount': advance_amount,
                            'advance_date': advance_date,
                            'balance_payment': balance_payment,
                            'balance_date': balance_date,
                            'payment_due': payment_due_display,
                            'payment_due_date': invoice.invoice_date_due or '-'
                        })
                shipment_status = ''
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
                    if i == 0:
                        write_center(sheet, row, 0, full_name)
                        write_center(sheet, row, 1, account_name)
                        write_center(sheet, row, 2, line.product_id.name)
                        write_center(sheet, row, 3, line.e_description)
                        # Use the stored quantity from sale order line
                        line_quantity = order_line_quantities.get(line.id, line.product_uom_qty)
                        write_center(sheet, row, 4, line_quantity, center_format)
                        write_center(sheet, row, 5, order.name)
                        write_center(sheet, row, 6, order.date_order, date_format)
                        write_center(sheet, row, 7, order.e_purchase_order_id.name or '-')
                        if order.e_purchase_order_id and order.e_purchase_order_id.e_ref_date:
                            write_center(sheet, row, 8, order.e_purchase_order_id.e_ref_date, date_format)
                        else:
                            write_center(sheet, row, 8, '-')
                        write_center(sheet, row, 12, order.payment_term_id.name or '100% Against Delivery')
                        write_center(sheet, row, 22, order.note or '-')

                        # Add to grand total only once per line (first invoice)
                        grand_total_qty += line_quantity
                    else:
                        # For subsequent invoices, don't show quantity again
                        write_center(sheet, row, 4, '', center_format)

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

                    write_center(sheet, row, 19, inv['payment_due'])

                    if inv['payment_due_date'] != '-':
                        write_center(sheet, row, 20, inv['payment_due_date'], date_format)
                    else:
                        write_center(sheet, row, 20, inv['payment_due_date'])

                    write_center(sheet, row, 21, shipment_status)

                    grand_total_value += inv['invoice_value']

                    row += 1

                line_end_row = row - 1

                if line_end_row > line_start_row:
                    sheet.merge_range(line_start_row, 2, line_end_row, 2, line.product_id.name, merge_format)
                    sheet.merge_range(line_start_row, 3, line_end_row, 3, line.e_description, merge_format)
                    # Only merge the first cell (where quantity is shown)
                    sheet.merge_range(line_start_row, 4, line_start_row, 4,
                                      order_line_quantities.get(line.id, line.product_uom_qty), merge_format)

            order_end_row = row - 1

            if order_end_row > order_start_row:
                sheet.merge_range(order_start_row, 0, order_end_row, 0, full_name, merge_format)
                sheet.merge_range(order_start_row, 1, order_end_row, 1, account_name, merge_format)
                sheet.merge_range(order_start_row, 5, order_end_row, 5, order.name, merge_format)
                if order.date_order:
                    sheet.merge_range(order_start_row, 6, order_end_row, 6, order.date_order, merge_date_format)
                else:
                    sheet.merge_range(order_start_row, 6, order_end_row, 6, '', merge_format)
                sheet.merge_range(order_start_row, 7, order_end_row, 7, order.e_purchase_order_id.name or '-',
                                  merge_format)
                if order.e_purchase_order_id and order.e_purchase_order_id.e_ref_date:
                    sheet.merge_range(order_start_row, 8, order_end_row, 8, order.e_purchase_order_id.e_ref_date,
                                      merge_date_format)
                else:
                    sheet.merge_range(order_start_row, 8, order_end_row, 8, '-', merge_format)
                sheet.merge_range(order_start_row, 12, order_end_row, 12,
                                  order.payment_term_id.name or '100% Against Delivery', merge_format)
                sheet.merge_range(order_start_row, 22, order_end_row, 22, order.note or '-', merge_format)

        # Write grand totals
        write_center(sheet, row, 0, 'Grand Total', header_format)
        sheet.write(row, 4, grand_total_qty, center_format)
        sheet.write(row, 11, grand_total_value, currency_format)

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