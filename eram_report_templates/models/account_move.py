# -*- coding: utf-8 -*-
import logging
import time
from datetime import date
from odoo import api, fields, models

_logger = logging.getLogger(__name__)


class AccountMove(models.Model):
    _inherit = 'account.move'
    _sql_constraints = [('name_unique', 'unique(company_id, name, partner_id)',
                         "Invoice name should be unique!")]

    e_transport_mode = fields.Char(string="Transport Mode")
    e_reverse_charge = fields.Boolean(string="Reverse Charge")
    e_buyer_order_no = fields.Char(string="Buyer's Order No")
    e_shipment_terms = fields.Char(string="Shipment Terms")
    e_ship_to_party = fields.Many2one("res.partner", string="Ship to Party")
    e_bill_tel = fields.Char(related="partner_id.phone", string="Tel")
    e_bill_email = fields.Char(related="partner_id.email", string="E-mail")
    e_bill_vat = fields.Char(related="partner_id.vat", string="GSTIN")
    e_ship_tel = fields.Char(related="e_ship_to_party.phone", string="Tel")
    e_ship_email = fields.Char(related="e_ship_to_party.email", string="E-mail")
    e_ship_vat = fields.Char(related="e_ship_to_party.vat", string="GSTIN")
    narration = fields.Html(default="""<span><u>Bank Details</u></span><br/>
        <span>Bank Name: HDFC Bank Limited</span><br/>
        <span>Bank A/C No: 50200064073923</span><br/>
        <span>Bank IFSC: HDFC0000549</span><br/>
        <span>Branch: Electronic City</span>""")
    invoice_date = fields.Date(default=fields.Date.context_today)
    e_invoice_received_date = fields.Date(default=fields.Date.context_today,
                                          string="Received Date")
    e_amount_total_words = fields.Char(
        string="Total amount in words",
    )
    e_sequence = fields.Integer("Sequence", defualt=1,
                                help="The order in which the invoices are taken in sale order report")

    def _compute_tax_totals(self):
        """Override to sort tax groups alphabetically by group_name"""
        res = super()._compute_tax_totals()
        for move in self:
            if move.is_invoice(include_receipts=True) and move.tax_totals:
                if 'subtotals' in move.tax_totals:
                    for subtotal in move.tax_totals['subtotals']:
                        if 'tax_groups' in subtotal and subtotal['tax_groups']:
                            subtotal['tax_groups'] = sorted(
                                subtotal['tax_groups'],
                                key=lambda x: x.get('group_name', '').lower()
                            )

        return res

    def _onchange_name_warning(self):
        pass

    def eram_alert_invoice_overdue(self):
        today = date.today()
        overdue_invoices = self.search([
            ('invoice_date_due', '<', today),
            ('state', '=', 'posted'),
            ('move_type','=', 'out_invoice'),
            ('invoice_date_due', '!=', False),
            ('amount_residual', '>', 0)
        ])
        overdue_invoices = overdue_invoices.filtered(lambda i: i.status_in_payment != 'paid')
        if overdue_invoices:
            overdue_invoices._send_due_date_alert_notification()

        return True

    def _send_due_date_alert_notification(self):
        """
        Send email notification to user when invoice due date is less than today
        """
        today = date.today()
        overdue_invoices = self.filtered(
            lambda inv: inv.invoice_date_due and
                        inv.invoice_date_due < today and
                        inv.state == 'posted' and
                        inv.move_type == 'out_invoice'
        )

        for invoice in overdue_invoices:
            try:
                recipient_user = invoice.invoice_user_id or invoice.create_uid or self.env.user

                if recipient_user and recipient_user.email:
                    mail_template = self.env.ref('eram_report_templates.eram_email_template_invoice_overdue')

                    if not mail_template:
                        return
                    mail_template.send_mail(invoice.id, force_send=True)

                    invoice.message_post(
                        body=f"Overdue notification email sent to {recipient_user.name}",
                        message_type='comment',
                        subtype_xmlid='mail.mt_comment'
                    )
                    time.sleep(0.5)

            except Exception as e:
                _logger.error("Failed to send overdue notification for invoice %s: %s", invoice.name, str(e))

    def _constrains_date_sequence(self):
        pass

class AccountMoveLine(models.Model):
    _inherit = 'account.move.line'

    e_line_index = fields.Integer(
        string='SI.No.',
        compute='_compute_line_index',
        store=True,
        help="The sequential number of this line in the purchase order"
    )

    e_description = fields.Html(string="Description")

    @api.depends('move_id', 'move_id.invoice_line_ids')
    def _compute_line_index(self):
        for move in self.mapped('move_id'):
            for index, line in enumerate(move.invoice_line_ids, start=1):
                line.e_line_index = index
