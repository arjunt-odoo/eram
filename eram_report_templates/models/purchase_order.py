# -*- coding: utf-8 -*-
from odoo import api, fields, models
from odoo.tools import formatLang

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    street = fields.Char(related="partner_id.street")
    street2 = fields.Char(related="partner_id.street2")
    city = fields.Char(related="partner_id.city")
    state_id = fields.Many2one(related="partner_id.state_id")
    zip = fields.Char(related="partner_id.zip")
    country_id = fields.Many2one(related="partner_id.country_id")
    email = fields.Char(related="partner_id.email")
    e_gst_no = fields.Char(related="partner_id.vat")
    e_rev_no = fields.Char(string="REV.NO.")
    e_supplier_code = fields.Char(string="Supplier #")
    e_req_no = fields.Char(string="REQ.NO.")
    e_supplier_ref = fields.Char(string="Supplier Ref")
    e_mode_of_shipment = fields.Char(string="Mode of Shipment", default="Courier")
    e_delivery_terms = fields.Char(string="Delivery Terms", default="Refer T & C")
    e_ship_to = fields.Char(string="Ship to", default="EPEPL BENGALURU")
    e_delivery_period = fields.Char(string="Delivery Period")
    e_attn = fields.Char(string="Attn")
    e_charge_code = fields.Char(string="Charge Code")
    e_order_line_count = fields.Integer(compute="_compute_e_order_line_count",
                                        store=True, string="No. of Lines")
    e_ref_date = fields.Date(string="REF Date")
    amount_total_words = fields.Char(
        string="TOTAL PURCHASE ORDER VALUE:",
    )
    purchase_cash_rounding_id = fields.Many2one("account.cash.rounding")
    e_date = fields.Date("Date")

    @api.depends('order_line')
    def _compute_e_order_line_count(self):
        for rec in self:
            rec.e_order_line_count = len(rec.order_line)

    def _compute_tax_totals(self):
        """Override to sort tax groups alphabetically by group_name"""
        res = super()._compute_tax_totals()
        for order in self:
            if order.tax_totals and 'subtotals' in order.tax_totals:
                for subtotal in order.tax_totals['subtotals']:
                    if 'tax_groups' in subtotal and subtotal['tax_groups']:
                        subtotal['tax_groups'] = sorted(
                            subtotal['tax_groups'],
                            key=lambda x: x.get('group_name', '').lower()
                        )
        return res

    @api.onchange('e_date')
    def _onchange_e_date(self):
        if self.e_date:
            self.date_order = self.e_date

    def _amount_all(self):
        AccountTax = self.env['account.tax']
        for order in self:
            order_lines = order.order_line.filtered(lambda x: not x.display_type)
            base_lines = [line._prepare_base_line_for_taxes_computation() for line in order_lines]
            AccountTax._add_tax_details_in_base_lines(base_lines, order.company_id)
            AccountTax._round_base_lines_tax_details(base_lines, order.company_id)
            tax_totals = AccountTax._get_tax_totals_summary(
                base_lines=base_lines,
                currency=order.currency_id or order.company_id.currency_id,
                company=order.company_id,
                cash_rounding=order.purchase_cash_rounding_id
            )
            order.amount_untaxed = tax_totals['base_amount_currency']
            order.amount_tax = tax_totals['tax_amount_currency']
            order.amount_total = tax_totals['total_amount_currency']
            order.amount_total_cc = tax_totals['total_amount']

    def _compute_tax_totals(self):
        AccountTax = self.env['account.tax']
        for order in self:
            if not order.company_id:
                order.tax_totals = False
                continue
            order_lines = order.order_line.filtered(lambda x: not x.display_type)
            base_lines = [line._prepare_base_line_for_taxes_computation() for line in order_lines]
            AccountTax._add_tax_details_in_base_lines(base_lines, order.company_id)
            AccountTax._round_base_lines_tax_details(base_lines, order.company_id)
            order.tax_totals = AccountTax._get_tax_totals_summary(
                base_lines=base_lines,
                currency=order.currency_id or order.company_id.currency_id,
                company=order.company_id,
                cash_rounding=order.purchase_cash_rounding_id
            )
            if order.currency_id != order.company_currency_id:
                order.tax_totals['amount_total_cc'] = f"({formatLang(self.env, order.amount_total_cc, currency_obj=self.company_currency_id)})"

class PurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    e_line_index = fields.Integer(
        string='S.NO.',
        compute='_compute_line_index',
        store=True,
        help="The sequential number of this line in the purchase order"
    )

    e_description = fields.Html(string="Description")
    e_charge_code = fields.Html(string="ITEM # / CHARGE CODE")

    @api.depends('order_id', 'order_id.order_line')
    def _compute_line_index(self):
        for order in self.mapped('order_id'):
            for index, line in enumerate(order.order_line, start=1):
                line.e_line_index = index
