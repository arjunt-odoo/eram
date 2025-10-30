# -*- coding: utf-8 -*-
from odoo import api, fields, models


class SaleOrder(models.Model):
    _inherit = 'sale.order'

    e_revision = fields.Char(string="Revision")
    e_attn = fields.Char(string="Attn")
    e_subject = fields.Char(string="Sub")
    e_ref_code = fields.Char(string="REF #")
    e_place = fields.Char(string="Place")
    amount_total_words = fields.Char(
        string="TOTAL QUOTE VALUE:",
    )
    e_packing_and_forwarding = fields.Char(string="PACKING & FORWARDING")
    e_purchase_order_id = fields.Many2one("purchase.order",
                                          string="Ref#")
    e_customer_po_ids = fields.One2many("eram.customer.po", "sale_id")
    e_ref = fields.Char(string="Ref#")

    def _prepare_invoice(self):
        values = super()._prepare_invoice()
        values.update({
            "narration": """<span><u>Bank Details</u></span><br/>
                            <span>Bank Name: HDFC Bank Limited</span><br/>
                            <span>Bank A/C No: 50200064073923</span><br/>
                            <span>Bank IFSC: HDFC0000549</span><br/>
                            <span>Branch: Electronic City</span>""",
        })
        return values

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

    def _prepare_confirmation_values(self):
        vals = super()._prepare_confirmation_values()
        vals.update({
            'date_order': self.date_order
        })
        return vals

class SaleOrderLine(models.Model):
    _inherit = 'sale.order.line'

    e_line_index = fields.Integer(
        string='SL.NO.',
        compute='_compute_line_index',
        store=True,
        help="The sequential number of this line in the purchase order"
    )

    e_description = fields.Html(string="Product Description")

    @api.depends('order_id', 'order_id.order_line')
    def _compute_line_index(self):
        for order in self.mapped('order_id'):
            for index, line in enumerate(order.order_line, start=1):
                line.e_line_index = index
