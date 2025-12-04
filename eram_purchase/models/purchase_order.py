# -*- coding: utf-8 -*-
from odoo import _, api, fields, models

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    partner_id = fields.Many2one("res.partner", required=False)
    state = fields.Selection(selection_add=[('draft', "Draft"),
                                            ('sent', "Sent")])
    e_supplier_quote_id = fields.Many2one("eram.supplier.quote", string="Supplier Quote")

    def action_set_as_sent(self):
        self.write({
            'state': 'sent'
        })



