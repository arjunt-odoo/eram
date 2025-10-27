# -*- coding: utf-8 -*-
from odoo import _, api, fields, models

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    partner_id = fields.Many2one("res.partner", required=False)
    eram_pr_id = fields.Many2one("eram.purchase.req")
    e_our_ref = fields.Char("Our REF")
    e_our_bid_closing_date = fields.Date("Our Bid Closing Date")
    e_send_offer_by = fields.Date("Send your Offer by")
    e_quote_ids = fields.One2many("eram.supplier.quote", "purchase_id")


    def action_view_pr(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Purchase Request"),
            'res_model': 'eram.purchase.req',
            'res_id': self.eram_pr_id.id,
            'view_mode': 'form'
        }

    def action_set_as_sent(self):
        self.write({
            'state': 'sent'
        })


