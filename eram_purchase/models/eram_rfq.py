# -*- coding: utf-8 -*-
from odoo import _, api, fields, models


class EramPurchaseReq(models.Model):
    _name = "eram.rfq"
    _description = "Eram RFQ"
    _rec_name = "our_ref"

    eram_pr_id = fields.Many2one("eram.purchase.req")
    our_ref = fields.Char("Our REF")
    date = fields.Date()
    our_bid_closing_date = fields.Date("Our Bid Closing Date")
    send_offer_by = fields.Char("Send your Offer by")
    quote_ids = fields.One2many("eram.supplier.quote", "rfq_id")
    no_of_items = fields.Integer("Total Number of Items", compute="_compute_no_of_items", store=True)
    line_ids = fields.One2many("eram.rfq.line", "rfq_id")

    @api.depends('line_ids')
    def _compute_no_of_items(self):
        for rec in self:
            rec.no_of_items = len(rec.line_ids)

    def action_view_pr(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Purchase Request"),
            'res_model': 'eram.purchase.req',
            'res_id': self.eram_pr_id.id,
            'view_mode': 'form'
        }

    def print_rfq(self):
        return self.env.ref('eram_purchase.eram_rfq').report_action(self)


class EramRfqLine(models.Model):
    _name = "eram.rfq.line"
    _description = "Eram RFQ Line"

    sl_no = fields.Integer("S.NO", compute="_compute_sl_no")
    product_id = fields.Many2one("product.product", required=True)
    item_no = fields.Char(required=True)
    description = fields.Char(required=True)
    part_no = fields.Char(required=True)
    uom = fields.Many2one("uom.uom")
    qty = fields.Float("Quantity")
    rfq_id = fields.Many2one("eram.rfq")

    @api.depends('rfq_id', 'rfq_id.line_ids')
    def _compute_sl_no(self):
        for req in self.mapped('rfq_id'):
            for index, line in enumerate(req.line_ids, start=1):
                line.sl_no = index
