# -*- coding: utf-8 -*-
from odoo import _, api, fields, models
from odoo.exceptions import ValidationError


class EramPurchaseReq(models.Model):
    _name = "eram.purchase.req"
    _description = "Eram Purchase Request"
    _rec_name = "pr_number"

    pr_number = fields.Char("PR No")
    pr_date = fields.Date("PR Date")
    closing_date = fields.Date()
    line_ids = fields.One2many("eram.purchase.req.line", "request_id")
    rfq_id = fields.Many2one("eram.rfq")


    def action_create_rfq(self):
        if not self.line_ids:
            raise ValidationError("Please add at least one product in PR!")
        order_lines = []
        for line in self.line_ids:
            order_lines.append(fields.Command.create({
                'product_id': line.product_id.id,
                'description': line.description,
                'qty': line.qty,
                'part_no': line.part_no,
                'item_no': line.item_no

            }))
        self.rfq_id =  self.rfq_id.create({
            'eram_pr_id': self.id,
            'our_bid_closing_date': self.closing_date,
            'line_ids': order_lines
        })

    def action_view_rfq(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Request For Quotation"),
            'res_model': 'eram.rfq',
            'res_id': self.rfq_id.id,
            'view_mode': 'form'
        }


class EramPurchaseReqLine(models.Model):
    _name = "eram.purchase.req.line"
    _description = "Eram Purchase Request Line"

    request_id = fields.Many2one("eram.purchase.req")
    sl_number = fields.Integer("Sl No", compute="_compute_sl_number", store=True)
    product_id = fields.Many2one("product.product", string="Item")
    description = fields.Char()
    part_no = fields.Char()
    item_no = fields.Char()
    qty = fields.Float("Quantity")

    @api.depends('request_id', 'request_id.line_ids')
    def _compute_sl_number(self):
        for req in self.mapped('request_id'):
            for index, line in enumerate(req.line_ids, start=1):
                line.sl_number = index
