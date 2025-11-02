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
    project_code = fields.Char()
    line_ids = fields.One2many("eram.purchase.req.line", "request_id")
    rfq_ids = fields.One2many("eram.rfq", "eram_pr_id")


    def action_create_rfq(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Order Lines"),
            'res_model': 'eram.purchase.req.line',
            'view_mode': 'list',
            'domain': [('request_id', '=', self.id)],
            'target': 'new'
        }

    def action_view_rfq(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Request For Quotation"),
            'res_model': 'eram.rfq',
            'view_mode': 'list,form',
            'domain': [('eram_pr_id', '=', self.id)]
        }


class EramPurchaseReqLine(models.Model):
    _name = "eram.purchase.req.line"
    _description = "Eram Purchase Request Line"

    request_id = fields.Many2one("eram.purchase.req")
    sl_number = fields.Integer("Sl No", compute="_compute_sl_number", store=True)
    product_id = fields.Many2one("product.product", string="Item")
    description = fields.Html()
    make = fields.Html()
    part_no = fields.Html()
    item_no = fields.Html()
    qty = fields.Float("Quantity")

    @api.depends('request_id', 'request_id.line_ids')
    def _compute_sl_number(self):
        for req in self.mapped('request_id'):
            for index, line in enumerate(req.line_ids, start=1):
                line.sl_number = index

    def create_rfq(self):
        if not self:
            raise ValidationError("Please select at least one line for RFQ!")
        order_lines = []
        for line in self:
            order_lines.append(fields.Command.create({
                'product_id': line.product_id.id,
                'description': line.description,
                'qty': line.qty,
                'part_no': line.part_no,
                'item_no': line.item_no

            }))
        self.request_id.write({
            'rfq_ids': [fields.Command.create({
                'eram_pr_id': self.request_id.id,
                'our_bid_closing_date': self.request_id.closing_date,
                'line_ids': order_lines
            })]
        })
