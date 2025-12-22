# -*- coding: utf-8 -*-
from odoo import fields, models


class EramGrn(models.Model):
    _name = "eram.grn"

    name = fields.Char()
    partner_id = fields.Many2one("res.partner", related="picking_id.partner_id", readonly=False)
    project_code = fields.Char(related="picking_id.e_project_code", readonly=False)
    date_received = fields.Date()
    pr_no = fields.Char(related="picking_id.e_pr_no", readonly=False)
    purchase_id = fields.Many2one("purchase.order", related="picking_id.purchase_id")
    bill_id = fields.Many2one("account.move", related="picking_id.e_bill_id")
    picking_id = fields.Many2one("stock.picking")
    line_ids = fields.One2many("stock.move", "e_grn_id",
                               related="picking_id.move_ids", readonly=False)

    def print_grn(self):
        return self.env.ref('eram_inventory.eram_grn').report_action(self)
