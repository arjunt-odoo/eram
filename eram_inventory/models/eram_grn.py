# -*- coding: utf-8 -*-
from odoo import fields, models


class EramGrn(models.Model):
    _name = "eram.grn"
    _description = "Eram GRN"

    name = fields.Char()
    partner_id = fields.Many2one("res.partner", related="picking_id.partner_id", readonly=False)
    project_id = fields.Many2one("project.project")
    date_received = fields.Date()
    pr_id = fields.Many2one("eram.purchase.req", related="picking_id.e_pr_id", readonly=False)
    purchase_id = fields.Many2one("purchase.order", related="picking_id.purchase_id")
    po_no = fields.Char("PO. No:", related="picking_id.e_po_no", readonly=False)
    bill_id = fields.Many2one("account.move", related="picking_id.e_bill_id")
    picking_id = fields.Many2one("stock.picking")
    line_ids = fields.One2many("stock.move", "e_grn_id",
                               related="picking_id.move_ids", readonly=False)

    def print_grn(self):
        return self.env.ref('eram_inventory.eram_grn').report_action(self)
