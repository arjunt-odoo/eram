# -*- coding: utf-8 -*-
from odoo import fields, models


class EramMaterialInspection(models.Model):
    _name = "eram.material.inspection"

    name = fields.Char("Doc No")
    inspected_date = fields.Date("Inspected Date")
    partner_id = fields.Many2one("res.partner", related="picking_id.partner_id", readonly=False,
                                 string="Supplier Name")
    project_code = fields.Char(related="picking_id.e_project_code", readonly=False)
    date_received = fields.Date(string="Date of item receival in inventory")
    pr_no = fields.Char(related="picking_id.e_pr_no", readonly=False)
    purchase_id = fields.Many2one("purchase.order", related="picking_id.purchase_id", readonly=False,
                                  string="P.O. No.")
    bill_id = fields.Many2one("account.move", related="picking_id.e_bill_id")
    picking_id = fields.Many2one("stock.picking")
    line_ids = fields.Many2many("stock.move", compute="_compute_line_ids", readonly=False)

    def _compute_line_ids(self):
        for rec in self:
            rec.line_ids = rec.picking_id.move_ids.filtered(lambda m: m.product_id.product_tmpl_id.e_allow_inspection)

    def print_inspection_report(self):
        return self.env.ref('eram_inventory.eram_material_inspection').report_action(self)
