# -*- coding: utf-8 -*-
from odoo import _, models, fields

class ProductProduct(models.Model):
    _inherit = 'product.product'

    e_categ_id = fields.Many2one("eram.categ")

    def action_view_invoice_lines(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "account.move.line",
            "name": _("Invoice Lines"),
            "view_mode": "list,form",
            "views": [(self.env.ref("eram_inventory.view_account_move_line_list_service_product").id, 'list'),
                      (False, 'form')],
            "domain": [('product_id', 'in', self.ids)],
        }


