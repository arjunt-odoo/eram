# -*- coding: utf-8 -*-
from odoo import _, api, fields, models

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    partner_id = fields.Many2one("res.partner", required=False)
    state = fields.Selection(selection_add=[('draft', "Draft"),
                                            ('sent', "Sent")])
    e_supplier_quote_id = fields.Many2one("eram.supplier.quote", string="Supplier Quote")
    e_bill_ids = fields.Many2many("account.move", "eram_po_bill_rel", string="Bills")
    e_project_id = fields.Many2one("project.project", string="Project", readonly=False, store=True,
                                   related="e_supplier_quote_id.rfq_id.eram_pr_id.project_id")
    task_id = fields.Many2one("project.task", string="Task", readonly=False, store=True,
                                   related="e_supplier_quote_id.rfq_id.eram_pr_id.task_id")

    def action_set_as_sent(self):
        self.write({
            'state': 'sent'
        })

class PurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    e_supplier_quote_line_id = fields.Many2one("eram.supplier.quote.line")




