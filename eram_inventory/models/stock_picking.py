# -*- coding: utf-8 -*-
from odoo import api, models, fields


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    active = fields.Boolean("Active", default=True)
    e_grn_id = fields.Many2one("eram.grn", string="Grn. No:")
    e_bill_id = fields.Many2one("account.move", "Invoice No:")
    e_pr_id = fields.Many2one("eram.purchase.req", related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id",
                          store=True, readonly=False)
    e_pr_no = fields.Char("PR. No:", related="e_pr_id.pr_number",
                          store=True)
    e_project_code = fields.Char("Project Code", readonly=False, store=True,
                                 related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id.project_code")
    e_project_id = fields.Many2one("project.project", string="Project", store=True, readonly=False,
                                   related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id.project_id")
    e_task_id = fields.Many2one("project.task", string="Task")
    e_invoice_date = fields.Date("Invoice Date", related="e_bill_id.invoice_date")
    e_invoice_received_date = fields.Date("Invoice Received Date", related="e_bill_id.e_invoice_received_date",
                                          readonly=False)
    currency_id = fields.Many2one("res.currency",
                                  related="purchase_id.currency_id")
    e_total_untaxed = fields.Monetary("Sub Total", compute="_compute_amount", store=True)
    e_amount_total = fields.Monetary("Grand Total", compute="_compute_amount", store=True)
    e_material_inspection_id = fields.Many2one("eram.material.inspection", string="Material Inspection")
    e_additional_charges = fields.Float("Additional Charges")
    department_id = fields.Many2one("hr.department")

    @api.depends('move_ids_without_package.e_total_untaxed',
                 'move_ids_without_package.e_price_total', 'e_additional_charges')
    def _compute_amount(self):
        for rec in self:
            rec.e_total_untaxed = sum(rec.move_ids_without_package.mapped('e_total_untaxed')) + rec.e_additional_charges
            rec.e_amount_total = sum(rec.move_ids_without_package.mapped('e_price_total')) + rec.e_additional_charges

    def _compute_e_bill_ids(self):
        for rec in self:
            rec.e_bill_ids = rec.purchase_id.invoice_ids


