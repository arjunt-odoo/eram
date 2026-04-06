# -*- coding: utf-8 -*-
from odoo import _, api, fields, models


class StockLocation(models.Model):
    _inherit = 'stock.location'

    task_id = fields.Many2one('project.task')


class StockValuationLayer(models.Model):
    _inherit = 'stock.valuation.layer'

    active = fields.Boolean("Active", default=True)
    department_id = fields.Many2one("hr.department", related="stock_move_id.department_id", readonly=False)
    project_id = fields.Many2one("project.project", string="Project Code", store=True,
                                 related="stock_move_id.project_id", readonly=False)
    task_id = fields.Many2one("project.task", string="Task", related="stock_move_id.task_id", store=True,
                              readonly=False)
    tax_ids = fields.Many2many("account.tax")
    total_taxed = fields.Monetary("Total Taxed", compute='_compute_total_taxed')
    qty_moved = fields.Float("Qty Moved")
    is_empty = fields.Boolean("Is Empty", compute="_compute_is_empty", store=True)
    eram_ow_move_id = fields.Many2one("stock.move")

    @api.depends('qty_moved', 'quantity')
    def _compute_is_empty(self):
        for rec in self:
            if rec.qty_moved == rec.quantity:
                rec.is_empty = True

    def _compute_total_taxed(self):
        for rec in self:
            if rec.tax_ids:
                taxes = rec.tax_ids.compute_all(rec.value)
                rec.total_taxed = taxes["total_included"]
            else:
                rec.total_taxed = rec.value