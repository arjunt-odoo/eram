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
                                 related="stock_move_id.project_id")
    task_id = fields.Many2one("project.task", string="Task", related="stock_move_id.task_id", store=True)