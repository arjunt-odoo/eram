from odoo import fields, models


class StockValuationLayer(models.Model):
    _inherit = 'stock.valuation.layer'

    active = fields.Boolean("Active", default=True)
    department_id = fields.Many2one("hr.department", related="stock_move_id.department_id", readonly=False)
    project_id = fields.Many2one("project.project", string="Project Code")
    task_id = fields.Many2one("project.task", string="Task")

