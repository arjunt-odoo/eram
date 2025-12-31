from odoo import fields, models


class StockValuationLayer(models.Model):
    _inherit = 'stock.valuation.layer'

    department_id = fields.Many2one("hr.department", related="stock_move_id.department_id", readonly=False)
    lot_id = fields.Many2one(string="Project Code")
