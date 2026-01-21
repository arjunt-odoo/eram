# -*- coding: utf-8 -*-
from odoo import api, models, fields


class StockPickingType(models.Model):
    _inherit = 'stock.picking.type'

    task_id = fields.Many2one('project.task')