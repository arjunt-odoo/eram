# -*- coding: utf-8 -*-
from odoo import _, api, fields, models


class StockLocation(models.Model):
    _inherit = 'stock.location'

    task_id = fields.Many2one('project.task')