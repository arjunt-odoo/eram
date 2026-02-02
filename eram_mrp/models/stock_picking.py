# -*- coding: utf-8 -*-
from odoo import api, models, fields


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    e_date = fields.Date(string="Date")
    e_justification = fields.Char(string="Justification")
    e_requested_by = fields.Many2one("hr.employee", string="Requested by")
    e_approved_by = fields.Many2one("hr.employee", string="Approved by")
    e_requested_by_dept_id = fields.Many2one("hr.department", string="Dept",
                                             related="e_requested_by.department_id", store=True)
    e_approved_by_dept_id = fields.Many2one("hr.department", string="Dept",
                                            related="e_approved_by.department_id", store=True)