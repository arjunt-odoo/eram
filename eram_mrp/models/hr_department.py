# -*- coding: utf-8 -*-
from odoo import _, fields, models


class HrDepartment(models.Model):
    _inherit = 'hr.department'

    eram_outward_req_count = fields.Integer(string="Outward Count", compute="_compute_eram_outward_count")
    eram_outward_app_count = fields.Integer(string="Outward Count", compute="_compute_eram_outward_count")

    def _compute_eram_outward_count(self):
        for rec in self:
            rec.eram_outward_req_count = self.env['mrp.production'].search_count([('e_requested_by_dept_id', '=', rec.id)])
            rec.eram_outward_app_count = self.env['mrp.production'].search_count([('e_approved_by_dept_id', '=', rec.id)])

    def action_view_requested_mo(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_requested_by_dept_id', '=', self.id)],
        }

    def action_view_approved_mo(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_approved_by_dept_id', '=', self.id)],
        }