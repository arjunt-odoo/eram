# -*- coding: utf-8 -*-
from odoo import _, models


class ProjectProject(models.Model):
    _inherit = "project.project"

    def action_view_inwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Inwards"),
            "view_mode": "list,form",
            "domain": [('e_project_id', '=', self.id), ('picking_type_code', '=', 'incoming')],
        }