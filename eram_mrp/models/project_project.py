# -*- coding: utf-8 -*-
from odoo import _, models


class ProjectProject(models.Model):
    _inherit = "project.project"

    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('project_id', '=', self.id)],
        }