# -*- coding: utf-8 -*-
from odoo import _, api, models


class ProjectTask(models.Model):
    _inherit = "project.task"

    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_task_id', '=', self.id)],
        }