# -*- coding: utf-8 -*-
from odoo import fields, models


class ProjectProject(models.Model):
    _inherit = "project.project"

    eram_purchase_request_ids = fields.One2many("eram.purchase.req", "project_id")