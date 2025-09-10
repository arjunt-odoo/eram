# -*- coding: utf-8 -*-
from odoo import fields, models

class ResCompany(models.Model):
    _inherit = "res.company"

    alert_user_ids = fields.Many2many("res.users",
                                      relation="eram_alert_user_company_rel")