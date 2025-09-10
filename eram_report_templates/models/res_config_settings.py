# -*- coding: utf-8 -*-
from odoo import fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    alert_user_ids = fields.Many2many("res.users",
                                      relation="eram_alert_user_config_settings_rel",
                                      readonly=False,
                                      related="company_id.alert_user_ids")