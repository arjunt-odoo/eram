# -*- coding: utf-8 -*-
from odoo import fields, models

class ResCompany(models.Model):
    _inherit = "res.company"

    alert_user_ids = fields.Many2many("res.users",
                                      relation="eram_alert_user_company_rel")
    footer_image = fields.Image(
        string='Footer Image',
        help='Upload the image to use as dynamic footer in reports. It can be updated over time.')


class BaseDocumentLayout(models.TransientModel):
    _inherit = "base.document.layout"

    footer_image = fields.Image(related='company_id.footer_image')