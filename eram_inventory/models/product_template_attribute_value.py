# -*- coding: utf-8 -*-
from odoo import fields, models


class EramCateg(models.Model):
    _inherit = "product.template.attribute.value"
    _rec_name = 'name'