# -*- coding: utf-8 -*-
from odoo import  models


class ProductTemplateAttributeValue(models.Model):
    _inherit = "product.template.attribute.value"

    def _compute_display_name(self):
        for value in self:
            value.display_name = value.name