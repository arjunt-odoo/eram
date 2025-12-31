# -*- coding: utf-8 -*-
from odoo import fields, models


class EramCateg(models.Model):
    _name = "eram.categ"
    _description = "Eram Categ"

    name = fields.Char()
    parent_id = fields.Many2one("eram.categ", "Parent")
    child_ids = fields.One2many("eram.categ", "parent_id")
    product_tmpl_id = fields.Many2one("product.template")

    def get_hierarchy_path(self):  # Recursive path builder
        """Method to get substitute name."""
        path = self.name
        current = self
        while current.parent_id:
            path = f"{current.parent_id.name} / {path}"
            current = current.parent_id
        return path