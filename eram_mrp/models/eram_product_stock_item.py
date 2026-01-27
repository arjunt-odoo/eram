# -*- coding: utf-8 -*-
from odoo import fields, models


class EramProductStockItem(models.TransientModel):
    _name = 'eram.product.stock.item'
    _description = 'Eram Product Stock Item'

    product_id = fields.Many2one("product.product")
    description = fields.Char(string="Description", compute="_compute_description")
    total_received = fields.Float(string="Total Received")
    total_consumed = fields.Float(string="Total Consumed")
    balance = fields.Float(string="Balance")

    def _compute_description(self):
        for rec in self:
            if rec.product_id:
                description = rec.product_id.name

                if rec.product_id.product_template_attribute_value_ids:
                    attribute_values = []
                    for value in rec.product_id.product_template_attribute_value_ids:
                        attr_str = f"{value.name} {value.attribute_id.name}"
                        attribute_values.append(attr_str)
                    if attribute_values:
                        description = f"{description} {' '.join(attribute_values)}"

                rec.description = description
            else:
                rec.description = False