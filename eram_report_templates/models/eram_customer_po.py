# -*- coding: utf-8 -*-
from odoo import fields, models


class EramCustomerPo(models.Model):
    _name = "eram.customer.po"

    name = fields.Char(string="PO No")
    date = fields.Date(string="PO Date")
    currency_id = fields.Many2one("res.currency",  default=lambda self: self.env.company.currency_id)
    amount = fields.Monetary(string="PO Value")
    partner_id = fields.Many2one("res.partner", string="Customer")
    company_id = fields.Many2one("res.company",  default=lambda self: self.env.company)