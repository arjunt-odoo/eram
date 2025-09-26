# -*- coding: utf-8 -*-
from odoo import api, fields, models


class EramCustomerPo(models.Model):
    _name = "eram.customer.po"

    name = fields.Char(string="PO No")
    date = fields.Date(string="PO Date")
    currency_id = fields.Many2one("res.currency",  default=lambda self: self.env.company.currency_id)
    amount = fields.Monetary(string="PO Value")
    partner_id = fields.Many2one("res.partner", string="Customer",
                                 related="sale_id.partner_id")
    company_id = fields.Many2one("res.company",  default=lambda self: self.env.company)
    tax_ids = fields.Many2many("account.tax", string="Taxes")
    amount_total = fields.Monetary(string="Total amount",
                                   compute="_compute_amount_total",
                                   store=True)
    sale_id = fields.Many2one("sale.order")
    line_ids = fields.One2many("eram.customer.po.line", "customer_po_id")

    @api.depends('amount', 'currency_id', 'tax_ids')
    def _compute_amount_total(self):
        for rec in self:
            if rec.tax_ids:
                taxes = rec.tax_ids.compute_all(rec.amount)
                rec.amount_total = taxes['total_included']
            else:
                rec.amount_total = rec.amount


class EramCustomerPoLine(models.Model):
    _name = 'eram.customer.po.line'

    product_id = fields.Many2one("product.product")
    quantity = fields.Float()
    customer_po_id = fields.Many2one("eram.customer.po")
