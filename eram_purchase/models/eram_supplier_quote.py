# -*- coding: utf-8 -*-
from odoo import _, api, fields, models


class EramSupplierQuote(models.Model):
    _name = "eram.supplier.quote"
    _description = "Eram Supplier Quote"
    _rec_name = "quote_no"

    quote_no = fields.Char("Quote No")
    quote_date = fields.Date("Quote Date")
    partner_id = fields.Many2one("res.partner", string="Supplier Details")
    line_ids = fields.One2many("eram.supplier.quote.line", "quote_id")
    company_id = fields.Many2one("res.company", default=lambda self: self.env.company)
    rfq_id = fields.Many2one("eram.rfq")
    purchase_id = fields.Many2one("purchase.order")

    def action_view_rfq(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Request For Quotation"),
            'res_model': 'eram.rfq',
            'res_id': self.rfq_id.id,
            'view_mode': 'form'
        }

    def action_view_po(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Purchase Order"),
            'res_model': 'purchase.order',
            'res_id': self.purchase_id.id,
            'view_mode': 'form'
        }

    def open_record(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _("Supplier Quote"),
            'res_model': 'eram.supplier.quote',
            'res_id': self.id,
            'view_mode': 'form'
        }

    def action_create_po(self):
        order_lines = []
        for line in self.line_ids:
            order_lines.append(fields.Command.create({
                'product_id': line.product_id.id,
                'e_description': line.description,
                'product_qty': line.qty,
                'price_unit': line.price_unit,
            }))
        self.purchase_id = self.purchase_id.create({
            'partner_id': self.partner_id.id,
            'order_line': order_lines
        })


class EramSupplierQuoteLine(models.Model):
    _name = "eram.supplier.quote.line"
    _description = "Eram Supplier Quote Line"

    quote_id = fields.Many2one("eram.supplier.quote")
    sl_no = fields.Integer("Sl No", compute="_compute_sl_no")
    product_id = fields.Many2one("product.product", "Item")
    description = fields.Char()
    qty = fields.Float()
    company_id = fields.Many2one("res.company", related="quote_id.company_id")
    currency_id = fields.Many2one("res.currency",  default=lambda self: self.env.company.currency_id)
    price_unit = fields.Monetary("Unit Price")
    total_untaxed = fields.Monetary("Total", compute="_compute_total_untaxed", store=True)
    tax_ids = fields.Many2many("account.tax", string="Tax")
    total_amount = fields.Monetary(string="Sub Total", compute="_compute_total_amount" , store=True)

    @api.depends('quote_id', 'quote_id.line_ids')
    def _compute_sl_no(self):
        for quote in self.mapped('quote_id'):
            for index, line in enumerate(quote.line_ids, start=1):
                line.sl_no = index

    @api.depends('currency_id', 'tax_ids', 'price_unit', 'qty', 'product_id')
    def _compute_total_untaxed(self):
        for rec in self:
            rec.total_untaxed = rec.qty * rec.price_unit

    @api.depends('currency_id', 'tax_ids', 'price_unit', 'total_untaxed', 'qty', 'product_id')
    def _compute_total_amount(self):
        for rec in self:
            if rec.tax_ids:
                taxes = rec.tax_ids.compute_all(rec.total_untaxed)
                rec.total_amount = taxes['total_included']
            else:
                rec.total_amount = rec.total_untaxed