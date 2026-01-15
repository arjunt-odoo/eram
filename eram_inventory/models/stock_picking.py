# -*- coding: utf-8 -*-
from odoo import api, models, fields


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    e_grn_id = fields.Many2one("eram.grn", string="Grn. No:")
    e_bill_id = fields.Many2one("account.move", "Invoice No:")
    e_pr_no = fields.Char("PR. No:", related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id.pr_number",
                          store=True, readonly=False)
    e_project_code = fields.Char("Project Code", readonly=False, store=True,
                                 related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id.project_code")
    e_project_id = fields.Many2one("project.project", string="Project", store=True, readonly=False,
                                   related="purchase_id.e_supplier_quote_id.rfq_id.eram_pr_id.project_id")
    e_invoice_date = fields.Date("Invoice Date", related="e_bill_id.invoice_date")
    e_invoice_received_date = fields.Date("Invoice Received Date", related="e_bill_id.e_invoice_received_date",
                                          readonly=False)
    currency_id = fields.Many2one("res.currency",
                                  related="purchase_id.currency_id")
    e_total_untaxed = fields.Monetary("Sub Total", compute="_compute_amount", store=True)
    e_amount_total = fields.Monetary("Grand Total", compute="_compute_amount", store=True)
    e_material_inspection_id = fields.Many2one("eram.material.inspection", string="Material Inspection")

    @api.depends('move_ids_without_package.e_total_untaxed',
                 'move_ids_without_package.e_price_total')
    def _compute_amount(self):
        for rec in self:
            rec.e_total_untaxed = sum(rec.move_ids_without_package.mapped('e_total_untaxed'))
            rec.e_amount_total = sum(rec.move_ids_without_package.mapped('e_price_total'))

    def _compute_e_bill_ids(self):
        for rec in self:
            rec.e_bill_ids = rec.purchase_id.invoice_ids


class StockMove(models.Model):
    _inherit = 'stock.move'

    currency_id = fields.Many2one("res.currency",
                                  related="purchase_line_id.currency_id")
    e_grn_id = fields.Many2one("eram.grn",
                               related="picking_id.e_grn_id", store=True)
    e_si_no = fields.Integer("SI. NO", default=1, store=True,
                             compute="_compute_e_si_no")
    e_item_code = fields.Html("ITEM CODE", related="purchase_line_id.e_supplier_quote_line_id.item_code", readonly=False)
    e_description = fields.Html("DESCRIPTION", related="purchase_line_id.e_description", readonly=False)
    e_part_no = fields.Html("PART NO.", related="purchase_line_id.e_supplier_quote_line_id.part_no", readonly=False)
    e_make = fields.Html("MAKE", related="purchase_line_id.e_supplier_quote_line_id.make", readonly=False)
    e_uom_id = fields.Many2one("uom.uom", "UOM", related="purchase_line_id.product_uom",
                               store=True, readonly=False)
    e_price_unit = fields.Float("UNIT PRICE", related="purchase_line_id.price_unit")
    e_total_untaxed = fields.Monetary(compute="_compute_amount", store=True)
    e_tax_ids = fields.Many2many("account.tax", string="TAX", related="purchase_line_id.taxes_id")
    e_price_total = fields.Monetary("TOTAL PRICE",store=True,
                                    compute="_compute_amount")
    e_discrepancy = fields.Html("DISCREPANCY DETAILS")
    e_receival_status = fields.Many2one("eram.receival.status", "Received Status")
    e_qty_accepted = fields.Float("Quantity Accepted")
    e_qty_rejected = fields.Float("Quantity Received")
    e_remarks = fields.Html("Remarks")
    e_visual_inspection = fields.Boolean(string="Visual Inspection")
    e_dimensional_inspection = fields.Boolean(string="Dimensional Inspection")
    e_functional_inspection = fields.Boolean(string="Functional Inspection")
    e_rejection_reason = fields.Char(string="Reason for Rejection")

    @api.depends('picking_id', 'picking_id.move_ids')
    def _compute_e_si_no(self):
        for picking in self.mapped('picking_id'):
            for index, line in enumerate(picking.move_ids, start=1):
                line.e_si_no = index

    @api.depends('product_id', 'product_id.lst_price', 'e_price_unit',
                 'e_total_untaxed', 'e_tax_ids', 'quantity', 'product_uom_qty')
    def _compute_amount(self):
        for rec in self:
            rec.e_total_untaxed = rec.quantity * rec.e_price_unit
            if rec.e_tax_ids:
                taxes = rec.e_tax_ids.compute_all(rec.e_total_untaxed)
                rec.e_price_total = taxes["total_included"]
            else:
                rec.e_price_total = rec.e_total_untaxed
