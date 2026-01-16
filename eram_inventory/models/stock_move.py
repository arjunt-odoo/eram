# -*- coding: utf-8 -*-
from odoo import api, fields, models


class StockMove(models.Model):
    _inherit = 'stock.move'

    active = fields.Boolean("Active", default=True)
    currency_id = fields.Many2one("res.currency",
                                  related="purchase_line_id.currency_id")
    department_id = fields.Many2one("hr.department")
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

    def _get_in_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_in_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            move = self.filtered(lambda m: m.id == vals.get('stock_move_id', False))
            if move:
                vals["project_id"] = move.move_line_ids[0].project_id.id if move.move_line_ids else False
        return svl_vals_list

    def _get_out_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_out_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            move = self.filtered(lambda m: m.id == vals.get('stock_move_id', False))
            if move:
                vals["project_id"] = move.move_line_ids[0].project_id.id if move.move_line_ids else False
        return svl_vals_list

    def _get_dropshipped_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_dropshipped_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            move = self.filtered(lambda m: m.id == vals.get('stock_move_id', False))
            if move:
                vals["project_id"] = move.move_line_ids[0].project_id.id if move.move_line_ids else False
        return svl_vals_list

class StockMoveLine(models.Model):
    _inherit = 'stock.move.line'

    active = fields.Boolean("Active", default=True)
    project_id = fields.Many2one("project.project")


class StockQuant(models.Model):
    _inherit = 'stock.quant'

    active = fields.Boolean("Active", default=True)


class StockLot(models.Model):
    _inherit = 'stock.lot'

    active = fields.Boolean("Active", default=True)




