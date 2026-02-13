
from odoo import models


class PurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    def _prepare_stock_move_vals(self, picking, price_unit, product_uom_qty, product_uom):
        data = super()._prepare_stock_move_vals( picking, price_unit, product_uom_qty, product_uom)
        data.update({
            'e_description': self.e_description,
            'e_part_no': self.e_supplier_quote_line_id.part_no,
            'e_make': self.e_supplier_quote_line_id.make
        })
        return data