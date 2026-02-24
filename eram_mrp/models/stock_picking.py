# -*- coding: utf-8 -*-
from odoo import models, fields
from odoo.tools.float_utils import float_is_zero, float_compare, float_round


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    name = fields.Char(readonly=False)
    e_date = fields.Date(string="Date")
    e_justification = fields.Char(string="Justification")
    e_requested_by = fields.Many2one("hr.employee", string="Requested by")
    e_approved_by = fields.Many2one("hr.employee", string="Approved by")
    e_requested_by_dept_id = fields.Many2one("hr.department", string="Dept",
                                             related="e_requested_by.department_id", store=True)
    e_approved_by_dept_id = fields.Many2one("hr.department", string="Dept",
                                            related="e_approved_by.department_id", store=True)

    def _action_done(self):
        res = super()._action_done()
        self._create_svl_for_internal()
        return res

    def _create_svl_for_internal(self):
        if self.picking_type_id.code != 'internal' or not self.e_transfer_task_id or not self.location_id.task_id or not self.location_dest_id.task_id:
            return

        svl_vals = []

        company = self.company_id
        rounding = self.env['decimal.precision'].precision_get('Product Unit of Measure')

        for move in self.move_ids:
            if float_is_zero(move.quantity, precision_rounding=rounding):
                continue

            from_task = move.location_id.task_id
            to_task   = move.location_dest_id.task_id

            if not (from_task and to_task and from_task != to_task):
                continue

            product = move.product_id

            if product.categ_id.property_cost_method != 'fifo':
                continue

            qty_todo = move.quantity

            source_layers = product.stock_valuation_layer_ids.filtered(
                lambda svl: (
                    svl.company_id == company and
                    svl.task_id == from_task and
                    svl.quantity > 0 and
                    float_compare(svl.remaining_qty, 0.0, precision_rounding=rounding) > 0
                )
            ).sorted(key=lambda svl: (svl.create_date or fields.Datetime.now(), svl.id))

            remaining_to_consume = qty_todo

            for layer in source_layers:
                if float_is_zero(remaining_to_consume, precision_rounding=rounding):
                    break

                available = layer.remaining_qty
                consume_qty = min(remaining_to_consume, available)

                if float_compare(consume_qty, 0.0, precision_rounding=rounding) <= 0:
                    continue

                unit_cost = layer.unit_cost
                out_value = -consume_qty * unit_cost

                svl_vals.append({
                    'product_id': product.id,
                    'company_id': company.id,
                    'quantity': -consume_qty,
                    'value': float_round(out_value, precision_digits=2),
                    'unit_cost': unit_cost,
                    'stock_move_id': move.id,
                    'description': f"Transfer to → {to_task.name}",
                    'task_id': from_task.id,
                    'project_id': from_task.project_id.id,
                })

                remaining_to_consume -= consume_qty

            if float_compare(remaining_to_consume, 0.0, precision_rounding=rounding) > 0:
                fallback_cost = product.with_context(force_company=company.id).standard_price or 0.0
                svl_vals.append({
                    'product_id': product.id,
                    'company_id': company.id,
                    'quantity': -remaining_to_consume,
                    'value': float_round(-remaining_to_consume * fallback_cost, precision_digits=2),
                    'unit_cost': fallback_cost,
                    'stock_move_id': move.id,
                    'description': f"Transfer to → {to_task.name}",
                    'task_id': from_task.id,
                    'project_id': from_task.project_id.id,
                })
                remaining_to_consume = 0.0

            total_in_value = sum(
                -v['value'] for v in svl_vals
                if v['quantity'] < 0 and v.get('stock_move_id') == move.id
            )

            if float_compare(qty_todo, 0.0, precision_rounding=rounding) > 0:
                in_unit_cost = total_in_value / qty_todo if qty_todo != 0 else 0.0

                svl_vals.append({
                    'product_id': product.id,
                    'company_id': company.id,
                    'quantity': qty_todo,
                    'value': float_round(total_in_value, precision_digits=2),
                    'unit_cost': in_unit_cost,
                    'stock_move_id': move.id,
                    'description': f"Transferred from ← {from_task.name}",
                    'task_id': to_task.id,
                    'project_id': to_task.project_id.id,
                })

        if svl_vals:
            self.env['stock.valuation.layer'].sudo().create(svl_vals)