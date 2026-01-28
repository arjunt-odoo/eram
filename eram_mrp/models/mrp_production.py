# -*- coding: utf-8 -*-
from odoo import api, fields, models


class MrpProduction(models.Model):
    _inherit = 'mrp.production'

    name = fields.Char(readonly=False)
    # e_date = fields.Date(string="Date")
    # e_justification = fields.Char(string="Justification")
    # e_requested_by = fields.Many2one("hr.employee",string="Requested by")
    # e_approved_by = fields.Many2one("hr.employee",string="Approved by")
    # e_requested_by_dept_id = fields.Many2one("hr.department",string="Dept",
    #                                          related="e_requested_by.department_id", store=True)
    # e_approved_by_dept_id = fields.Many2one("hr.department",string="Dept",
    #                                         related="e_approved_by.department_id", store=True)
    e_project_id = fields.Many2one("project.project",string="Project")
    e_task_id = fields.Many2one("project.task",string="Task")
    eram_outward_ids = fields.One2many("eram.outward", "production_id")

    @api.onchange("eram_outward_ids")
    def _onchange_eram_outward_ids(self):
        if not self.product_id:
            return

        move_commands = []

        for outward in self.eram_outward_ids:
            for line in outward.line_ids:
                if not line.product_id:
                    continue

                product_uom = line.uom_id or line.product_id.uom_id

                move_vals = {
                    'name': f"{line.product_id.name} - {line.description or 'From Outward'}",
                    'product_id': line.product_id.id,
                    'product_uom_qty': line.qty,
                    'product_uom': product_uom.id,
                    'e_description_out': line.description,
                    'e_uom_out_id': line.uom_id.id,
                    'e_required_date': line.required_date,
                    'e_remarks': line.remarks
                }

                move_commands.append((0, 0, move_vals))

        self.move_raw_ids = [(5, 0, 0)] + move_commands


    def _get_move_raw_values(self, product, product_uom_qty, product_uom, operation_id=False, bom_line=False):
        data = super()._get_move_raw_values(product, product_uom_qty, product_uom, operation_id, bom_line)
        data["project_id"] = self.e_project_id.id
        data["task_id"] = self.e_task_id.id
        return data

    def _get_move_finished_values(self, product_id, product_uom_qty, product_uom, operation_id=False,
                                  byproduct_id=False, cost_share=0):
        data = super()._get_move_finished_values(product_id, product_uom_qty, product_uom, operation_id, byproduct_id, cost_share)
        data["project_id"] = self.e_project_id.id
        data["task_id"] = self.e_task_id.id
        return data

    def _post_inventory(self, cancel_backorder=False):
        """Override to ensure task_id is propagated during posting"""
        moves_to_do = self.move_raw_ids.filtered(lambda m: m.state not in ('done', 'cancel') and m.picked)
        result = super()._post_inventory(cancel_backorder)
        for move in moves_to_do:
            if move.task_id and move.stock_valuation_layer_ids:
                move.stock_valuation_layer_ids.write({'task_id': move.task_id.id})
        for move in self.move_finished_ids:
            if move.task_id and move.stock_valuation_layer_ids:
                move.stock_valuation_layer_ids.write({'task_id': move.task_id.id})
        return result


class StockMove(models.Model):
    _inherit = 'stock.move'

    e_description_out = fields.Html("Description")
    e_uom_out_id = fields.Many2one("uom.uom", "UOM")
    e_required_date = fields.Date(string="Required Date")

    @api.depends('raw_material_production_id', 'raw_material_production_id.move_raw_ids')
    def _compute_e_si_no(self):
        super()._compute_e_si_no()
        for mo in self.mapped('raw_material_production_id'):
            for index, line in enumerate(mo.move_raw_ids, start=1):
                line.e_si_no = index