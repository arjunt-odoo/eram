# -*- coding: utf-8 -*-
from odoo import api, fields, models


class MrpProduction(models.Model):
    _inherit = 'mrp.production'

    name = fields.Char(readonly=False)
    e_date = fields.Date(string="Date")
    e_justification = fields.Char(string="Justification")
    e_requested_by = fields.Many2one("hr.employee",string="Requested by")
    e_approved_by = fields.Many2one("hr.employee",string="Approved by")
    e_requested_by_dept_id = fields.Many2one("hr.department",string="Dept",
                                             related="e_requested_by.department_id", store=True)
    e_approved_by_dept_id = fields.Many2one("hr.department",string="Dept",
                                            related="e_approved_by.department_id", store=True)
    e_project_id = fields.Many2one("project.project",string="Project")
    e_task_id = fields.Many2one("project.task",string="Task")

    @api.model_create_multi
    def create(self, vals_list):
        today = fields.Date.today()
        current_month = today.month

        if current_month >= 4:
            current_fy_year = today.year
            next_fy_year = today.year + 1
        else:
            current_fy_year = today.year - 1
            next_fy_year = today.year

        for val in vals_list:
            batch_size = 50
            seq = self.env['ir.sequence'].sudo()
            position_str = seq.next_by_code('mrp.production.position')
            position = int(position_str)
            batch_seq = seq.search([('code', '=', 'mrp.production.batch')], limit=1)
            batch = batch_seq.number_next
            if position > batch_size:
                seq.next_by_code('mrp.production.batch')
                position_seq = seq.search([('code', '=', 'mrp.production.position')], limit=1)
                position_seq.write({'number_next': 2})
                position = 1
                batch += 1
            middle = f"{batch:03d}"
            last = f"{position:04d}"
            val['name'] = f"BLR-MWF-{current_fy_year}-{next_fy_year}-R0-{middle}-{last}"

        return super(MrpProduction, self).create(vals_list)

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
    e_part_no_out = fields.Html(string="Part No")
    e_uom_out_id = fields.Many2one("uom.uom", "UOM")
    e_required_date = fields.Date(string="Required Date")

    @api.depends('raw_material_production_id', 'raw_material_production_id.move_raw_ids')
    def _compute_e_si_no(self):
        super()._compute_e_si_no()
        for mo in self.mapped('raw_material_production_id'):
            for index, line in enumerate(mo.move_raw_ids, start=1):
                line.e_si_no = index