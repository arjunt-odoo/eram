# -*- coding: utf-8 -*-
from odoo import api, fields, models


class MrpProduction(models.Model):
    _inherit = 'mrp.production'

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
    def create(self, vals):
        for val in vals:
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
            val['name'] = f"R0-{middle}-{last}"

        return super(MrpProduction, self).create(vals)


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