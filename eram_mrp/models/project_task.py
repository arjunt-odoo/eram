# -*- coding: utf-8 -*-
from odoo import _, models
from odoo.exceptions import ValidationError


class ProjectTask(models.Model):
    _inherit = "project.task"

    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_task_id', '=', self.id)],
        }

    def action_view_stock_moves(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.move.line",
            "name": _("Moves"),
            "view_mode": "list,form",
            "domain": [
                '|', '|', '|', ('picking_id.e_task_id', '=', self.id), ('production_id.e_task_id', '=', self.id),
                ('move_id.raw_material_production_id.e_task_id', '=', self.id),
                ('move_id.production_id.e_task_id', '=', self.id)
            ],
        }

    def action_transfer_left_overs(self):
        self.ensure_one()
        source_location = self.env["stock.location"].search([('task_id', '=', self.id)], limit=1)
        stock_quant = self.env['stock.quant'].search([('location_id', '=', source_location.id)])
        if not stock_quant:
            raise ValidationError(_("Nothing left on project to move!"))
        internal_transfer_type = self.env["stock.picking.type"].search([('code', '=', 'internal'),
                                                                        ('company_id', 'in', self.env.company.ids)], limit=1)
        move_vals = []
        for quant in stock_quant:
            move_vals.append({
                "product_id": quant.product_id.id,
                "product_uom_qty": quant.available_quantity,
            })
        if not move_vals:
            raise ValidationError(_("Nothing left on project to move!"))
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Transfer"),
            "view_mode": "form",
            "target": "new",
            "context": {
                'default_picking_type_id': internal_transfer_type.id,
                'default_location_id': source_location.id,
                'default_move_ids_without_package': [(0, 0, vals) for vals in move_vals]
            }
        }