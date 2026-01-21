# -*- coding: utf-8 -*-
from odoo import _, api, models


class ProjectTask(models.Model):
    _inherit = "project.task"

    def action_view_inwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Inwards"),
            "view_mode": "list,form",
            "domain": [('e_project_id', '=', self.id), ('picking_type_code', '=', 'incoming')],
        }

    @api.model_create_multi
    def create(self, vals):
        res = super().create(vals)
        for rec in res:
            warehouse = self.env['stock.warehouse'].search([('company_id', '=', self.env.company.id)], limit=1)
            if warehouse:
                parent_location = self.env['stock.location'].search([('warehouse_id', '=', warehouse.id),
                                                                     ('name', '=', warehouse.code)], limit=1)
                if parent_location:
                    location = self.env['stock.location'].create([{
                        'name': rec.name,
                        'location_id': parent_location.id,
                        'warehouse_id': warehouse.id,
                        'usage': 'internal',
                        'task_id': rec.id,
                    }])
                    delivery = self.env['stock.picking.type'].create([{
                        'name': f"{rec.name}: Delivery",
                        'code': 'outgoing',
                        'default_location_src_id': location.id,
                        'sequence_code': f"{rec.name}OUT",
                        'task_id': rec.id
                    }])
                    receipt = self.env['stock.picking.type'].create([{
                        'name': f"{rec.name}: Inwards",
                        'code': 'incoming',
                        'default_location_dest_id': location.id,
                        'sequence_code': f"{rec.name}IN",
                        'return_picking_type_id': delivery.id,
                        'task_id': rec.id
                    }])
                    delivery.write({
                        'return_picking_type_id': receipt.id
                    })
                    manufacture = self.env['stock.picking.type'].create([{
                        'name': f"{rec.name}: Outwards",
                        'code': 'mrp_operation',
                        'default_location_src_id': location.id,
                        'default_location_dest_id': location.id,
                        'sequence_code': f"{rec.name}MO",
                        'task_id': rec.id
                    }])
        return res