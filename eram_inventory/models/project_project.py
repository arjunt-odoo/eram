# -*- coding: utf-8 -*-
from odoo import _, api, models


class ProjectProject(models.Model):
    _inherit = "project.project"

    def action_view_inwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Inwards"),
            "view_mode": "list,form",
            "domain": [('e_project_id', '=', self.id), ('picking_type_code', '=', 'incoming')],
        }


class ProjectTask(models.Model):
    _inherit = "project.task"

    def action_view_inwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Inwards"),
            "view_mode": "list,form",
            "domain": [('e_task_id', '=', self.id), ('picking_type_code', '=', 'incoming')],
        }

    def action_view_inward_items(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.move",
            "name": _("Inward Items"),
            "view_mode": "list,form",
            "views": [[self.env.ref("eram_inventory.view_stock_move_individual_eram").id, "list"], [False, "form"]],
            "domain": [('picking_id.e_task_id', '=', self.id), ('picking_id.picking_type_code', '=', 'incoming')],
        }

    def action_view_purchase_order(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "purchase.order",
            "name": _("Purchase Order"),
            "view_mode": "list,form",
            "domain": [('task_id', '=', self.id)],
        }

    def action_view_locations(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.location",
            "name": _("Locations"),
            "view_mode": "list,form",
            "domain": [('task_id', '=', self.id)],
        }

    def action_view_valuation(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.valuation.layer",
            "name": _("Valuation"),
            "view_mode": "list",
            "domain": [('task_id', '=', self.id)],
        }

    def action_view_purchase_request(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "eram.purchase.req",
            "name": _("Purchase Request"),
            "view_mode": "list",
            "domain": [('task_id', '=', self.id)],
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
                        'name': f"{rec.name}: Inward",
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
                        'name': f"{rec.name}: Outward",
                        'code': 'mrp_operation',
                        'default_location_src_id': location.id,
                        'default_location_dest_id': location.id,
                        'sequence_code': f"{rec.name}MO",
                        'task_id': rec.id
                    }])
        return res