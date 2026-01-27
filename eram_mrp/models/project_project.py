# -*- coding: utf-8 -*-
from odoo import _, fields, models
from odoo.exceptions import ValidationError


class ProjectProject(models.Model):
    _inherit = "project.project"

    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_project_id', '=', self.id)],
        }


class ProjectTask(models.Model):
    _inherit = "project.task"

    e_product_stock_item_ids = fields.Many2many("eram.product.stock.item", string="Stock Items",
                                                compute="_compute_e_product_stock_item_ids")

    def _compute_e_product_stock_item_ids(self):
        for task in self:
            move_lines = self.env["stock.move.line"].search([
                '|', '|', '|',
                ('picking_id.e_task_id', '=', task.id),
                ('production_id.e_task_id', '=', task.id),
                ('move_id.raw_material_production_id.e_task_id', '=', task.id),
                ('move_id.production_id.e_task_id', '=', task.id)
            ])

            product_data = {}

            for move_line in move_lines:
                product = move_line.product_id
                if product.id not in product_data:
                    product_data[product.id] = {
                        'product_id': product.id,
                        'total_received': 0,
                        'total_consumed': 0,
                        'balance': 0,
                    }

                if move_line.location_usage not in ('internal','transit') and move_line.location_dest_usage in ('internal','transit'):
                    product_data[product.id]['total_received'] += move_line.quantity
                elif move_line.location_usage in ('internal','transit') and move_line.location_dest_usage not in ('internal','transit'):
                    product_data[product.id]['total_consumed'] += move_line.quantity

            for product_id in product_data:
                data = product_data[product_id]
                data['balance'] = data['total_received'] - data['total_consumed']

            stock_item_ids = []
            for product_id, data in product_data.items():
                stock_item = self.env['eram.product.stock.item'].search([
                    ('product_id', '=', product_id),
                ], limit=1)

                if not stock_item:
                    stock_item = self.env['eram.product.stock.item'].create({
                        'product_id': product_id,
                        'total_received': data['total_received'],
                        'total_consumed': data['total_consumed'],
                        'balance': data['balance'],
                    })
                else:
                    stock_item.write({
                        'total_received': data['total_received'],
                        'total_consumed': data['total_consumed'],
                        'balance': data['balance'],
                    })

                stock_item_ids.append(stock_item.id)

            task.e_product_stock_item_ids = [(6, 0, stock_item_ids)]


    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('e_task_id', '=', self.id)],
        }

    def action_view_outward_items(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.move",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "domain": [('raw_material_production_id.e_task_id', '=', self.id)],
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

    def action_view_internal_transfers(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Transfers"),
            "view_mode": "list,form",
            "views": [[False, "list"], [self.env.ref("eram_mrp.view_picking_form_left_over").id, "form"]],
            "domain": [('e_transfer_task_id', '=', self.id)],
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
                'name': quant.product_id.display_name or '/',
                'product_uom': quant.product_id.uom_id.id,
                'product_id': quant.product_id.id,
                'product_uom_qty': quant.quantity,
            })
        if not move_vals:
            raise ValidationError(_("Nothing left on project to move!"))
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Transfer"),
            "view_mode": "form",
            "target": "new",
            "view_id": self.env.ref("eram_mrp.view_picking_form_left_over").id,
            "context": {
                'default_picking_type_id': internal_transfer_type.id,
                'default_location_id': source_location.id,
                'default_move_ids_without_package': [(0, 0, vals) for vals in move_vals],
                'default_e_transfer_task_id': self.id,
            }
        }