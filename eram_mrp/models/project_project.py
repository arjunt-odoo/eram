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
            task_location = self.env["stock.location"].search(
                [("task_id", "=", task.id)], limit=1
            )
            if not task_location:
                task.e_product_stock_item_ids = False
                continue

            quants = self.env["stock.quant"].search([
                ("location_id", "=", task_location.id),
            ])

            product_quant_map = {q.product_id.id: q.quantity for q in quants}

            if not product_quant_map:
                task.e_product_stock_item_ids = False
                continue

            incoming_lines = self.env["stock.move.line"].search([
                ("location_dest_id", "=", task_location.id)
            ])

            product_data = {}

            for line in incoming_lines:
                product_id = line.product_id.id
                if product_id not in product_data:
                    product_data[product_id] = {
                        "product_id": product_id,
                        "total_received": 0.0,
                        "balance": 0.0,
                        "total_consumed": 0.0,
                    }
                product_data[product_id]["total_received"] += line.quantity

            stock_item_ids = []

            for product_id, data in product_data.items():
                current_balance = product_quant_map.get(product_id, 0.0)
                data["balance"] = current_balance

                data["total_consumed"] = data["total_received"] - current_balance

                if data["total_consumed"] < 0:
                    data["total_consumed"] = 0.0

                stock_item = self.env["eram.product.stock.item"].search([
                    ("product_id", "=", product_id),
                ], limit=1)

                vals = {
                    "product_id": product_id,
                    "total_received": data["total_received"],
                    "total_consumed": data["total_consumed"],
                    "balance": data["balance"],
                }

                if not stock_item:
                    stock_item = self.env["eram.product.stock.item"].create(vals)
                else:
                    stock_item.write(vals)

                stock_item_ids.append(stock_item.id)

            task.e_product_stock_item_ids = [(6, 0, stock_item_ids)] if stock_item_ids else False


    def action_view_production(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "mrp.production",
            "name": _("Manufacture"),
            "view_mode": "list,form",
            "domain": [('e_task_id', '=', self.id)],
        }

    def action_view_outwards(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.picking",
            "name": _("Outwards"),
            "view_mode": "list,form",
            "views": [[self.env.ref("eram_mrp.view_picking_list_outward").id, "list"],
                      [self.env.ref("eram_mrp.view_picking_form_outward").id, "form"]],
            "domain": [('e_task_id', '=', self.id),
                       ('picking_type_code', '=', 'internal'),
                       ('picking_type_id.task_id', '=', self.id)],
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

    # def action_view_stock_moves(self):
    #     self.ensure_one()
    #     return {
    #         "type": "ir.actions.act_window",
    #         "res_model": "stock.move.line",
    #         "name": _("Moves"),
    #         "view_mode": "list,form",
    #         "domain": [
    #             '|', '|', '|', ('picking_id.e_task_id', '=', self.id), ('production_id.e_task_id', '=', self.id),
    #             ('move_id.raw_material_production_id.e_task_id', '=', self.id),
    #             ('move_id.production_id.e_task_id', '=', self.id)
    #         ],
    #     }

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
            if quant.quantity < 1:
                continue
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