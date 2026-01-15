# -*- coding: utf-8 -*-
from odoo import fields, models


class StockMove(models.Model):
    _inherit = 'stock.move'

    department_id = fields.Many2one("hr.department")

    def _get_in_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_in_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            lines = self._get_in_move_lines()
            if forced_quantity:
                lot = forced_quantity[0]
                line = lines.filtered(lambda l: l.lot_id == lot)
            else:
                lot = self.env['stock.lot'].browse(vals.get('lot_id'))
                line = lines.filtered(lambda l: l.lot_id == lot)
            if line:
                vals['project_id'] = line[0].project_id.id
        return svl_vals_list

    def _get_out_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_out_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            lines = self._get_out_move_lines()
            if forced_quantity:
                lot = forced_quantity[0]
                line = lines.filtered(lambda l: l.lot_id == lot)
            else:
                lot = self.env['stock.lot'].browse(vals.get('lot_id'))
                line = lines.filtered(lambda l: l.lot_id == lot)
            if line:
                vals['project_id'] = line[0].project_id.id
        return svl_vals_list

    def _get_dropshipped_svl_vals(self, forced_quantity):
        svl_vals_list = super()._get_dropshipped_svl_vals(forced_quantity)
        for vals in svl_vals_list:
            lines = self.move_line_ids
            lot_id = vals.get('lot_id')
            lot = self.env['stock.lot'].browse(lot_id) if lot_id else self.env['stock.lot']
            if forced_quantity:
                lot = forced_quantity[0]
            line = lines.filtered(lambda l: l.lot_id == lot)
            if line:
                vals['project_id'] = line[0].project_id.id
        return svl_vals_list

class StockMoveLine(models.Model):
    _inherit = 'stock.move.line'

    project_id = fields.Many2one("project.project")



