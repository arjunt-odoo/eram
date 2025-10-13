# -*- coding: utf-8 -*-
from odoo import _, api, fields, models
from odoo.exceptions import UserError


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    e_invoice_id = fields.Many2one("account.move", string="Invoice Number")

    def action_reset_to_draft(self):
        """
        Reset a done picking back to draft state and reverse all stock movements
        """
        for picking in self:
            if picking.state != 'done':
                raise UserError(
                    _("You can only reset pickings that are in 'Done' state."))

            if not picking._can_reset_to_draft():
                raise UserError(_("This transfer cannot be reset to draft."))

            # Reverse all stock moves
            picking._reverse_stock_moves()

            # Reset picking state
            picking.write({
                'state': 'draft',
                'date_done': False,
                'is_locked': False,
            })

            # Reset related moves
            picking.move_ids.write({
                'state': 'draft',
                'quantity': 0,
                'picked': False,
            })

            # Delete move lines
            picking.move_line_ids.unlink()

        return True

    def _can_reset_to_draft(self):
        """
        Check if the picking can be reset to draft
        """
        self.ensure_one()

        # Can't reset if there are backorders or returns
        if self.backorder_ids or self.return_ids:
            return False

        # Can't reset if there are destination moves that are done
        done_dest_moves = self.move_ids.move_dest_ids.filtered(
            lambda m: m.state == 'done')
        if done_dest_moves:
            return False

        return True

    def _reverse_stock_moves(self):
        """
        Reverse all stock movements by creating reverse moves
        """
        for move in self.move_ids:
            if move.state == 'done' and not move.scrapped:
                # Create reverse quant movement
                self._reverse_quant_movement(move)

    def _reverse_quant_movement(self, move):
        """
        Reverse the quant movement for a specific move
        """
        for move_line in move.move_line_ids:
            # Reverse the quant movement: from destination back to source
            quantity = move_line.product_uom_id._compute_quantity(
                move_line.quantity,
                move_line.product_id.uom_id
            )

            # Remove from destination location
            self.env['stock.quant']._update_available_quantity(
                move_line.product_id,
                move_line.location_dest_id,
                -quantity,
                lot_id=move_line.lot_id,
                package_id=move_line.package_id,
                owner_id=move_line.owner_id
            )

            # Add back to source location
            self.env['stock.quant']._update_available_quantity(
                move_line.product_id,
                move_line.location_id,
                quantity,
                lot_id=move_line.lot_id,
                package_id=move_line.package_id,
                owner_id=move_line.owner_id
            )