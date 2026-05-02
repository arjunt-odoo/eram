# -*- coding: utf-8 -*-
from odoo import _, fields, models
from odoo.exceptions import UserError


class GrnImportWizard(models.TransientModel):
    """
    Thin wizard: accepts the Excel file, delegates ALL parsing to
    GrnImportQueue.create_from_excel(), then shows the created queue(s).
    """
    _name        = 'grn.import.wizard'
    _description = 'GRN Inward Inventory Import Wizard'

    file_data  = fields.Binary(string='Excel File (.xlsx)', attachment=False)
    file_name  = fields.Char(string='File Name')
    queue_ids  = fields.Many2many('grn.import.queue', string='Created Queues',
                                  readonly=True)
    info_msg   = fields.Char(string='Info', readonly=True)

    def action_import(self):
        self.ensure_one()
        if not self.file_data:
            raise UserError(_("Please upload an Excel file before importing."))

        queues = self.env['grn.import.queue'].create_from_excel(
            file_data   = self.file_data,
            file_name   = self.file_name,
            company_id  = self.env.company.id,
        )

        total_lines = sum(queues.mapped('total_lines'))
        self.write({
            'queue_ids': [fields.Command.set(queues.ids)],
            'info_msg':  (
                f"{len(queues)} queue(s) created with {total_lines} GRN line(s) total. "
                f"The scheduled action will process one line per run."
            ),
        })

        return {
            'type':      'ir.actions.act_window',
            'res_model': self._name,
            'res_id':    self.id,
            'view_mode': 'form',
            'target':    'new',
        }

    def action_view_queues(self):
        self.ensure_one()
        if len(self.queue_ids) == 1:
            return {
                'type':      'ir.actions.act_window',
                'res_model': 'grn.import.queue',
                'res_id':    self.queue_ids.id,
                'view_mode': 'form',
                'target':    'current',
            }
        return {
            'type':      'ir.actions.act_window',
            'res_model': 'grn.import.queue',
            'name':      _('Import Queues'),
            'view_mode': 'list,form',
            'domain':    [('id', 'in', self.queue_ids.ids)],
            'target':    'current',
        }
