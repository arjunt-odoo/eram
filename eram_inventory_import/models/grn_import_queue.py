# -*- coding: utf-8 -*-
import base64
import io
import logging
from datetime import datetime

from odoo import _, api, fields, models
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ---------------------------------------------------------------------------
# Column indices (0-based) matching the GRN Inward Inventory Excel layout
# ---------------------------------------------------------------------------
COL_GRN_NO        = 1
COL_INVOICE_DATE  = 2
COL_RECEIVED_DATE = 3
COL_PROJECT_CODE  = 4
COL_PR_NO         = 5
COL_PO_NUMBER     = 6
COL_INVOICE_NUMBER= 7
COL_DESCRIPTION   = 8
COL_PART_NO       = 9
COL_PO_QTY        = 10
COL_RECEIVED_QTY  = 11
COL_UNIT          = 12
COL_RATE          = 13
COL_GST           = 14
COL_SUPPLIER      = 17

DATA_START_ROW    = 3   # 0-based; rows 0-2 are title / header / blank
MAX_QUEUE_LINES   = 100


def _cell_str(row, idx):
    try:
        val = row[idx]
        if val is None:
            return None
        s = str(val).strip()
        return s if s and s.lower() not in ('none', 'nan', 'n/a', 'na', '-') else None
    except IndexError:
        return None


def _cell_float(row, idx):
    try:
        val = row[idx]
        return float(val) if val is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _cell_date(row, idx):
    val = row[idx]
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date().isoformat()
    try:
        return datetime.strptime(str(val).strip(), '%Y-%m-%d').date().isoformat()
    except ValueError:
        return None


class GrnImportQueue(models.Model):
    """
    One queue = one Excel upload.  The wizard parses the file and splits
    the GRN rows (max MAX_QUEUE_LINES per queue) into grn.import.queue.line
    records.  A permanent scheduled action processes ONE pending queue line
    per run so the server is never overloaded.
    """
    _name = 'grn.import.queue'
    _description = 'GRN Import Queue'
    _order = 'create_date desc'

    # ------------------------------------------------------------------
    # Fields
    # ------------------------------------------------------------------
    name = fields.Char(string='Queue Name', required=True, copy=False,
                       default=lambda self: _('New'))
    file_name = fields.Char(string='Source File')
    state = fields.Selection([
        ('draft',     'Draft'),
        ('in_progress','In Progress'),
        ('done',      'Done'),
        ('partial',   'Partial (with errors)'),
    ], default='draft', string='Status', copy=False)

    company_id  = fields.Many2one('res.company',  default=lambda s: s.env.company)
    user_id     = fields.Many2one('res.users',     default=lambda s: s.env.user,
                                  string='Uploaded by')
    line_ids    = fields.One2many('grn.import.queue.line', 'queue_id',
                                  string='Queue Lines')

    # Computed counters
    total_lines   = fields.Integer(compute='_compute_counts', store=True)
    pending_lines = fields.Integer(compute='_compute_counts', store=True)
    done_lines    = fields.Integer(compute='_compute_counts', store=True)
    error_lines   = fields.Integer(compute='_compute_counts', store=True)
    progress      = fields.Float(compute='_compute_counts', store=True,
                                 string='Progress (%)', digits=(5, 1))

    @api.depends('line_ids.state')
    def _compute_counts(self):
        for rec in self:
            lines = rec.line_ids
            total   = len(lines)
            pending = len(lines.filtered(lambda l: l.state == 'pending'))
            done    = len(lines.filtered(lambda l: l.state == 'done'))
            error   = len(lines.filtered(lambda l: l.state == 'error'))
            rec.total_lines   = total
            rec.pending_lines = pending
            rec.done_lines    = done
            rec.error_lines   = error
            rec.progress = ((done + error) / total * 100) if total else 0.0

    # ------------------------------------------------------------------
    # State helpers
    # ------------------------------------------------------------------

    def _refresh_state(self):
        """Recompute queue-level state from its lines."""
        for rec in self:
            lines = rec.line_ids
            if not lines:
                continue
            states = set(lines.mapped('state'))
            if states == {'done'}:
                rec.state = 'done'
            elif states == {'pending'}:
                rec.state = 'draft'
            elif 'error' in states and not {'pending', 'running'} & states:
                rec.state = 'partial'
            else:
                rec.state = 'in_progress'

    # ------------------------------------------------------------------
    # Excel parsing (called from wizard)
    # ------------------------------------------------------------------

    @api.model
    def create_from_excel(self, file_data, file_name, company_id=None):
        """
        Parse an Excel file and create one or more GrnImportQueue records,
        each holding at most MAX_QUEUE_LINES lines.

        Returns a recordset of the created queues.
        """
        if not openpyxl:
            raise UserError(_(
                "openpyxl is not installed. Run: pip install openpyxl --break-system-packages"
            ))

        raw  = base64.b64decode(file_data)
        wb   = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))

        # Group Excel rows by GRN number ─ each GRN becomes one queue line
        grn_groups = {}       # {grn_name: [row_dict, ...]}
        grn_order  = []       # preserve insertion order
        for row_idx, row in enumerate(rows):
            if row_idx < DATA_START_ROW:
                continue
            grn_name = _cell_str(row, COL_GRN_NO)
            if not grn_name:
                continue
            if grn_name not in grn_groups:
                grn_groups[grn_name] = []
                grn_order.append(grn_name)
            grn_groups[grn_name].append({
                'grn_no':         grn_name,
                'invoice_date':   _cell_date(row, COL_INVOICE_DATE),
                'received_date':  _cell_date(row, COL_RECEIVED_DATE),
                'project_code':   _cell_str(row,  COL_PROJECT_CODE),
                'pr_number':      _cell_str(row,  COL_PR_NO),
                'po_number':      _cell_str(row,  COL_PO_NUMBER),
                'invoice_number': _cell_str(row,  COL_INVOICE_NUMBER),
                'description':    _cell_str(row,  COL_DESCRIPTION),
                'part_no':        _cell_str(row,  COL_PART_NO),
                'po_qty':         _cell_float(row, COL_PO_QTY),
                'received_qty':   _cell_float(row, COL_RECEIVED_QTY),
                'unit':           _cell_str(row,  COL_UNIT),
                'rate':           _cell_float(row, COL_RATE),
                'gst':            _cell_float(row, COL_GST),
                'supplier':       _cell_str(row,  COL_SUPPLIER),
            })

        if not grn_order:
            raise UserError(_("No data rows found in the uploaded file."))

        company = self.env['res.company'].browse(company_id) if company_id \
            else self.env.company

        # Split into batches of MAX_QUEUE_LINES
        created_queues = self.env['grn.import.queue']
        batch_num  = 1
        batch_grns = []

        def _flush(batch):
            nonlocal batch_num
            seq = self.env['ir.sequence'].next_by_code('grn.import.queue') or \
                  f'GRNQ-{fields.Datetime.now():%Y%m%d%H%M%S}-{batch_num:02d}'
            queue = self.create({
                'name':      f"{seq} (batch {batch_num})" if batch_num > 1 else seq,
                'file_name': file_name,
                'company_id': company.id,
                'user_id':   self.env.uid,
            })
            import json
            line_vals = []
            for grn_name in batch:
                rows_data = grn_groups[grn_name]
                line_vals.append({
                    'queue_id':   queue.id,
                    'grn_no':     grn_name,
                    'sequence':   len(line_vals) + 1,
                    'payload':    json.dumps(rows_data),
                    'line_count': len(rows_data),
                })
            self.env['grn.import.queue.line'].create(line_vals)
            batch_num += 1
            return queue

        for grn_name in grn_order:
            batch_grns.append(grn_name)
            if len(batch_grns) >= MAX_QUEUE_LINES:
                created_queues |= _flush(batch_grns)
                batch_grns = []

        if batch_grns:
            created_queues |= _flush(batch_grns)

        return created_queues

    # ------------------------------------------------------------------
    # Manual actions
    # ------------------------------------------------------------------

    def action_reset_errors(self):
        """Re-queue all error lines so they will be retried."""
        self.mapped('line_ids').filtered(
            lambda l: l.state == 'error'
        ).write({'state': 'pending', 'error_message': False, 'picking_id': False})
        self._refresh_state()

    def action_view_lines(self):
        self.ensure_one()
        return {
            'type':      'ir.actions.act_window',
            'res_model': 'grn.import.queue.line',
            'name':      _('Queue Lines'),
            'view_mode': 'list,form',
            'domain':    [('queue_id', '=', self.id)],
            'context':   {'default_queue_id': self.id},
        }
