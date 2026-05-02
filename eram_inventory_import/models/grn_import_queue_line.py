# -*- coding: utf-8 -*-
import json
import logging

from odoo import _, fields, models
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)


class GrnImportQueueLine(models.Model):
    """
    One line = one GRN number = one future stock.picking.

    The permanent scheduled action picks lines in FIFO order
    (oldest queue first, lowest sequence within queue first),
    processes exactly ONE line per run, commits, then stops.
    """
    _name        = 'grn.import.queue.line'
    _description = 'GRN Import Queue Line'
    _order       = 'queue_id asc, sequence asc, id asc'

    # ------------------------------------------------------------------
    # Fields
    # ------------------------------------------------------------------
    queue_id      = fields.Many2one('grn.import.queue', string='Queue',
                                    required=True, ondelete='cascade', index=True)
    sequence      = fields.Integer(default=10)
    grn_no        = fields.Char(string='GRN No.', required=True)
    line_count    = fields.Integer(string='Excel Rows')
    payload       = fields.Text(string='Row Data (JSON)', help='Serialised list of Excel row dicts')
    state         = fields.Selection([
        ('pending', 'Pending'),
        ('running', 'Running'),
        ('done',    'Done'),
        ('error',   'Error'),
    ], default='pending', string='Status', index=True)
    error_message = fields.Text(string='Error Detail')
    picking_id    = fields.Many2one('stock.picking', string='Created Receipt')
    processed_at  = fields.Datetime(string='Processed At')

    # ------------------------------------------------------------------
    # Scheduled-action entry point
    # ------------------------------------------------------------------

    @classmethod
    def _cron_process_next(cls, env):
        """
        Process exactly ONE pending queue line and commit.
        Called by the permanent ir.cron every N minutes.
        """
        lines = env['grn.import.queue.line'].search(
            [('state', '=', 'pending')],
            order='queue_id asc, sequence asc, id asc',
            limit=50,
        )
        if not lines:
            _logger.info("GRN import queue: nothing pending.")
            return
        for line in lines:
            line._process()

    def _process(self):
        """
        Process this single line: create stock.picking + stock.moves for the GRN.
        Commits on success, rolls back and marks error on failure.
        Updates the parent queue state after each line.

        IMPORTANT: we stash line_id and queue_id as plain ints BEFORE any
        try/except so that after a rollback (which wipes the ORM cache and
        invalidates self) we can re-browse both records cleanly.
        """
        self.ensure_one()
        line_id  = self.id
        queue_id = self.queue_id.id

        self.write({'state': 'running'})
        self.env.cr.commit()

        try:
            row_list = json.loads(self.payload or '[]')
            company  = self.queue_id.company_id or self.env.company
            picking  = self._build_picking(self.grn_no, row_list, company)

            self.write({
                'state':         'done',
                'picking_id':    picking.id,
                'processed_at':  fields.Datetime.now(),
                'error_message': False,
            })
            self.env.cr.commit()
            _logger.info("GRN queue line %s: picking %s created.", line_id, picking.name)

        except Exception as exc:
            # Log with local vars only — self.grn_no would hit the dead transaction
            _logger.exception("GRN queue line %s failed: %s", line_id, exc)
            # Rollback resets the connection; ORM cache is now stale
            self.env.cr.rollback()
            # Re-browse in the now-clean transaction
            self.env['grn.import.queue.line'].browse(line_id).write({
                'state':         'error',
                'error_message': str(exc),
                'processed_at':  fields.Datetime.now(),
            })
            self.env.cr.commit()

        # Always refresh queue state in a clean transaction (no finally needed)
        self.env['grn.import.queue'].browse(queue_id)._refresh_state()
        self.env.cr.commit()

    # ------------------------------------------------------------------
    # Search-or-create helpers (same logic as before, scoped to this model)
    # ------------------------------------------------------------------

    def _get_or_create_partner(self, name):
        if not name:
            return self.env['res.partner']
        p = self.env['res.partner'].search([('name', 'ilike', name)], limit=1)
        if not p:
            p = self.env['res.partner'].create({'name': name, 'supplier_rank': 1})
        return p

    def _get_or_create_grn(self, grn_name):
        if not grn_name:
            return self.env['eram.grn']
        grn = self.env['eram.grn'].search([('name', '=', grn_name)], limit=1)
        if not grn:
            grn = self.env['eram.grn'].create({'name': grn_name})
        return grn

    def _get_or_create_pr(self, pr_number):
        if not pr_number:
            return self.env['eram.purchase.req']
        pr = self.env['eram.purchase.req'].search(
            [('pr_number', '=', pr_number)], limit=1)
        if not pr:
            pr = self.env['eram.purchase.req'].create({'pr_number': pr_number})
        return pr

    def _get_or_create_project_task(self, project_code):
        if not project_code:
            return self.env['project.project'], self.env['project.task']

        parts        = project_code.split('-', 1)
        project_name = parts[0].strip()
        task_name    = parts[1].strip() if len(parts) > 1 else project_code.strip()

        project = self.env['project.project'].search(
            [('name', '=', project_name)], limit=1)
        if not project:
            project = self.env['project.project'].create({'name': project_name})

        task = self.env['project.task'].search([
            ('name', '=', task_name),
            ('project_id', '=', project.id),
        ], limit=1)
        if not task:
            task = self.env['project.task'].create({
                'name':       task_name,
                'project_id': project.id,
            })
            # ProjectTask.create() writes receipt_type_id back via rec.write().
            # Invalidate cache so we read the freshly committed value.
            task.invalidate_recordset(['receipt_type_id'])

        return project, task

    def _get_or_create_uom(self, unit_name):
        if not unit_name:
            return self.env.ref('uom.product_uom_unit', raise_if_not_found=False) \
                   or self.env['uom.uom']
        uom = self.env['uom.uom'].search([('name', 'ilike', unit_name)], limit=1)
        if not uom:
            categ = self.env.ref('uom.product_uom_categ_unit', raise_if_not_found=False) \
                    or self.env['uom.category'].search([], limit=1)
            uom = self.env['uom.uom'].create({
                'name':        unit_name,
                'category_id': categ.id if categ else False,
                'uom_type':    'bigger',
                'factor':      1.0,
            })
        return uom

    def _get_or_create_product(self, description, uom):
        name = (description or 'Unknown Product').strip()
        tmpl = self.env['product.template'].search([('name', 'ilike', name)], limit=1)
        if not tmpl:
            tmpl = self.env['product.template'].create({
                'name':       name,
                'type':       'consu',
                'uom_id':     uom.id if uom else False,
                'uom_po_id':  uom.id if uom else False,
                'purchase_ok': True,
            })
        return tmpl.product_variant_ids[:1]

    def _find_closest_tax(self, gst_amount, subtotal, company_id):
        if not subtotal or not gst_amount:
            return self.env['account.tax']
        approx_pct   = round((gst_amount / subtotal) * 100)
        slabs        = [0, 5, 12, 18, 28]
        closest_slab = min(slabs, key=lambda s: abs(s - approx_pct))
        if closest_slab == 0:
            return self.env['account.tax']
        tax = self.env['account.tax'].search([
            ('type_tax_use', '=', 'purchase'),
            ('amount',       '=', closest_slab),
            ('amount_type',  '=', 'percent'),
            ('company_id',   '=', company_id),
        ], limit=1)
        if not tax:
            # tax_group_id is NOT NULL in Odoo 18 — find or create a GST group
            tax_group = self.env['account.tax.group'].search([
                ('name', 'ilike', 'GST'),
                ('company_id', '=', company_id),
            ], limit=1)
            if not tax_group:
                tax_group = self.env['account.tax.group'].search([
                    ('company_id', '=', company_id),
                ], limit=1)
            tax = self.env['account.tax'].create({
                'name':         f'GST {closest_slab}%',
                'type_tax_use': 'purchase',
                'amount_type':  'percent',
                'amount':       closest_slab,
                'company_id':   company_id,
                'tax_group_id': tax_group.id if tax_group else False,
            })
        return tax

    def _get_or_create_invoice(self, invoice_number, invoice_date,
                               partner, move_vals_list, company):
        if not invoice_number:
            return False

        invoice = self.env['account.move'].search([
            ('name',      '=', invoice_number),
            ('move_type', '=', 'in_invoice'),
        ], limit=1)
        if invoice:
            return invoice

        line_cmds = []
        for mv in move_vals_list:
            line_cmds.append(fields.Command.create({
                'product_id':    mv.get('product_id'),
                'name':          mv.get('description') or '/',
                'quantity':      mv.get('quantity', 1.0),
                'price_unit':    mv.get('price_unit', 0.0),
                'tax_ids':       [fields.Command.set(mv.get('tax_ids', []))],
                **(({'product_uom_id': mv['product_uom_id']}
                    if mv.get('product_uom_id') else {})),
            }))

        invoice = self.env['account.move'].create({
            'move_type':         'in_invoice',
            'partner_id':        partner.id if partner else False,
            'invoice_date':      invoice_date,
            'company_id':        company.id,
            'invoice_line_ids':  line_cmds,
        })
        invoice.write({'name': invoice_number})
        return invoice

    # ------------------------------------------------------------------
    # Core: build one stock.picking from the row list
    # ------------------------------------------------------------------

    def _build_picking(self, grn_name, row_list, company):
        """
        Create stock.picking + stock.moves for a single GRN.
        row_list is a list of row dicts as stored in the payload JSON.
        """
        first = row_list[0]

        invoice_date   = first.get('invoice_date')
        received_date  = first.get('received_date')
        project_code   = first.get('project_code')
        pr_number      = first.get('pr_number')
        po_number      = first.get('po_number')
        invoice_number = first.get('invoice_number')
        supplier_name  = first.get('supplier')

        partner             = self._get_or_create_partner(supplier_name)
        grn                 = self._get_or_create_grn(grn_name)
        project, task       = self._get_or_create_project_task(project_code)
        pr                  = self._get_or_create_pr(pr_number) if pr_number \
                              else self.env['eram.purchase.req']

        # --- Resolve operation type ---
        picking_type = (task.receipt_type_id
                        if task and task.receipt_type_id else False)

        if not picking_type and task and project:
            picking_type = self.env['stock.picking.type'].search([
                ('name',          '=', f"{task.name}: Inward"),
                ('code',          '=', 'incoming'),
                ('sequence_code', '=', f"{project.name}-{task.name}-IN"),
            ], limit=1)

        if not picking_type:
            raise UserError(_(
                "No incoming operation type found for task '%s' (project '%s'). "
                "Expected '%s: Inward' with sequence code '%s-%s-IN'."
            ) % (
                task.name    if task    else 'N/A',
                project.name if project else 'N/A',
                task.name    if task    else '?',
                project.name if project else '?',
                task.name    if task    else '?',
            ))

        # --- Build move value dicts ---
        move_vals_list = []
        for row in row_list:
            description  = row.get('description') or 'Unknown'
            part_no      = row.get('part_no')
            po_qty       = row.get('po_qty',       0.0)
            received_qty = row.get('received_qty', 0.0)
            unit_name    = row.get('unit')
            rate         = row.get('rate',  0.0)
            gst_amount   = row.get('gst',   0.0)
            subtotal     = rate * (received_qty or po_qty)

            uom     = self._get_or_create_uom(unit_name)
            product = self._get_or_create_product(description, uom)
            tax     = self._find_closest_tax(gst_amount, subtotal, company.id)

            move_vals_list.append({
                'description':    description,
                'part_no':        part_no,
                'product_id':     product.id if product else False,
                'product_uom_id': uom.id     if uom     else False,
                'po_qty':         po_qty,
                'quantity':       received_qty or po_qty,
                'price_unit':     rate,
                'tax_ids':        tax.ids,
                'uom':            uom,
                'product':        product,
                'tax':            tax,
            })

        # --- Create vendor bill ---
        invoice = self._get_or_create_invoice(
            invoice_number, invoice_date, partner, move_vals_list, company
        )

        # --- Create stock.picking ---
        picking_vals = {
            'picking_type_id': picking_type.id,
            'partner_id':      partner.id if partner else False,
            'scheduled_date':  received_date or fields.Datetime.now(),
            'company_id':      company.id,
            'e_grn_id':        grn.id if grn else False,
            'e_project_id':    project.id if project else False,
            'e_task_id':       task.id    if task    else False,
            'e_po_no':         po_number,
        }
        if pr:
            picking_vals['e_pr_id'] = pr.id
        if invoice:
            picking_vals['e_bill_id'] = invoice.id

        picking = self.env['stock.picking'].create(picking_vals)
        if grn:
            grn.picking_id = picking

        # --- Create stock.move per row ---
        dest_location = picking_type.default_location_dest_id
        src_location  = picking_type.default_location_src_id or \
            self.env.ref('stock.stock_location_suppliers', raise_if_not_found=False)

        for mv in move_vals_list:
            if not mv['product']:
                continue
            self.env['stock.move'].create({
                'name':             mv['description'],
                'picking_id':       picking.id,
                'product_id':       mv['product'].id,
                'product_uom_qty':  mv['po_qty'] or mv['quantity'],
                'quantity':         mv['quantity'],
                'product_uom':      mv['uom'].id,
                'price_unit':       mv['price_unit'],
                'location_id':      src_location.id  if src_location  else dest_location.id,
                'location_dest_id': dest_location.id,
                'e_description':    mv['description'],
                'e_part_no':        mv['part_no'] or '',
                'e_tax_ids':        [fields.Command.set(mv['tax'].ids)],
                'state':            'draft',
            })

        return picking
