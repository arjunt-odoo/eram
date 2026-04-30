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
    raise UserError(_("openpyxl is required. Install it via: pip install openpyxl"))


# ---------------------------------------------------------------------------
# Column indices (0-based) matching the Excel layout
# ---------------------------------------------------------------------------
COL_SL_NO = 0
COL_GRN_NO = 1
COL_INVOICE_DATE = 2
COL_RECEIVED_DATE = 3
COL_PROJECT_CODE = 4
COL_PR_NO = 5
COL_PO_NUMBER = 6
COL_INVOICE_NUMBER = 7
COL_DESCRIPTION = 8
COL_PART_NO = 9
COL_PO_QTY = 10
COL_RECEIVED_QTY = 11
COL_UNIT = 12
COL_RATE = 13
COL_GST = 14
COL_TOTAL_AMOUNT = 15
COL_GRAND_TOTAL = 16
COL_SUPPLIER = 17

HEADER_ROW_INDEX = 1   # 0-based: row 2 in Excel
DATA_START_ROW = 3     # 0-based: row 4 in Excel (skip title + header + blank)


def _cell_val(row, col_idx):
    """Return stripped string value or None from a tuple-row."""
    try:
        val = row[col_idx]
        if val is None:
            return None
        s = str(val).strip()
        return s if s and s.lower() not in ('none', 'nan', 'n/a', 'na') else None
    except IndexError:
        return None


def _cell_float(row, col_idx):
    try:
        val = row[col_idx]
        if val is None:
            return 0.0
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _cell_date(row, col_idx):
    val = row[col_idx]
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date()
    try:
        return datetime.strptime(str(val).strip(), '%Y-%m-%d').date()
    except ValueError:
        return None


class GrnImportWizard(models.TransientModel):
    _name = 'grn.import.wizard'
    _description = 'GRN Inward Inventory Excel Import Wizard'

    file_data = fields.Binary(string='Excel File (.xlsx)', required=True, attachment=False)
    file_name = fields.Char(string='File Name')
    import_log = fields.Text(string='Import Log', readonly=True)
    state = fields.Selection([
        ('draft', 'Draft'),
        ('done', 'Done'),
    ], default='draft')

    # ------------------------------------------------------------------
    # Helpers: search-or-create patterns
    # ------------------------------------------------------------------

    def _get_or_create_partner(self, name):
        if not name:
            return self.env['res.partner']
        partner = self.env['res.partner'].search([('name', 'ilike', name)], limit=1)
        if not partner:
            partner = self.env['res.partner'].create({'name': name, 'supplier_rank': 1})
        return partner

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
        pr = self.env['eram.purchase.req'].search([('pr_number', '=', pr_number)], limit=1)
        if not pr:
            # Minimal creation – adapt field names to your eram.purchase.req model
            pr = self.env['eram.purchase.req'].create({'pr_number': pr_number})
        return pr

    def _get_or_create_project_task(self, project_code):
        """
        project_code format: '<PROJECT_NAME>-<TASK_NAME>'
        Everything before the first '-' is the project name;
        everything after is the task name.

        When a new task is created, the ProjectTask.create() hook automatically
        builds all operation types (receipt_type_id, delivery_type_id, etc.) and
        writes them back onto the task.  We invalidate the cache after creation so
        that task.receipt_type_id is visible in the same transaction.
        """
        if not project_code:
            return self.env['project.project'], self.env['project.task']

        parts = project_code.split('-', 1)
        project_name = parts[0].strip()
        task_name = parts[1].strip() if len(parts) > 1 else project_code.strip()

        project = self.env['project.project'].search([('name', '=', project_name)], limit=1)
        if not project:
            project = self.env['project.project'].create({'name': project_name})

        task = self.env['project.task'].search([
            ('name', '=', task_name),
            ('project_id', '=', project.id),
        ], limit=1)
        if not task:
            task = self.env['project.task'].create({
                'name': task_name,
                'project_id': project.id,
            })
            # The create hook writes receipt_type_id back via rec.write({...}).
            # Invalidate the ORM cache so we read the freshly committed value.
            task.invalidate_recordset(['receipt_type_id'])

        return project, task

    def _get_or_create_uom(self, unit_name):
        if not unit_name:
            uom = self.env.ref('uom.product_uom_unit', raise_if_not_found=False)
            return uom or self.env['uom.uom']
        uom = self.env['uom.uom'].search([('name', 'ilike', unit_name)], limit=1)
        if not uom:
            categ = self.env.ref('uom.product_uom_categ_unit', raise_if_not_found=False)
            if not categ:
                categ = self.env['uom.category'].search([], limit=1)
            uom = self.env['uom.uom'].create({
                'name': unit_name,
                'category_id': categ.id if categ else False,
                'uom_type': 'bigger',
                'factor': 1.0,
            })
        return uom

    def _get_or_create_product(self, description, uom):
        if not description:
            description = 'Unknown Product'
        product_tmpl = self.env['product.template'].search(
            [('name', 'ilike', description.strip())], limit=1
        )
        if not product_tmpl:
            product_tmpl = self.env['product.template'].create({
                'name': description.strip(),
                'type': 'consu',
                'uom_id': uom.id if uom else False,
                'uom_po_id': uom.id if uom else False,
                'purchase_ok': True,
            })
        return product_tmpl.product_variant_ids[:1]

    def _find_closest_tax(self, gst_amount, subtotal, company_id):
        """
        Derive approximate tax % from gst_amount / subtotal * 100
        and find the nearest purchase tax in the system.
        Returns a recordset of account.tax.
        """
        if not subtotal or not gst_amount:
            return self.env['account.tax']

        approx_pct = round((gst_amount / subtotal) * 100)
        # Round to nearest common GST slab: 0, 5, 12, 18, 28
        slabs = [0, 5, 12, 18, 28]
        closest_slab = min(slabs, key=lambda s: abs(s - approx_pct))

        if closest_slab == 0:
            return self.env['account.tax']

        tax = self.env['account.tax'].search([
            ('type_tax_use', '=', 'purchase'),
            ('amount', '=', closest_slab),
            ('amount_type', '=', 'percent'),
            ('company_id', '=', company_id),
        ], limit=1)

        if not tax:
            tax = self.env['account.tax'].create({
                'name': f'GST {closest_slab}%',
                'type_tax_use': 'purchase',
                'amount_type': 'percent',
                'amount': closest_slab,
                'company_id': company_id,
            })
        return tax

    def _get_or_create_invoice(self, invoice_number, invoice_date, partner, picking_move_vals, company):
        """
        Search for an existing vendor bill with this number, or create a new draft one.
        Lines will be added based on move_vals list: [{product, qty, price_unit, tax_ids, name}, ...]
        """
        invoice = self.env['account.move'].search([
            ('name', '=', invoice_number),
            ('move_type', '=', 'in_invoice'),
        ], limit=1)

        if invoice:
            return invoice

        invoice_line_vals = []
        for mv in picking_move_vals:
            line_vals = {
                'product_id': mv.get('product_id'),
                'name': mv.get('description') or '/',
                'quantity': mv.get('quantity', 1.0),
                'price_unit': mv.get('price_unit', 0.0),
                'tax_ids': [fields.Command.set(mv.get('tax_ids', []))],
            }
            if mv.get('product_uom_id'):
                line_vals['product_uom_id'] = mv['product_uom_id']
            invoice_line_vals.append(fields.Command.create(line_vals))

        invoice = self.env['account.move'].create({
            'move_type': 'in_invoice',
            'partner_id': partner.id if partner else False,
            'invoice_date': invoice_date,
            'company_id': company.id,
            'invoice_line_ids': invoice_line_vals,
        })

        # Override the auto-generated name with the one from Excel
        if invoice_number:
            invoice.write({'name': invoice_number})

        return invoice

    # ------------------------------------------------------------------
    # Core import logic
    # ------------------------------------------------------------------

    def action_import(self):
        self.ensure_one()
        if not self.file_data:
            raise UserError(_("Please upload an Excel file before importing."))

        raw = base64.b64decode(self.file_data)
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))

        company = self.env.company
        logs = []
        errors = []

        # ---------------------------------------------------------------
        # Group rows by GRN number (each GRN → one stock.picking)
        # ---------------------------------------------------------------
        grn_groups = {}   # {grn_name: [row, ...]}
        for row_idx, row in enumerate(rows):
            if row_idx < DATA_START_ROW:
                continue
            grn_name = _cell_val(row, COL_GRN_NO)
            if not grn_name:
                continue
            grn_groups.setdefault(grn_name, []).append((row_idx, row))

        total_pickings = 0
        total_moves = 0

        for grn_name, row_list in grn_groups.items():
            try:
                picking, move_count = self._process_grn_group(grn_name, row_list, company)
                total_pickings += 1
                total_moves += move_count
                logs.append(f"✓ GRN {grn_name}: created picking {picking.name} with {move_count} move(s).")
            except Exception as exc:
                _logger.exception("Error importing GRN %s", grn_name)
                errors.append(f"✗ GRN {grn_name}: {exc}")

        summary = [
            f"Import complete.",
            f"  Pickings created : {total_pickings}",
            f"  Stock moves      : {total_moves}",
            f"  Errors           : {len(errors)}",
            "",
        ]
        self.write({
            'import_log': "\n".join(summary + logs + errors),
            'state': 'done',
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def _process_grn_group(self, grn_name, row_list, company):
        """
        Process all Excel rows that share the same GRN number.
        Creates/updates:
          - eram.grn
          - res.partner  (supplier)
          - project.project + project.task
          - eram.purchase.req
          - uom.uom
          - product.product
          - account.tax
          - account.move (vendor bill)
          - stock.picking + stock.move lines
        """
        # Use first row for picking-level fields
        first_row = row_list[0][1]

        invoice_date = _cell_date(first_row, COL_INVOICE_DATE)
        received_date = _cell_date(first_row, COL_RECEIVED_DATE)
        project_code = _cell_val(first_row, COL_PROJECT_CODE)
        pr_number = _cell_val(first_row, COL_PR_NO)
        po_number = _cell_val(first_row, COL_PO_NUMBER)
        invoice_number = _cell_val(first_row, COL_INVOICE_NUMBER)
        supplier_name = _cell_val(first_row, COL_SUPPLIER)

        # --- Resolve picking-level records ---
        partner = self._get_or_create_partner(supplier_name)
        grn = self._get_or_create_grn(grn_name)
        project, task = self._get_or_create_project_task(project_code)
        pr = self._get_or_create_pr(pr_number) if pr_number else self.env['eram.purchase.req']

        # Determine operation type from task.receipt_type_id.
        # The ProjectTask.create() hook names it "{task_name}: Inward" with
        # sequence_code "{project_name}-{task_name}-IN", so we mirror that
        # pattern as a reliable fallback when the field is not yet populated.
        picking_type = task.receipt_type_id if task and task.receipt_type_id else False

        if not picking_type and task and project:
            task_name_str = task.name
            project_name_str = project.name
            picking_type = self.env['stock.picking.type'].search([
                ('name', '=', f"{task_name_str}: Inward"),
                ('code', '=', 'incoming'),
                ('sequence_code', '=', f"{project_name_str}-{task_name_str}-IN"),
            ], limit=1)

        if not picking_type:
            raise UserError(_(
                "No incoming operation type found for task '%s' (project '%s'). "
                "Expected an operation type named '%s: Inward' with sequence code "
                "'%s-%s-IN'. Please verify the task was created correctly."
            ) % (
                task.name if task else 'N/A',
                project.name if project else 'N/A',
                task.name if task else '?',
                project.name if project else '?',
                task.name if task else '?',
            ))

        # --- Build move_vals list (needed for invoice creation) ---
        move_vals_list = []
        for _row_idx, row in row_list:
            description = _cell_val(row, COL_DESCRIPTION) or 'Unknown'
            part_no = _cell_val(row, COL_PART_NO)
            po_qty = _cell_float(row, COL_PO_QTY)
            received_qty = _cell_float(row, COL_RECEIVED_QTY)
            unit_name = _cell_val(row, COL_UNIT)
            rate = _cell_float(row, COL_RATE)
            gst_amount = _cell_float(row, COL_GST)
            subtotal = rate * received_qty if received_qty else rate * po_qty

            uom = self._get_or_create_uom(unit_name)
            product = self._get_or_create_product(description, uom)
            tax = self._find_closest_tax(gst_amount, subtotal, company.id)

            move_vals_list.append({
                'description': description,
                'part_no': part_no,
                'product_id': product.id if product else False,
                'product_uom_id': uom.id if uom else False,
                'po_qty': po_qty,
                'quantity': received_qty or po_qty,
                'price_unit': rate,
                'tax_ids': tax.ids,
                'uom': uom,
                'product': product,
                'tax': tax,
            })

        # --- Create vendor bill ---
        invoice = False
        if invoice_number:
            invoice = self._get_or_create_invoice(
                invoice_number, invoice_date, partner, move_vals_list, company
            )

        # --- Create stock.picking ---
        picking_vals = {
            'picking_type_id': picking_type.id,
            'partner_id': partner.id if partner else False,
            'scheduled_date': received_date or fields.Datetime.now(),
            'company_id': company.id,
            'e_grn_id': grn.id,
            'e_project_id': project.id if project else False,
            'e_task_id': task.id if task else False,
            'e_po_no': po_number,
        }
        if pr:
            picking_vals['e_pr_id'] = pr.id
        if invoice:
            picking_vals['e_bill_id'] = invoice.id

        picking = self.env['stock.picking'].create(picking_vals)

        # Sync GRN → picking
        if grn:
            grn.picking_id = picking

        # --- Create stock.move records ---
        for mv in move_vals_list:
            product = mv['product']
            uom = mv['uom']
            tax = mv['tax']

            if not product:
                continue

            dest_location = picking_type.default_location_dest_id
            src_location = picking_type.default_location_src_id or \
                self.env.ref('stock.stock_location_suppliers', raise_if_not_found=False)

            move = self.env['stock.move'].create({
                'name': mv['description'],
                'picking_id': picking.id,
                'product_id': product.id,
                'product_uom_qty': mv['po_qty'] or mv['quantity'],
                'quantity': mv['quantity'],
                'product_uom': uom.id,
                'price_unit': mv['price_unit'],
                'location_id': src_location.id if src_location else dest_location.id,
                'location_dest_id': dest_location.id,
                'e_description': mv['description'],
                'e_part_no': mv['part_no'] or '',
                'e_tax_ids': [fields.Command.set(tax.ids)],
                'state': 'draft',
            })

        return picking, len(move_vals_list)
