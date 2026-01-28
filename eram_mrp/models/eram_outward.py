from odoo import api, fields, models


class EramOutward(models.Model):
    _name = "eram.outward"
    _description = "Eram Outward"

    name = fields.Char("Name")
    date = fields.Date()
    project_id = fields.Many2one("project.project")
    task_id = fields.Many2one("project.task")
    justification = fields.Char()
    requested_by_id = fields.Many2one("hr.employee")
    approved_by_id = fields.Many2one("hr.employee")
    requested_by_dept_id = fields.Many2one("hr.department", related="requested_by_id.department_id")
    approved_by_dept_id = fields.Many2one("hr.department", related="approved_by_id.department_id")
    line_ids = fields.One2many("eram.outward.line", "outward_id")
    production_id = fields.Many2one("mrp.production")

    @api.model_create_multi
    def create(self, vals_list):
        today = fields.Date.today()
        current_month = today.month

        if current_month >= 4:
            current_fy_year = today.year
            next_fy_year = today.year + 1
        else:
            current_fy_year = today.year - 1
            next_fy_year = today.year

        for val in vals_list:
            batch_size = 50
            seq = self.env['ir.sequence'].sudo()
            position_str = seq.next_by_code('eram.outward.position')
            position = int(position_str)
            batch_seq = seq.search([('code', '=', 'eram.outward.batch')], limit=1)
            batch = batch_seq.number_next
            if position > batch_size:
                seq.next_by_code('eram.outward.batch')
                position_seq = seq.search([('code', '=', 'eram.outward.position')], limit=1)
                position_seq.write({'number_next': 2})
                position = 1
                batch += 1
            middle = f"{batch:03d}"
            last = f"{position:04d}"
            val['name'] = f"BLR-MWF-{current_fy_year}-{next_fy_year}-R0-{middle}-{last}"

        return super().create(vals_list)


class EramOutwardLine(models.Model):
    _name = "eram.outward.line"
    _description = "Eram outward line"

    si_no = fields.Integer(compute="_compute_si_no", store=True)
    product_id = fields.Many2one("product.product")
    description = fields.Char()
    uom_id = fields.Many2one("uom.uom")
    product_uom_category_id = fields.Many2one("uom.category", related="product_id.uom_category_id")
    qty = fields.Float()
    required_date = fields.Date()
    remarks = fields.Char()
    outward_id = fields.Many2one("eram.outward")

    @api.depends('outward_id', 'outward_id.line_ids')
    def _compute_si_no(self):
        for outward in self.mapped('outward_id'):
            for index, line in enumerate(outward.line_ids, start=1):
                line.si_no = index