from odoo import models, fields, api

class VariantConfiguratorWizard(models.TransientModel):
    _name = 'variant.configurator.wizard'
    _description = 'Hierarchical Variant Configurator'

    product_id = fields.Many2one('product.template', required=True)
    current_level = fields.Many2one('product.attribute.value', string='Current Selection')
    available_children = fields.Many2many('product.attribute.value', compute='_compute_children')
    config_path = fields.Char(string='Configuration Path', compute='_compute_path')
    final_variant_id = fields.Many2one('product.product', string='Generated Variant')

    @api.depends('current_level')
    def _compute_children(self):
        for rec in self:
            rec.available_children = rec.current_level.child_ids if rec.current_level else self.product_id.attribute_line_ids.product_tmpl_id_value_ids

    @api.depends('current_level')
    def _compute_path(self):
        for rec in self:
            rec.config_path = rec.current_level.get_hierarchy_path() if rec.current_level else ''

    def action_next_level(self):
        # Advance to child selection; recurse via wizard reopen
        return {'type': 'ir.actions.act_window', 'res_model': self._name, 'view_mode': 'form', 'res_id': self.id}

    def action_generate_variant(self):
        # Create variant on-the-fly if full path selected
        if self.current_level:
            variant_name = self.config_path
            variant = self.env['product.product'].create({
                'product_tmpl_id': self.product_id.id,
                'name': variant_name,
                # Add other fields: default_code, list_price, etc., based on path
            })
            self.final_variant_id = variant
        return {'type': 'ir.actions.act_window_close'}