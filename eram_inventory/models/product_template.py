# -*- coding: utf-8 -*-
from odoo import _, models, fields


class ProductTemplate(models.Model):
    _inherit = 'product.template'

    e_categ_ids = fields.One2many("eram.categ", "product_tmpl_id")
    e_allow_inspection = fields.Boolean("Allow Inspection", default=False)
    is_storable = fields.Boolean(default=True)

    def _get_leaf_categories(self):
        """Get all leaf categories from selected e_categ_ids"""
        leaves_ids = set()

        def collect_leaves(cat):
            if not cat.child_ids:
                leaves_ids.add(cat.id)
            else:
                for child in cat.child_ids:
                    collect_leaves(child)

        for rec in self.e_categ_ids:
            collect_leaves(rec)

        return self.env['eram.categ'].browse(list(leaves_ids))

    def _setup_eram_attribute(self):
        """Setup single 'test' attribute with all leaf categories as values"""
        # Create or find the 'test' attribute
        attribute = self.env['product.attribute'].search(
            [('name', '=', '/')], limit=1)
        if not attribute:
            attribute = self.env['product.attribute'].create({
                'name': '/',
                'create_variant': 'always',
                'display_type': 'select',
            })

        # Get all leaf categories
        leaves = self._get_leaf_categories()
        attribute_values = []

        for leaf in leaves:
            # Use the full hierarchy path as the value name
            value_name = leaf.get_hierarchy_path()

            # Create or find attribute value
            value = self.env['product.attribute.value'].search([
                ('attribute_id', '=', attribute.id),
                ('name', '=', value_name)
            ], limit=1)

            if not value:
                value = self.env['product.attribute.value'].create({
                    'attribute_id': attribute.id,
                    'name': value_name,
                })

            attribute_values.append((4, value.id))

        # Ensure attribute line exists
        attribute_line = self.attribute_line_ids.filtered(
            lambda l: l.attribute_id == attribute
        )

        if attribute_line:
            attribute_line.write({'value_ids': attribute_values})
        else:
            self.write({
                'attribute_line_ids': [(0, 0, {
                    'attribute_id': attribute.id,
                    'value_ids': attribute_values,
                })]
            })

    def action_create_eram_variants(self):
        """Create variants using attribute system"""
        self.ensure_one()

        # Setup the single attribute first
        self._setup_eram_attribute()

        # Trigger variant creation
        self._create_variant_ids()

        # Update e_categ_id on variants
        leaves = self._get_leaf_categories()
        for variant in self.product_variant_ids:
            # Find matching leaf category based on attribute value name
            variant_attribute_value = variant.product_template_attribute_value_ids.filtered(
                lambda v: v.attribute_id._name == 'product.attribute'
            )
            if variant_attribute_value:
                path = variant_attribute_value.name
                matching_leaf = leaves.filtered(
                    lambda l: l.get_hierarchy_path() == path
                )
                if matching_leaf:
                    variant.e_categ_id = matching_leaf.id

    def action_archive(self):
        stock_move_ids = self.env['stock.move'].search([('product_id', 'in', self.product_variant_ids.mapped('id'))])
        stock_move_ids.mapped('picking_id').action_archive()
        stock_move_ids.action_archive()
        self.env['stock.valuation.layer'].search([('product_id', 'in', self.product_variant_ids.mapped('id'))]).action_archive()
        self.env['stock.quant'].search([('product_id', 'in', self.product_variant_ids.mapped('id'))]).action_archive()
        self.env['stock.lot'].search([('product_id', 'in', self.product_variant_ids.mapped('id'))]).action_archive()
        return super().action_archive()

    def action_view_invoice_lines(self):
        self.ensure_one()
        return self.product_variant_ids.action_view_invoice_lines()

    def action_view_stock_valuation(self):
        self.ensure_one()
        return self.product_variant_ids.action_view_stock_valuation()


class ProductProduct(models.Model):
    _inherit = 'product.product'

    e_categ_id = fields.Many2one("eram.categ")

    def action_view_invoice_lines(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "account.move.line",
            "name": _("Invoice Lines"),
            "view_mode": "list,form",
            "views": [(self.env.ref("eram_inventory.view_account_move_line_list_service_product").id, 'list'),
                      (False, 'form')],
            "domain": [('product_id', 'in', self.ids)],
        }

    def action_view_stock_valuation(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "res_model": "stock.valuation.layer",
            "name": _("Valuation"),
            "view_mode": "list,form",
            "domain": [('product_id', 'in', self.ids)],
        }