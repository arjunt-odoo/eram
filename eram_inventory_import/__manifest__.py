# -*- coding: utf-8 -*-
{
    'name': 'ERAM Inventory Import',
    'version': '18.0.1.0.0',
    'category': 'Inventory',
    'summary': 'Bulk import GRN Inward Inventory from Excel',
    'description': """
        Import stock.picking (receipts) records in bulk from a structured
        Excel file matching the GRN Inward Inventory format.
    """,
    'author': 'ERAM',
    'depends': [
        'eram_inventory',
    ],
    'data': [
        'security/ir.model.access.csv',
        'wizard/grn_import_wizard_view.xml',
        'views/menu.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
