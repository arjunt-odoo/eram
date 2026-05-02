# -*- coding: utf-8 -*-
{
    'name': 'ERAM GRN Import Queue',
    'version': '18.0.3.0.0',
    'category': 'Inventory',
    'summary': 'Queue-based bulk import of GRN Inward Inventory from Excel',
    'description': """
        Upload a GRN Inward Inventory Excel file.  The wizard parses it and
        creates one or more import queues (max 100 GRN lines each).

        A permanent scheduled action processes exactly ONE queue line per run
        (default: every minute), creating the corresponding stock.picking and
        stock.moves without ever blocking an HTTP worker or hitting a timeout.
    """,
    'author': 'ERAM',
    'depends': [
        'eram_inventory',
    ],
    'data': [
        'security/ir.model.access.csv',
        'data/grn_import_cron.xml',
        'views/grn_import_views.xml',
        'views/menu.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
