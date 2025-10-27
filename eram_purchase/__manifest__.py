# -*- coding: utf-8 -*-
{
    'name': 'Eram Purchase',
    'version': '18.0.1.0.0',
    'category': 'purchase',
    'summary': 'Eram Reports',
    'license': 'LGPL-3',

    'depends': ['eram_report_templates'],

    'data': [
        'security/ir.model.access.csv',
        'views/eram_purchase_req_views.xml',
        'views/eram_supplier_quote_views.xml',
        'views/purchase_order_views.xml',
    ],
    'installable': True,
    'auto_install': False,
    'application': False,
}
