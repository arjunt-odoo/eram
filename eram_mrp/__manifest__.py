# -*- coding: utf-8 -*-
{
    'name': 'Eram Manufacturing',
    'version': '18.0.1.0.0',
    'category': 'Inventory',
    'summary': 'Eram Manufacturing',
    'license': 'LGPL-3',
    'depends': ['eram_inventory'],
    'data': [
        'security/ir.model.access.csv',
        'data/ir_sequence.xml',
        'views/hr_department_views.xml',
        'views/mrp_production_views.xml',
        'views/project_project_views.xml',
        'views/stock_picking_views.xml',
        'report/eram_outward_templates.xml',
        'report/eram_outward_reports.xml',
    ],
    'installable': True,
    'auto_install': False,
}