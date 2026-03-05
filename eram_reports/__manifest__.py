# -*- coding: utf-8 -*-
{
    'name': 'Eram Report',
    'version': '18.0.1.0.0',
    'category': 'sale',
    'summary': 'Eram Reports',
    'license': 'LGPL-3',

    'depends': ['eram_mrp'],

    'data': [
        'security/ir.model.access.csv',
        'wizard/eram_report_views.xml'
    ],
    'assets': {
        'web.assets_backend': [
            'eram_reports/static/src/js/action_manager.js',
        ]
    },
    'installable': True,
    'auto_install': False,
    'application': False,
}
