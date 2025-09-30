# -*- coding: utf-8 -*-
{
    'name': 'Eram Reports',
    'version': '18.0.1.0.0',
    'category': 'sale',
    'summary': 'Eram Reports',
    'license': 'LGPL-3',

    'depends': ['sale_management','purchase', 'l10n_in', 'stock'],

    'data': [
        'data/ir_cron_data.xml',
        'data/ir_sequence.xml',
        'data/mail_template_data.xml',
        'security/ir.model.access.csv',
        'report/external_layout_templates.xml',
        'report/ir_actions_report.xml',
        'report/invoice_report_templates.xml',
        'report/purchase_templates.xml',
        'report/sale_templates.xml',
        'views/eram_customer_po_views.xml',
        'views/account_move_views.xml',
        'views/purchase_order_views.xml',
        'views/res_config_settings_views.xml',
        'views/sale_order_views.xml',
        'views/stock_picking_views.xml',
        'wizard/eram_sale_order_report_views.xml'
    ],
    'assets': {
        'web.assets_backend': [
            'eram_report_templates/static/src/js/action_manager.js',
            'eram_report_templates/static/src/payment_notify/payment_notify_icon.xml',
            'eram_report_templates/static/src/payment_notify/payment_notify_icon.js',
        ]
    },
    'installable': True,
    'auto_install': False,
    'application': False,
}
