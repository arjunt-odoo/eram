{
    'name': 'Hierarchical Product Variants',
    'version': '18.0.1.0.0',
    'category': 'Inventory',
    'summary': 'Infinite nested variants',
    'depends': ['eram_purchase'],
    'data': [
        'security/ir.model.access.csv',
        'views/eram_grn_views.xml',
        'views/product_product_views.xml',
        'views/stock_picking_views.xml',
        'report/eram_grn_report_templates.xml',
        'report/eram_grn_reports.xml',
    ],
    'installable': True,
    'auto_install': False,
}