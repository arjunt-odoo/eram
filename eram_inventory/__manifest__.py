{
    'name': 'Hierarchical Product Variants',
    'version': '18.0.1.0.0',
    'category': 'Inventory',
    'summary': 'Infinite nested variants',
    'depends': ['eram_purchase'],
    'data': [
        'security/ir.model.access.csv',
        'views/product_product_views.xml',
        'views/stock_picking_views.xml',
    ],
    'installable': True,
    'auto_install': False,
}