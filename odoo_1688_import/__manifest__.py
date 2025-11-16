# -*- coding: utf-8 -*-
{
    'name': '1688采购订单导入',
    'version': '16.0.1.2.0',
    'category': 'Purchases',
    'summary': '从1688 Excel文件导入采购订单到Odoo',
    'description': """
        1688采购订单导入模块
        ==================

        功能特性：
        ---------
        * 从1688导出的Excel文件导入采购订单
        * 自动创建供应商（如果不存在）
        * 自动创建产品（如果不存在）
        * 支持一个订单多个产品行
        * 支持选择采购订单货币，默认为人民币(CNY)
        * 运费自动添加为单独订单行
        * 直接使用Excel价格数据，无需二次计算
        * 默认不添加税率，避免价格重复计算

        使用方法：
        ---------
        1. 从1688导出订单Excel文件
        2. 进入 采购 > 1688订单导入
        3. 上传Excel文件，选择货币（默认CNY）
        4. 点击导入
    """,
    'author': 'Your Company',
    'website': 'https://www.yourcompany.com',
    'depends': ['base', 'purchase', 'product'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/import_1688_orders_views.xml',
        'views/menu_views.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'LGPL-3',
}
