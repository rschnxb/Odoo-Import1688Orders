# -*- coding: utf-8 -*-
{
    'name': '1688采购订单导入',
    'version': '16.0.2.0.0',
    'category': 'Purchases',
    'summary': '从1688 Excel文件导入采购订单到Odoo，支持Odoo商品编号匹配和运费分摊',
    'description': """
        1688采购订单导入模块
        ==================

        功能特性：
        ---------
        * 从1688导出的Excel文件导入采购订单
        * 自动创建供应商（如果不存在）
        * 使用"Odoo 商品编号"(AA列)精确匹配产品
        * 产品不存在时智能跳过并提示，不自动创建
        * 支持一个订单多个产品行
        * 支持选择采购订单货币，默认为人民币(CNY)
        * 运费按比例智能分摊到产品单价中
        * 直接使用Excel价格数据，无需二次计算
        * 默认不添加税率，避免价格重复计算
        * 详细的导入摘要，包含跳过产品行明细

        重要变更(v16.0.2.0.0)：
        ---------------------
        * AA列"Odoo 商品编号"为必填项
        * 不再自动创建产品，必须预先在Odoo中创建
        * 运费不再作为单独行，而是分摊到产品单价

        使用方法：
        ---------
        1. 在Odoo中预先创建好产品，设置Internal Reference
        2. 从1688导出订单Excel文件
        3. 在Excel的AA列填写Odoo产品的Internal Reference
        4. 进入 采购 > 1688订单导入
        5. 上传Excel文件，选择货币（默认CNY）
        6. 点击导入，查看导入摘要
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
