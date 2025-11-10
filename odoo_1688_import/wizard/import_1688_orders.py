# -*- coding: utf-8 -*-

import base64
import io
import logging
from datetime import datetime
from odoo import models, fields, api, _
from odoo.exceptions import UserError

try:
    import openpyxl
except ImportError:
    openpyxl = None

_logger = logging.getLogger(__name__)


class Import1688Orders(models.TransientModel):
    _name = 'import.1688.orders'
    _description = '导入1688采购订单'

    excel_file = fields.Binary(string='Excel文件', required=True)
    file_name = fields.Char(string='文件名')
    import_summary = fields.Text(string='导入摘要', readonly=True)
    state = fields.Selection([
        ('draft', '草稿'),
        ('done', '完成'),
    ], default='draft', string='状态')

    def action_import(self):
        """导入1688订单"""
        self.ensure_one()

        if not openpyxl:
            raise UserError(_('请先安装openpyxl库：pip install openpyxl'))

        if not self.excel_file:
            raise UserError(_('请先上传Excel文件'))

        try:
            # 解码Excel文件
            file_data = base64.b64decode(self.excel_file)
            workbook = openpyxl.load_workbook(io.BytesIO(file_data))
            sheet = workbook.active

            # 解析订单数据
            orders_data = self._parse_excel_data(sheet)

            # 创建采购订单
            created_orders = self._create_purchase_orders(orders_data)

            # 生成导入摘要
            summary = self._generate_summary(created_orders)

            self.write({
                'import_summary': summary,
                'state': 'done',
            })

            return {
                'type': 'ir.actions.act_window',
                'res_model': 'import.1688.orders',
                'view_mode': 'form',
                'res_id': self.id,
                'target': 'new',
            }

        except Exception as e:
            _logger.error(f'导入1688订单时出错: {str(e)}', exc_info=True)
            raise UserError(_('导入失败：%s') % str(e))

    def _parse_excel_data(self, sheet):
        """解析Excel数据"""
        orders_dict = {}

        # 从第2行开始读取（第1行是表头）
        for row_idx in range(2, sheet.max_row + 1):
            row = list(sheet[row_idx])

            # 获取订单编号（第1列）
            order_no = self._get_cell_value(row[0])

            # 如果订单编号为空，说明是同一订单的额外产品行
            if not order_no and orders_dict:
                # 使用上一个订单编号
                order_no = last_order_no
            else:
                last_order_no = order_no

            if not order_no:
                continue

            # 初始化订单字典
            if order_no not in orders_dict:
                orders_dict[order_no] = {
                    'order_no': order_no,
                    'seller_company': self._get_cell_value(row[3]),  # 卖家公司名
                    'seller_name': self._get_cell_value(row[4]),     # 卖家会员名
                    'total_price': self._get_cell_value(row[5]),     # 货品总价
                    'shipping_fee': self._get_cell_value(row[6]),    # 运费
                    'discount': self._get_cell_value(row[7]),        # 涨价或折扣
                    'actual_payment': self._get_cell_value(row[8]),  # 实付款
                    'order_status': self._get_cell_value(row[9]),    # 订单状态
                    'create_date': self._get_cell_value(row[10]),    # 订单创建时间
                    'payment_date': self._get_cell_value(row[11]),   # 订单付款时间
                    'tracking_no': self._get_cell_value(row[31]),    # 运单号
                    'logistics_company': self._get_cell_value(row[30]),  # 物流公司
                    'buyer_note': self._get_cell_value(row[29]),     # 买家留言
                    'lines': []
                }

            # 添加订单行
            product_name = self._get_cell_value(row[18])  # 货品标题
            if product_name:
                line_data = {
                    'product_name': product_name,
                    'unit_price': self._get_cell_value(row[19]),     # 单价
                    'quantity': self._get_cell_value(row[20]),       # 数量
                    'uom': self._get_cell_value(row[21]),           # 单位
                    'product_code': self._get_cell_value(row[22]),   # 货号
                    'model': self._get_cell_value(row[23]),         # 型号
                    'sku_id': self._get_cell_value(row[25]),        # SKU ID
                }
                orders_dict[order_no]['lines'].append(line_data)

        return list(orders_dict.values())

    def _get_cell_value(self, cell):
        """获取单元格值"""
        if cell.value is None:
            return False
        if isinstance(cell.value, datetime):
            return cell.value
        return str(cell.value).strip() if cell.value else False

    def _create_purchase_orders(self, orders_data):
        """创建采购订单"""
        PurchaseOrder = self.env['purchase.order']
        created_orders = []

        for order_data in orders_data:
            try:
                # 获取或创建供应商
                partner = self._get_or_create_partner(
                    order_data['seller_company'],
                    order_data['seller_name']
                )

                # 准备采购订单数据
                po_vals = {
                    'partner_id': partner.id,
                    'date_order': order_data['create_date'] if isinstance(order_data['create_date'], datetime) else fields.Datetime.now(),
                    'origin': f"1688-{order_data['order_no']}",
                    'notes': self._generate_notes(order_data),
                    'order_line': [],
                }

                # 创建订单行
                for line_data in order_data['lines']:
                    # 获取或创建产品
                    product = self._get_or_create_product(line_data)

                    # 准备订单行数据
                    line_vals = {
                        'product_id': product.id,
                        'name': line_data['product_name'],
                        'product_qty': float(line_data['quantity']) if line_data['quantity'] else 1.0,
                        'price_unit': float(line_data['unit_price']) if line_data['unit_price'] else 0.0,
                        'date_planned': fields.Datetime.now(),
                    }
                    po_vals['order_line'].append((0, 0, line_vals))

                # 创建采购订单
                if po_vals['order_line']:
                    purchase_order = PurchaseOrder.create(po_vals)
                    created_orders.append({
                        'order_no': order_data['order_no'],
                        'po_id': purchase_order.id,
                        'po_name': purchase_order.name,
                        'partner_name': partner.name,
                        'amount_total': purchase_order.amount_total,
                        'status': 'success',
                    })
                    _logger.info(f"成功创建采购订单: {purchase_order.name} (1688订单号: {order_data['order_no']})")
                else:
                    _logger.warning(f"订单 {order_data['order_no']} 没有有效的订单行，跳过")
                    created_orders.append({
                        'order_no': order_data['order_no'],
                        'status': 'skipped',
                        'message': '没有有效的订单行',
                    })

            except Exception as e:
                _logger.error(f"创建采购订单时出错 (1688订单号: {order_data['order_no']}): {str(e)}", exc_info=True)
                created_orders.append({
                    'order_no': order_data['order_no'],
                    'status': 'error',
                    'message': str(e),
                })

        return created_orders

    def _get_or_create_partner(self, company_name, member_name):
        """获取或创建供应商"""
        Partner = self.env['res.partner']

        # 首先尝试根据公司名称查找
        partner = False
        if company_name:
            partner = Partner.search([
                ('name', '=', company_name),
                ('supplier_rank', '>', 0)
            ], limit=1)

        # 如果找不到，创建新供应商
        if not partner:
            partner_vals = {
                'name': company_name or member_name or '未知供应商',
                'supplier_rank': 1,
                'customer_rank': 0,
                'comment': f'从1688自动导入\n会员名: {member_name or "无"}',
            }
            partner = Partner.create(partner_vals)
            _logger.info(f"创建新供应商: {partner.name}")

        return partner

    def _get_or_create_product(self, line_data):
        """获取或创建产品"""
        Product = self.env['product.product']

        # 尝试根据产品编码查找
        product = False
        if line_data['product_code']:
            product = Product.search([
                ('default_code', '=', line_data['product_code'])
            ], limit=1)

        # 如果找不到，尝试根据SKU ID查找
        if not product and line_data['sku_id']:
            product = Product.search([
                ('default_code', '=', line_data['sku_id'])
            ], limit=1)

        # 如果还是找不到，创建新产品
        if not product:
            # 生成产品编码
            default_code = line_data['product_code'] or line_data['sku_id'] or f"1688-{self.env['ir.sequence'].next_by_code('product.product') or '00000'}"

            product_vals = {
                'name': line_data['product_name'],
                'default_code': default_code,
                'type': 'product',
                'purchase_ok': True,
                'sale_ok': True,
                'description_purchase': f"型号: {line_data['model'] or '无'}\nSKU ID: {line_data['sku_id'] or '无'}",
            }
            product = Product.create(product_vals)
            _logger.info(f"创建新产品: {product.name} [{product.default_code}]")

        return product

    def _generate_notes(self, order_data):
        """生成订单备注"""
        notes = f"1688订单信息\n"
        notes += f"=" * 50 + "\n"
        notes += f"订单编号: {order_data['order_no']}\n"
        notes += f"卖家会员名: {order_data['seller_name'] or '无'}\n"
        notes += f"订单状态: {order_data['order_status'] or '无'}\n"

        if order_data['payment_date']:
            notes += f"付款时间: {order_data['payment_date']}\n"

        if order_data['logistics_company']:
            notes += f"物流公司: {order_data['logistics_company']}\n"

        if order_data['tracking_no']:
            notes += f"运单号: {order_data['tracking_no']}\n"

        if order_data['buyer_note']:
            notes += f"买家留言: {order_data['buyer_note']}\n"

        if order_data['shipping_fee']:
            notes += f"运费: ¥{order_data['shipping_fee']}\n"

        if order_data['discount']:
            notes += f"折扣: ¥{order_data['discount']}\n"

        if order_data['actual_payment']:
            notes += f"实付款: ¥{order_data['actual_payment']}\n"

        return notes

    def _generate_summary(self, created_orders):
        """生成导入摘要"""
        total = len(created_orders)
        success = len([o for o in created_orders if o['status'] == 'success'])
        skipped = len([o for o in created_orders if o['status'] == 'skipped'])
        failed = len([o for o in created_orders if o['status'] == 'error'])

        summary = f"导入完成！\n\n"
        summary += f"总计: {total} 个订单\n"
        summary += f"成功: {success} 个\n"
        summary += f"跳过: {skipped} 个\n"
        summary += f"失败: {failed} 个\n\n"

        if success > 0:
            summary += "成功导入的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'success':
                    summary += f"  • {order['po_name']} - {order['partner_name']} - ¥{order['amount_total']:.2f}\n"
                    summary += f"    (1688订单号: {order['order_no']})\n"

        if failed > 0:
            summary += "\n失败的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'error':
                    summary += f"  • 1688订单号: {order['order_no']}\n"
                    summary += f"    错误: {order['message']}\n"

        if skipped > 0:
            summary += "\n跳过的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'skipped':
                    summary += f"  • 1688订单号: {order['order_no']}\n"
                    summary += f"    原因: {order['message']}\n"

        return summary

    def action_view_orders(self):
        """查看创建的采购订单"""
        self.ensure_one()

        return {
            'type': 'ir.actions.act_window',
            'name': '导入的采购订单',
            'res_model': 'purchase.order',
            'view_mode': 'tree,form',
            'domain': [('origin', 'like', '1688-')],
            'context': {'search_default_filter_to_approve': 1},
        }
