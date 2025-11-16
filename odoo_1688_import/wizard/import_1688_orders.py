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
    currency_id = fields.Many2one('res.currency', string='货币', required=True,
                                   default=lambda self: self._get_default_currency())
    import_summary = fields.Text(string='导入摘要', readonly=True)
    state = fields.Selection([
        ('draft', '草稿'),
        ('done', '完成'),
    ], default='draft', string='状态')

    def _get_default_currency(self):
        """获取默认货币（人民币）"""
        cny = self.env['res.currency'].search([('name', '=', 'CNY')], limit=1)
        if cny:
            return cny.id
        # 如果找不到人民币，使用公司默认货币
        return self.env.company.currency_id.id

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
                    'odoo_product_ref': self._get_cell_value(row[26]),  # Odoo 商品编号 (AA列)
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
                    'currency_id': self.currency_id.id,
                    'date_order': order_data['create_date'] if isinstance(order_data['create_date'], datetime) else fields.Datetime.now(),
                    'origin': f"1688-{order_data['order_no']}",
                    'notes': self._generate_notes(order_data),
                    'order_line': [],
                }

                # 计算运费分摊比例
                shipping_fee = float(order_data['shipping_fee']) if order_data['shipping_fee'] else 0.0
                total_amount = 0.0
                for line_data in order_data['lines']:
                    unit_price = float(line_data['unit_price']) if line_data['unit_price'] else 0.0
                    quantity = float(line_data['quantity']) if line_data['quantity'] else 0.0
                    total_amount += unit_price * quantity

                # 用于记录跳过的产品行
                skipped_lines = []

                # 创建订单行
                for line_data in order_data['lines']:
                    # 根据Odoo商品编号匹配产品
                    product_result = self._find_product_by_reference(line_data)

                    if not product_result['found']:
                        # 产品不存在，记录跳过的行
                        skipped_lines.append({
                            'product_name': line_data['product_name'],
                            'odoo_ref': line_data.get('odoo_product_ref', ''),
                            'reason': product_result['message']
                        })
                        continue

                    product = product_result['product']

                    # 计算含运费的单价
                    unit_price = float(line_data['unit_price']) if line_data['unit_price'] else 0.0
                    quantity = float(line_data['quantity']) if line_data['quantity'] else 0.0

                    # 按比例分摊运费到单价
                    if total_amount > 0 and shipping_fee > 0:
                        line_amount = unit_price * quantity
                        shipping_allocation = (line_amount / total_amount) * shipping_fee
                        unit_price_with_shipping = unit_price + (shipping_allocation / quantity if quantity > 0 else 0)
                    else:
                        unit_price_with_shipping = unit_price

                    # 准备订单行数据
                    line_vals = {
                        'product_id': product.id,
                        'name': line_data['product_name'],
                        'product_qty': quantity,
                        'price_unit': unit_price_with_shipping,
                        'date_planned': fields.Datetime.now(),
                        'taxes_id': [(6, 0, [])],  # 清空税率
                    }
                    po_vals['order_line'].append((0, 0, line_vals))

                # 创建采购订单
                if po_vals['order_line']:
                    purchase_order = PurchaseOrder.create(po_vals)
                    result_data = {
                        'order_no': order_data['order_no'],
                        'po_id': purchase_order.id,
                        'po_name': purchase_order.name,
                        'partner_name': partner.name,
                        'amount_total': purchase_order.amount_total,
                        'status': 'success',
                    }
                    # 如果有跳过的产品行，添加警告信息
                    if skipped_lines:
                        result_data['skipped_lines'] = skipped_lines
                        result_data['status'] = 'partial'  # 部分成功
                    created_orders.append(result_data)
                    _logger.info(f"成功创建采购订单: {purchase_order.name} (1688订单号: {order_data['order_no']})")
                    if skipped_lines:
                        _logger.warning(f"订单 {order_data['order_no']} 有 {len(skipped_lines)} 个产品行被跳过")
                else:
                    _logger.warning(f"订单 {order_data['order_no']} 没有有效的订单行，跳过")
                    created_orders.append({
                        'order_no': order_data['order_no'],
                        'status': 'skipped',
                        'message': '没有有效的订单行' + (f"，{len(skipped_lines)}个产品缺少Odoo商品编号" if skipped_lines else ''),
                        'skipped_lines': skipped_lines if skipped_lines else None,
                    })

            except Exception as e:
                _logger.error(f"创建采购订单时出错 (1688订单号: {order_data['order_no']}): {str(e)}", exc_info=True)
                created_orders.append({
                    'order_no': order_data['order_no'],
                    'status': 'error',
                    'message': str(e),
                })

        return created_orders

    def _find_product_by_reference(self, line_data):
        """
        根据Odoo商品编号查找产品
        返回: {'found': bool, 'product': product对象或None, 'message': str}
        """
        Product = self.env['product.product']

        odoo_ref = line_data.get('odoo_product_ref', '')

        # 如果Odoo商品编号为空，返回错误
        if not odoo_ref:
            return {
                'found': False,
                'product': None,
                'message': 'Excel中Odoo商品编号(AA列)为空，需要手动关联产品'
            }

        # 根据Internal Reference (default_code)查找产品
        product = Product.search([
            ('default_code', '=', odoo_ref)
        ], limit=1)

        if product:
            return {
                'found': True,
                'product': product,
                'message': ''
            }
        else:
            return {
                'found': False,
                'product': None,
                'message': f'在Odoo系统中未找到Internal Reference为"{odoo_ref}"的产品，请先创建该产品或检查编号'
            }

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
        partial = len([o for o in created_orders if o['status'] == 'partial'])
        skipped = len([o for o in created_orders if o['status'] == 'skipped'])
        failed = len([o for o in created_orders if o['status'] == 'error'])

        summary = f"导入完成！\n\n"
        summary += f"总计: {total} 个订单\n"
        summary += f"完全成功: {success} 个\n"
        summary += f"部分成功: {partial} 个（有产品行被跳过）\n"
        summary += f"完全跳过: {skipped} 个\n"
        summary += f"失败: {failed} 个\n\n"

        # 显示完全成功的订单
        if success > 0:
            summary += "完全成功导入的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'success':
                    summary += f"  • {order['po_name']} - {order['partner_name']} - ¥{order['amount_total']:.2f}\n"
                    summary += f"    (1688订单号: {order['order_no']})\n"

        # 显示部分成功的订单（有跳过的产品行）
        if partial > 0:
            summary += "\n部分成功的订单（含跳过的产品行）:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'partial':
                    summary += f"  • {order['po_name']} - {order['partner_name']} - ¥{order['amount_total']:.2f}\n"
                    summary += f"    (1688订单号: {order['order_no']})\n"
                    if 'skipped_lines' in order and order['skipped_lines']:
                        summary += f"    ⚠️  跳过了 {len(order['skipped_lines'])} 个产品行:\n"
                        for skipped_line in order['skipped_lines']:
                            summary += f"       - {skipped_line['product_name']}\n"
                            summary += f"         原因: {skipped_line['reason']}\n"

        if failed > 0:
            summary += "\n失败的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'error':
                    summary += f"  • 1688订单号: {order['order_no']}\n"
                    summary += f"    错误: {order['message']}\n"

        if skipped > 0:
            summary += "\n完全跳过的订单:\n"
            summary += "-" * 80 + "\n"
            for order in created_orders:
                if order['status'] == 'skipped':
                    summary += f"  • 1688订单号: {order['order_no']}\n"
                    summary += f"    原因: {order['message']}\n"
                    if 'skipped_lines' in order and order['skipped_lines']:
                        summary += f"    跳过的产品行:\n"
                        for skipped_line in order['skipped_lines']:
                            summary += f"       - {skipped_line['product_name']}\n"
                            summary += f"         原因: {skipped_line['reason']}\n"

        # 添加提示信息
        if partial > 0 or skipped > 0:
            summary += "\n" + "=" * 80 + "\n"
            summary += "⚠️  重要提示：\n"
            summary += "有产品行因为缺少Odoo商品编号或编号不存在而被跳过。\n"
            summary += "请在Excel文件的AA列（Odoo 商品编号）中填写正确的产品Internal Reference，\n"
            summary += "或者在Odoo系统中创建对应的产品后重新导入。\n"

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
