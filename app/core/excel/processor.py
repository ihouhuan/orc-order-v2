"""
Excel处理核心模块
--------------
提供Excel文件处理功能，包括表格解析、数据提取和处理。
"""

import os
import re
import pandas as pd
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy as xlcopy
from typing import Dict, List, Optional, Tuple, Union, Any
from datetime import datetime

from ...config.settings import ConfigManager
from ..utils.log_utils import get_logger
from ..utils.file_utils import (
    ensure_dir,
    get_file_extension,
    get_latest_file,
    load_json,
    save_json
)
from ..utils.string_utils import (
    clean_string,
    clean_barcode,
    extract_number,
    format_barcode
)
from .converter import UnitConverter

logger = get_logger(__name__)

class ExcelProcessor:
    """
    Excel处理器：处理OCR识别后的Excel文件，
    提取条码、单价和数量，并按照采购单模板的格式填充
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化Excel处理器
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        logger.info("初始化ExcelProcessor")
        self.config = config or ConfigManager()
        
        # 获取配置
        self.output_dir = self.config.get_path('Paths', 'output_folder', 'data/output', create=True)
        self.temp_dir = self.config.get_path('Paths', 'temp_folder', 'data/temp', create=True)
        
        # 获取模板文件路径
        template_folder = self.config.get('Paths', 'template_folder', 'templates')
        template_name = self.config.get('Templates', 'purchase_order', '银豹-采购单模板.xls')
        
        self.template_path = os.path.join(template_folder, template_name)
        
        # 检查模板文件是否存在
        if not os.path.exists(self.template_path):
            logger.error(f"模板文件不存在: {self.template_path}")
            raise FileNotFoundError(f"模板文件不存在: {self.template_path}")
        
        # 用于记录已处理的文件
        self.cache_file = os.path.join(self.output_dir, "processed_files.json")
        self.processed_files = self._load_processed_files()
        
        # 创建单位转换器
        self.unit_converter = UnitConverter()
        
        logger.info(f"初始化完成，模板文件: {self.template_path}")
    
    def _load_processed_files(self) -> Dict[str, str]:
        """
        加载已处理文件的缓存
        
        Returns:
            处理记录字典
        """
        return load_json(self.cache_file, {})
        
    def _save_processed_files(self) -> None:
        """保存已处理文件的缓存"""
        save_json(self.processed_files, self.cache_file)
    
    def get_latest_excel(self) -> Optional[str]:
        """
        获取output目录下最新的Excel文件（排除采购单文件）
        
        Returns:
            最新Excel文件的路径，如果未找到则返回None
        """
        logger.info(f"搜索目录 {self.output_dir} 中的Excel文件")
        
        # 使用文件工具获取最新文件
        latest_file = get_latest_file(
            self.output_dir,
            pattern="",  # 不限制文件名
            extensions=['.xlsx', '.xls']  # 限制为Excel文件
        )
        
        # 如果没有找到文件
        if not latest_file:
            logger.warning(f"未在 {self.output_dir} 目录下找到未处理的Excel文件")
            return None
        
        # 检查是否是采购单（以"采购单_"开头的文件）
        file_name = os.path.basename(latest_file)
        if file_name.startswith('采购单_'):
            logger.warning(f"找到的最新文件是采购单，不作处理: {latest_file}")
            return None
        
        logger.info(f"找到最新的Excel文件: {latest_file}")
        return latest_file
    
    def validate_barcode(self, barcode: Any) -> bool:
        """
        验证条码是否有效
        新增功能：如果条码是"仓库"，则返回False以避免误认为有效条码
        
        Args:
            barcode: 条码值
            
        Returns:
            条码是否有效
        """
        # 处理"仓库"特殊情况
        if isinstance(barcode, str) and barcode.strip() in ["仓库", "仓库全名"]:
            logger.warning(f"条码为仓库标识: {barcode}")
            return False
            
        # 清理条码格式
        barcode_clean = clean_barcode(barcode)
        
        # 对特定的错误条码进行修正（开头改6开头）
        if len(barcode_clean) > 8 and barcode_clean.startswith('5') and not barcode_clean.startswith('53'):
            barcode_clean = '6' + barcode_clean[1:]
            logger.info(f"修正条码前缀 5->6: {barcode} -> {barcode_clean}")
            
        # 验证条码长度
        if len(barcode_clean) < 8 or len(barcode_clean) > 13:
            logger.warning(f"条码长度异常: {barcode_clean}, 长度={len(barcode_clean)}")
            return False
            
        # 验证条码是否全为数字
        if not barcode_clean.isdigit():
            logger.warning(f"条码包含非数字字符: {barcode_clean}")
            return False
            
        # 对于序号9的特殊情况，允许其条码格式
        if barcode_clean == "5321545613":
            logger.info(f"特殊条码验证通过: {barcode_clean}")
            return True
            
        logger.debug(f"条码验证通过: {barcode_clean}")
        return True
    
    def extract_barcode(self, df: pd.DataFrame) -> List[str]:
        """
        从数据帧中提取条码列名
        
        Args:
            df: 数据帧
            
        Returns:
            可能的条码列名列表
        """
        possible_barcode_columns = [
            '条码', '条形码', '商品条码', '商品条形码', 
            '商品编码', '商品编号', '条形码', '条码（必填）', 
            'barcode', 'Barcode', '编码', '条形码'
        ]
        
        found_columns = []
        for col in df.columns:
            col_str = str(col).strip()
            if col_str in possible_barcode_columns:
                found_columns.append(col)
        
        return found_columns
    
    def extract_product_info(self, df: pd.DataFrame) -> List[Dict]:
        """
        从数据帧中提取商品信息
        
        Args:
            df: 数据帧
            
        Returns:
            商品信息列表
        """
        # 提取有用的列
        barcode_cols = self.extract_barcode(df)
        
        # 如果没有找到条码列，无法继续处理
        if not barcode_cols:
            logger.error("未找到条码列，无法处理")
            return []
            
        # 定义列名映射
        column_mapping = {
            'name': ['商品名称', '名称', '品名', '商品', '商品名', '商品或服务名称', '品项名'],
            'specification': ['规格', '规格型号', '型号', '商品规格'],
            'quantity': ['数量', '采购数量', '购买数量', '采购数量', '订单数量', '数量（必填）'],
            'unit': ['单位', '采购单位', '计量单位', '单位（必填）'],
            'price': ['单价', '价格', '采购单价', '销售价', '进货价', '单价（必填）']
        }
        
        # 映射列名到标准名称
        mapped_columns = {'barcode': barcode_cols[0]}  # 使用第一个找到的条码列
        
        for target, possible_names in column_mapping.items():
            for col in df.columns:
                col_str = str(col).strip()
                for name in possible_names:
                    if col_str == name:
                        mapped_columns[target] = col
                        break
                if target in mapped_columns:
                    break
        
        logger.info(f"列名映射结果: {mapped_columns}")
        
        # 提取商品信息
        products = []
        
        for _, row in df.iterrows():
            barcode = row.get(mapped_columns['barcode'])
            
            # 跳过空行或无效条码
            if pd.isna(barcode) or not self.validate_barcode(barcode):
                continue
                
            # 创建商品信息字典
            product = {
                'barcode': format_barcode(barcode),
                'name': row.get(mapped_columns.get('name', ''), ''),
                'specification': row.get(mapped_columns.get('specification', ''), ''),
                'quantity': extract_number(str(row.get(mapped_columns.get('quantity', ''), 0))) or 0,
                'unit': str(row.get(mapped_columns.get('unit', ''), '')),
                'price': extract_number(str(row.get(mapped_columns.get('price', ''), 0))) or 0
            }
            
            # 如果商品名称为空但商品条码不为空，则使用条码作为名称
            if not product['name'] and product['barcode']:
                product['name'] = f"商品 ({product['barcode']})"
            
            # 推断规格
            if not product['specification'] and product['name']:
                inferred_spec = self.unit_converter.infer_specification_from_name(product['name'])
                if inferred_spec:
                    product['specification'] = inferred_spec
                    logger.info(f"从商品名称推断规格: {product['name']} -> {inferred_spec}")
            
            # 单位处理：如果单位为空但数量包含单位信息
            quantity_str = str(row.get(mapped_columns.get('quantity', ''), ''))
            if not product['unit'] and '数量' in mapped_columns:
                num, unit = self.unit_converter.extract_unit_from_quantity(quantity_str)
                if unit:
                    product['unit'] = unit
                    logger.info(f"从数量提取单位: {quantity_str} -> {unit}")
                    # 如果数量被提取出来，更新数量
                    if num is not None:
                        product['quantity'] = num
            
            # 应用单位转换规则
            product = self.unit_converter.process_unit_conversion(product)
            
            products.append(product)
        
        logger.info(f"提取到 {len(products)} 个商品信息")
        return products
    
    def fill_template(self, products: List[Dict], output_file_path: str) -> bool:
        """
        填充采购单模板
        
        Args:
            products: 商品信息列表
            output_file_path: 输出文件路径
            
        Returns:
            是否成功填充
        """
        try:
            # 打开模板文件
            template_workbook = xlrd.open_workbook(self.template_path, formatting_info=True)
            template_sheet = template_workbook.sheet_by_index(0)
            
            # 创建可写的副本
            output_workbook = xlcopy(template_workbook)
            output_sheet = output_workbook.get_sheet(0)
            
            # 先对产品按条码分组，区分正常商品和赠品
            barcode_groups = {}
            
            # 遍历所有产品，按条码分组
            logger.info(f"开始处理{len(products)} 个产品信息")
            for product in products:
                barcode = product.get('barcode', '')
                if not barcode:
                    logger.warning(f"跳过无条码商品")
                    continue
                
                # 获取数量和单价
                quantity = product.get('quantity', 0)
                price = product.get('price', 0)
                
                # 判断是否为赠品（价格为0）
                is_gift = price == 0
                
                logger.info(f"处理商品: 条码={barcode}, 数量={quantity}, 单价={price}, 是否赠品={is_gift}")
                
                if barcode not in barcode_groups:
                    barcode_groups[barcode] = {
                        'normal': None,  # 正常商品信息
                        'gift_quantity': 0  # 赠品数量
                    }
                
                if is_gift:
                    # 是赠品，累加赠品数量
                    barcode_groups[barcode]['gift_quantity'] += quantity
                    logger.info(f"发现赠品：条码{barcode}, 数量={quantity}")
                else:
                    # 是正常商品
                    if barcode_groups[barcode]['normal'] is None:
                        barcode_groups[barcode]['normal'] = {
                            'product': product,
                            'quantity': quantity,
                            'price': price
                        }
                        logger.info(f"发现正常商品：条码{barcode}, 数量={quantity}, 单价={price}")
                    else:
                        # 如果有多个正常商品记录，累加数量
                        barcode_groups[barcode]['normal']['quantity'] += quantity
                        logger.info(f"累加正常商品数量：条码{barcode}, 新增={quantity}, 累计={barcode_groups[barcode]['normal']['quantity']}")
                        
                        # 如果单价不同，取平均值
                        if price != barcode_groups[barcode]['normal']['price']:
                            avg_price = (barcode_groups[barcode]['normal']['price'] + price) / 2
                            barcode_groups[barcode]['normal']['price'] = avg_price
                            logger.info(f"调整单价(取平均值)：条码{barcode}, 原价={barcode_groups[barcode]['normal']['price']}, 新价={price}, 平均={avg_price}")
            
            # 输出调试信息
            logger.info(f"分组后共{len(barcode_groups)} 个不同条码的商品")
            for barcode, group in barcode_groups.items():
                if group['normal'] is not None:
                    logger.info(f"条码 {barcode} 处理结果：正常商品数量{group['normal']['quantity']}，单价{group['normal']['price']}，赠品数量{group['gift_quantity']}")
                else:
                    logger.info(f"条码 {barcode} 处理结果：只有赠品，数量={group['gift_quantity']}")
            
            # 准备填充数据
            row_index = 1  # 从第2行开始填充（索引从0开始）
            
            for barcode, group in barcode_groups.items():
                # 1. 列B(1): 条码（必填）
                output_sheet.write(row_index, 1, barcode)
                
                if group['normal'] is not None:
                    # 有正常商品
                    product = group['normal']['product']
                    
                    # 2. 列C(2): 采购量（必填） 使用正常商品的采购量
                    normal_quantity = group['normal']['quantity']
                    output_sheet.write(row_index, 2, normal_quantity)
                    
                    # 3. 列D(3): 赠送量 - 添加赠品数量
                    if group['gift_quantity'] > 0:
                        output_sheet.write(row_index, 3, group['gift_quantity'])
                        logger.info(f"条码 {barcode} 填充：采购量={normal_quantity}，赠品数量{group['gift_quantity']}")
                    
                    # 4. 列E(4): 采购单价（必填）
                    purchase_price = group['normal']['price']
                    style = xlwt.XFStyle()
                    style.num_format_str = '0.0000'
                    output_sheet.write(row_index, 4, round(purchase_price, 4), style)
                else:
                    # 只有赠品，没有正常商品
                    # 采购量填0，赠送量填赠品数量
                    output_sheet.write(row_index, 2, 0)  # 采购量为0
                    output_sheet.write(row_index, 3, group['gift_quantity'])  # 赠送量
                    output_sheet.write(row_index, 4, 0)  # 单价为0
                    
                    logger.info(f"条码 {barcode} 填充：仅有赠品，采购量=0，赠品数量={group['gift_quantity']}")
                
                # 移到下一行
                row_index += 1
            
            # 保存文件
            output_workbook.save(output_file_path)
            logger.info(f"采购单已保存到: {output_file_path}")
            return True
            
        except Exception as e:
            logger.error(f"填充模板时出错: {e}")
            return False
    
    def process_specific_file(self, file_path: str) -> Optional[str]:
        """
        处理指定的Excel文件
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            输出文件路径，如果处理失败则返回None
        """
        logger.info(f"开始处理Excel文件: {file_path}")
        
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return None
        
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            logger.info(f"成功读取Excel文件: {file_path}, 共 {len(df)} 行")
            
            # 提取商品信息
            products = self.extract_product_info(df)
            
            if not products:
                logger.warning("未提取到有效商品信息")
                return None
            
            # 生成输出文件名
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_dir, f"采购单_{file_name}.xls")
            
            # 填充模板并保存
            if self.fill_template(products, output_file):
                # 记录已处理文件
                self.processed_files[file_path] = output_file
                self._save_processed_files()
                return output_file
            
            return None
            
        except Exception as e:
            logger.error(f"处理Excel文件时出错: {file_path}, 错误: {e}")
            return None
    
    def process_latest_file(self) -> Optional[str]:
        """
        处理最新的Excel文件
        
        Returns:
            输出文件路径，如果处理失败则返回None
        """
        # 获取最新的Excel文件
        latest_file = self.get_latest_excel()
        if not latest_file:
            logger.warning("未找到可处理的Excel文件")
            return None
        
        # 处理文件
        return self.process_specific_file(latest_file) 