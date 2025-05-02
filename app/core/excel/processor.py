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
            'barcode', 'Barcode', '编码', '条形码', '电脑条码',
            '条码ID', '产品条码', 'BarCode'
        ]
        
        found_columns = []
        
        # 检查精确匹配
        for col in df.columns:
            col_str = str(col).strip()
            if col_str in possible_barcode_columns:
                found_columns.append(col)
                logger.info(f"找到精确匹配的条码列: {col_str}")
        
        # 如果找不到精确匹配，尝试部分匹配
        if not found_columns:
            for col in df.columns:
                col_str = str(col).strip().lower()
                for keyword in ['条码', '条形码', 'barcode', '编码']:
                    if keyword.lower() in col_str:
                        found_columns.append(col)
                        logger.info(f"找到部分匹配的条码列: {col} (包含关键词: {keyword})")
                        break
        
        # 如果仍然找不到，尝试使用数据特征识别
        if not found_columns and len(df) > 0:
            for col in df.columns:
                # 检查此列数据是否符合条码特征
                sample_values = df[col].dropna().astype(str).tolist()[:10]  # 取前10个非空值
                
                if sample_values and all(len(val) >= 8 and len(val) <= 14 for val in sample_values):
                    # 大多数条码长度在8-14之间
                    if all(val.isdigit() for val in sample_values):
                        found_columns.append(col)
                        logger.info(f"基于数据特征识别的可能条码列: {col}")
        
        return found_columns
    
    def extract_product_info(self, df: pd.DataFrame) -> List[Dict]:
        """
        从处理后的数据框中提取商品信息
        支持处理不同格式的Excel文件
        
        Args:
            df: 数据框
            
        Returns:
            商品信息列表，每个商品为一个字典
        """
        products = []
        
        # 检测表头位置和数据格式
        column_mapping = self._detect_column_mapping(df)
        logger.info(f"列名映射结果: {column_mapping}")
        
        # 检查是否有规格列
        has_specification_column = '规格' in df.columns
        logger.info(f"是否存在规格列: {has_specification_column}")
        
        # 处理每一行数据
        for idx, row in df.iterrows():
            try:
                # 条码处理 - 确保条码总是字符串格式且不带小数点
                barcode_raw = row[column_mapping['barcode']] if column_mapping.get('barcode') else ''
                if pd.isna(barcode_raw) or barcode_raw == '' or str(barcode_raw).strip() in ['nan', 'None']:
                    continue
                
                # 使用format_barcode函数处理条码，确保无小数点
                barcode = format_barcode(barcode_raw)
                
                # 处理数量字段，先提取数字部分再转换为浮点数
                quantity_value = 0
                quantity_str = ""
                if column_mapping.get('quantity') and not pd.isna(row[column_mapping['quantity']]):
                    quantity_str = str(row[column_mapping['quantity']])
                    # 使用提取数字的函数
                    quantity_num = extract_number(quantity_str)
                    if quantity_num is not None:
                        quantity_value = quantity_num
                
                # 基础信息
                product = {
                    'barcode': barcode,
                    'name': str(row[column_mapping['name']]) if column_mapping.get('name') else '',
                    'quantity': quantity_value,
                    'price': float(row[column_mapping['price']]) if column_mapping.get('price') and not pd.isna(row[column_mapping['price']]) else 0,
                    'unit': str(row[column_mapping['unit']]) if column_mapping.get('unit') and not pd.isna(row[column_mapping['unit']]) else '',
                    'specification': '',
                    'package_quantity': None
                }
                
                # 清理单位
                if product['unit'] == 'nan' or product['unit'] == 'None':
                    product['unit'] = ''
                
                # 打印每行提取出的信息
                logger.info(f"第{idx+1}行: 提取商品信息 条码={product['barcode']}, 名称={product['name']}, 规格={product['specification']}, 数量={product['quantity']}, 单位={product['unit']}, 单价={product['price']}")
                
                # 从数量字段中提取单位（如果单位字段为空）
                if not product['unit'] and quantity_str:
                    num, unit = self.unit_converter.extract_unit_from_quantity(quantity_str)
                    if unit:
                        product['unit'] = unit
                        logger.info(f"从数量提取单位: {quantity_str} -> {unit}")
                        # 如果数量被提取出来，更新数量
                        if num is not None:
                            product['quantity'] = num
                
                # 提取规格并解析包装数量
                if '规格' in df.columns and not pd.isna(row['规格']):
                    product['specification'] = str(row['规格'])
                    package_quantity = self.parse_specification(product['specification'])
                    if package_quantity:
                        product['package_quantity'] = package_quantity
                        logger.info(f"解析规格: {product['specification']} -> 包装数量={package_quantity}")
                else:
                    # 逻辑1: 如果规格为空，尝试从商品名称推断规格
                    if product['name']:
                        # 特殊处理："营养快线原味450g*15"或"娃哈哈瓶装大AD水蜜桃450ml*15"等形式的名称
                        weight_volume_pattern = r'.*?\d+(?:g|ml|毫升|克)[*xX×](\d+)'
                        match = re.search(weight_volume_pattern, product['name'])
                        if match:
                            inferred_spec = f"1*{match.group(1)}"
                            inferred_qty = int(match.group(1))
                            product['specification'] = inferred_spec
                            product['package_quantity'] = inferred_qty
                            logger.info(f"从商品名称提取重量/容量规格: {product['name']} -> {inferred_spec}, 包装数量={inferred_qty}")
                        else:
                            # 一般情况的规格推断
                            inferred_spec, inferred_qty = self.infer_specification_from_name(product['name'])
                            if inferred_spec:
                                product['specification'] = inferred_spec
                                product['package_quantity'] = inferred_qty
                                logger.info(f"从商品名称推断规格: {product['name']} -> {inferred_spec}, 包装数量={inferred_qty}")
                
                # 应用单位转换规则
                product = self.unit_converter.process_unit_conversion(product)
                
                products.append(product)
            except Exception as e:
                logger.error(f"提取第{idx+1}行商品信息时出错: {e}", exc_info=True)
                continue
                
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
                # 确保条码是整数字符串
                barcode = format_barcode(barcode)
                
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
    
    def _find_header_row(self, df: pd.DataFrame) -> Optional[int]:
        """
        自动识别表头行
        
        通过多种规则识别表头：
        1. 检查行是否包含典型的表头关键词（条码、商品名称、数量等）
        2. 检查是否是第一个非空行
        3. 检查行是否有较多的字符串类型单元格（表头通常是字符串）
        
        Args:
            df: 数据帧
            
        Returns:
            表头行索引，如果未找到则返回None
        """
        # 定义可能的表头关键词
        header_keywords = [
            '条码', '条形码', '商品条码', '商品名称', '名称', '数量', '单位', '单价', 
            '规格', '商品编码', '采购数量', '采购单位', '商品', '品名'
        ]
        
        # 存储每行的匹配分数
        row_scores = []
        
        # 遍历前10行（通常表头不会太靠后）
        max_rows_to_check = min(10, len(df))
        for row in range(max_rows_to_check):
            row_data = df.iloc[row]
            score = 0
            
            # 检查1: 关键词匹配
            for cell in row_data:
                if isinstance(cell, str):
                    cell_clean = str(cell).strip().lower()
                    for keyword in header_keywords:
                        if keyword.lower() in cell_clean:
                            score += 5  # 每匹配一个关键词加5分
            
            # 检查2: 非空单元格比例
            non_empty_cells = row_data.count()
            if non_empty_cells / len(row_data) > 0.5:  # 如果超过一半的单元格有内容
                score += 2
            
            # 检查3: 字符串类型单元格比例
            string_cells = sum(1 for cell in row_data if isinstance(cell, str))
            if string_cells / len(row_data) > 0.5:  # 如果超过一半的单元格是字符串
                score += 3
                
            row_scores.append((row, score))
            
            # 日志记录每行的评分情况
            logger.debug(f"第{row+1}行评分: {score}，内容: {row_data.values}")
        
        # 按评分排序
        row_scores.sort(key=lambda x: x[1], reverse=True)
        
        # 如果最高分达到一定阈值，认为是表头
        if row_scores and row_scores[0][1] >= 5:
            best_row = row_scores[0][0]
            logger.info(f"找到可能的表头行: 第{best_row+1}行，评分: {row_scores[0][1]}")
            return best_row
        
        # 如果没有找到明确的表头，尝试找第一个非空行
        for row in range(len(df)):
            if df.iloc[row].notna().sum() > 3:  # 至少有3个非空单元格
                logger.info(f"未找到明确表头，使用第一个有效行: 第{row+1}行")
                return row
                
        logger.warning("无法识别表头行")
        return None
    
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
            # 读取Excel文件时不立即指定表头
            df = pd.read_excel(file_path, header=None)
            logger.info(f"成功读取Excel文件: {file_path}, 共 {len(df)} 行")
            
            # 自动识别表头行
            header_row = self._find_header_row(df)
            if header_row is None:
                logger.error("无法识别表头行")
                return None
                
            logger.info(f"识别到表头在第 {header_row+1} 行")
            
            # 重新读取Excel，正确指定表头行
            df = pd.read_excel(file_path, header=header_row)
            logger.info(f"使用表头行重新读取数据，共 {len(df)} 行有效数据")
            
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
                
                # 不再自动打开输出目录
                logger.info(f"采购单已保存到: {output_file}")
                
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
    
    def _detect_column_mapping(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        检测和映射Excel表头列名
        
        Args:
            df: 数据框
            
        Returns:
            列名映射字典，键为标准列名，值为实际列名
        """
        # 提取有用的列
        barcode_cols = self.extract_barcode(df)
        
        # 如果没有找到条码列，无法继续处理
        if not barcode_cols:
            logger.error("未找到条码列，无法处理")
            return {}
        
        # 定义列名映射
        column_mapping = {
            'name': ['商品名称', '名称', '品名', '商品', '商品名', '商品或服务名称', '品项名', '产品名称', '品项'],
            'specification': ['规格', '规格型号', '型号', '商品规格', '产品规格', '包装规格'],
            'quantity': ['数量', '采购数量', '购买数量', '采购数量', '订单数量', '数量（必填）', '入库数', '入库数量'],
            'unit': ['单位', '采购单位', '计量单位', '单位（必填）', '单位名称', '计价单位'],
            'price': ['单价', '价格', '采购单价', '销售价', '进货价', '单价（必填）', '采购价', '参考价', '入库单价']
        }
        
        # 映射列名到标准名称
        mapped_columns = {'barcode': barcode_cols[0]}  # 使用第一个找到的条码列
        
        # 记录列名映射详情
        logger.info(f"使用条码列: {mapped_columns['barcode']}")
        
        for target, possible_names in column_mapping.items():
            for col in df.columns:
                col_str = str(col).strip()
                for name in possible_names:
                    if col_str == name:
                        mapped_columns[target] = col
                        logger.info(f"找到{target}列: {col}")
                        break
                if target in mapped_columns:
                    break
            
            # 如果没有找到精确匹配，尝试部分匹配
            if target not in mapped_columns:
                for col in df.columns:
                    col_str = str(col).strip().lower()
                    for name in possible_names:
                        if name.lower() in col_str:
                            mapped_columns[target] = col
                            logger.info(f"找到{target}列(部分匹配): {col}")
                            break
                    if target in mapped_columns:
                        break
        
        return mapped_columns 
    
    def infer_specification_from_name(self, product_name: str) -> Tuple[Optional[str], Optional[int]]:
        """
        从商品名称推断规格
        根据特定的命名规则匹配规格信息
        
        Args:
            product_name: 商品名称
            
        Returns:
            规格字符串和包装数量的元组
        """
        if not product_name or not isinstance(product_name, str):
            logger.warning(f"无效的商品名: {product_name}")
            return None, None
            
        product_name = product_name.strip()
        
        # 特殊处理：重量/容量*数字格式
        weight_volume_pattern = r'.*?\d+(?:g|ml|毫升|克)[*xX×](\d+)'
        match = re.search(weight_volume_pattern, product_name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            inferred_qty = int(match.group(1))
            logger.info(f"从商品名称提取重量/容量规格: {product_name} -> {inferred_spec}, 包装数量={inferred_qty}")
            return inferred_spec, inferred_qty
        
        # 使用单位转换器推断规格
        inferred_spec = self.unit_converter.infer_specification_from_name(product_name)
        if inferred_spec:
            # 解析规格中的包装数量
            package_quantity = self.parse_specification(inferred_spec)
            if package_quantity:
                logger.info(f"从商品名称推断规格: {product_name} -> {inferred_spec}, 包装数量={package_quantity}")
                return inferred_spec, package_quantity
        
        # 特定商品规则匹配
        spec_rules = [
            # XX入白膜格式，如"550纯净水24入白膜"
            (r'.*?(\d+)入白膜', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 白膜格式，如"550水24白膜"
            (r'.*?(\d+)白膜', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 445水溶C系列
            (r'445水溶C.*?(\d+)[入个]纸箱', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 东方树叶系列
            (r'东方树叶.*?(\d+\*\d+).*纸箱', lambda m: (m.group(1), int(m.group(1).split('*')[1]))),
            
            # 桶装
            (r'(\d+\.?\d*L)桶装', lambda m: (f"{m.group(1)}*1", 1)),
            
            # 树叶茶系
            (r'树叶.*?(\d+)[入个]纸箱', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 茶π系列
            (r'茶[πΠπ].*?(\d+)纸箱', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 通用入数匹配
            (r'.*?(\d+)[入个](?:纸箱|箱装|白膜)', lambda m: (f"1*{m.group(1)}", int(m.group(1)))),
            
            # 通用数字+纸箱格式
            (r'.*?(\d+)纸箱', lambda m: (f"1*{m.group(1)}", int(m.group(1))))
        ]
        
        # 尝试所有规则
        for pattern, formatter in spec_rules:
            match = re.search(pattern, product_name)
            if match:
                spec, qty = formatter(match)
                logger.info(f"根据特定规则推断规格: {product_name} -> {spec}, 包装数量={qty}")
                return spec, qty
        
        # 尝试直接从名称中提取数字*数字格式
        match = re.search(r'(\d+\*\d+)', product_name)
        if match:
            spec = match.group(1)
            package_quantity = self.parse_specification(spec)
            if package_quantity:
                logger.info(f"从名称中直接提取规格: {spec}, 包装数量={package_quantity}")
                return spec, package_quantity
        
        # 最后尝试提取任何位置的数字，默认典型件装数
        numbers = re.findall(r'\d+', product_name)
        if numbers:
            for num in numbers:
                # 检查是否为典型的件装数(12/15/24/30)
                if num in ['12', '15', '24', '30']:
                    spec = f"1*{num}"
                    logger.info(f"从名称中提取可能的件装数: {spec}, 包装数量={int(num)}")
                    return spec, int(num)
            
        logger.warning(f"无法从商品名'{product_name}' 推断规格")
        return None, None 
    
    def parse_specification(self, spec_str: str) -> Optional[int]:
        """
        解析规格字符串，提取包装数量
        支持格式：1*15, 1x15, 1*5*10
        
        Args:
            spec_str: 规格字符串
            
        Returns:
            包装数量，如果无法解析则返回None
        """
        if not spec_str or not isinstance(spec_str, str):
            return None
        
        try:
            # 清理规格字符串
            spec_str = clean_string(spec_str)
            
            # 匹配重量/容量格式，如"450g*15"、"450ml*15"
            match = re.search(r'\d+(?:g|ml|毫升|克)[*xX×](\d+)', spec_str)
            if match:
                # 返回后面的数量
                return int(match.group(1))
            
            # 匹配1*5*10 格式的三级规格
            match = re.search(r'(\d+)[\*xX×](\d+)[\*xX×](\d+)', spec_str)
            if match:
                # 取最后一个数字作为袋数量
                return int(match.group(3))
            
            # 匹配1*15, 1x15 格式
            match = re.search(r'(\d+)[\*xX×](\d+)', spec_str)
            if match:
                # 取第二个数字作为包装数量
                return int(match.group(2))
                
            # 匹配24瓶/件等格式
            match = re.search(r'(\d+)[瓶个支袋][/／](件|箱)', spec_str)
            if match:
                return int(match.group(1))
                
            # 匹配4L格式
            match = re.search(r'(\d+(?:\.\d+)?)\s*[Ll升][*×]?(\d+)?', spec_str)
            if match:
                # 如果有第二个数字，返回它；否则返回1
                return int(match.group(2)) if match.group(2) else 1
                
        except Exception as e:
            logger.warning(f"解析规格'{spec_str}'时出错: {e}")
            
        return None 