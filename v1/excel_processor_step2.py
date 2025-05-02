#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel处理程序 - 第二步
-------------------
读取OCR识别后的Excel文件，提取条码、单价和数量，
并创建采购单Excel文件。
"""

import os
import sys
import re
import logging
import pandas as pd
import numpy as np
import xlrd
import xlwt
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Any
from datetime import datetime
import random
from xlutils.copy import copy as xlcopy
import time
import json

# 配置日志
logger = logging.getLogger(__name__)
if not logger.handlers:
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs', 'excel_processor.log')
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)
logger.info("初始化日志系统")

class ExcelProcessorStep2:
    """
    Excel处理器第二步：处理OCR识别后的Excel文件，
    提取条码、单价和数量，并按照银豹采购单模板的格式填充
    """
    
    def __init__(self, output_dir="output"):
        """
        初始化Excel处理器，并设置输出目录
        """
        logger.info("初始化ExcelProcessorStep2")
        self.output_dir = output_dir
        
        # 确保输出目录存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logger.info(f"创建输出目录: {output_dir}")
        
        # 设置路径
        self.template_path = os.path.join("templets", "银豹-采购单模板.xls")
        
        # 检查模板文件是否存在
        if not os.path.exists(self.template_path):
            logger.error(f"模板文件不存在: {self.template_path}")
            raise FileNotFoundError(f"模板文件不存在: {self.template_path}")
        
        # 用于记录已处理的文件
        self.cache_file = os.path.join(output_dir, "processed_files.json")
        self.processed_files = {}  # 清空已处理文件记录
        
        # 特殊条码配置
        self.special_barcodes = {
            '6925019900087': {
                'multiplier': 10,  # 数量乘以10
                'target_unit': '瓶',  # 目标单位
                'description': '特殊处理：数量*10，单位转换为瓶'
            }
            # 可以在这里添加更多特殊条码的配置
        }
        
        logger.info(f"初始化完成，模板文件: {self.template_path}")
    
    def _load_processed_files(self):
        """加载已处理文件的缓存"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                logger.info(f"加载已处理文件缓存，共{len(cache)} 条记录")
                return cache
            except Exception as e:
                logger.warning(f"读取缓存文件失败: {e}")
        return {}
        
    def _save_processed_files(self):
        """保存已处理文件的缓存"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.processed_files, f, ensure_ascii=False, indent=2)
            logger.info(f"已更新处理文件缓存，共{len(self.processed_files)} 条记录")
        except Exception as e:
            logger.warning(f"保存缓存文件失败: {e}")
    
    def get_latest_excel(self):
        """
        获取output目录下最新的Excel文件
        """
        logger.info(f"搜索目录 {self.output_dir} 中的Excel文件")
        excel_files = []
        
        for file in os.listdir(self.output_dir):
            # 忽略临时文件（以~$开头的文件）和已处理的文件（以"采购单_"开头的文件）
            if file.lower().endswith('.xlsx') and not file.startswith('~$') and not file.startswith('采购单_'):
                file_path = os.path.join(self.output_dir, file)
                excel_files.append((file_path, os.path.getmtime(file_path)))
            
            if not excel_files:
                logger.warning(f"未在 {self.output_dir} 目录下找到未处理的Excel文件")
                return None
        
        # 按修改时间排序，获取最新的文件
        latest_file = sorted(excel_files, key=lambda x: x[1], reverse=True)[0][0]
        logger.info(f"找到最新的Excel文件: {latest_file}")
        return latest_file
    
    def validate_barcode(self, barcode):
        """
        验证条码是否有效
        新增功能：如果条码是"仓库"，则返回False以避免误认为有效条码
        """
        # 处理"仓库"特殊情况
        if isinstance(barcode, str) and barcode.strip() in ["仓库", "仓库全名"]:
            logger.warning(f"条码为仓库标识: {barcode}")
            return False
            
        # 处理科学计数法
        if isinstance(barcode, (int, float)):
            barcode = f"{barcode:.0f}"
            
        # 清理条码格式，移除可能的非数字字符（包括小数点）
        barcode_clean = re.sub(r'\.0+$', '', str(barcode))  # 移除末尾0
        barcode_clean = re.sub(r'\D', '', barcode_clean)  # 只保留数字
        
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
            
        logger.info(f"条码验证通过: {barcode_clean}")
        return True
    
    def parse_specification(self, spec_str):
        """
        解析规格字符串，提取包装数量
        支持格式1*15 1x15 格式
        
        新增支持*5*10 格式，其中最后的数字表示包装数量（例如：1袋）
        """
        if not spec_str or not isinstance(spec_str, str):
            logger.warning(f"无效的规格字符串: {spec_str}")
            return None
    
        try:
            # 清理规格字符串
            spec_str = spec_str.strip()
            
            # 新增：匹配1*5*10 格式的三级规格
            match = re.search(r'(\d+)[\*xX×](\d+)[\*xX×](\d+)', spec_str)
            if match:
                # 取最后一个数字作为袋数量
                return int(match.group(3))
            
            # 1. 匹配 1*15 1x15 格式
            match = re.search(r'(\d+)[\*xX×](\d+)', spec_str)
            if match:
                # 取第二个数字作为包装数量
                return int(match.group(2))
                
            # 2. 匹配 24瓶个支袋格式
            match = re.search(r'(\d+)[瓶个支袋][/／](件|箱)', spec_str)
            if match:
                return int(match.group(1))
                
            # 3. 匹配 500ml*15 格式
            match = re.search(r'\d+(?:ml|ML|毫升)[\*xX×](\d+)', spec_str)
            if match:
                return int(match.group(1))
            
            # 4. 提取最后一个数字作为包装数量（兜底方案）
            numbers = re.findall(r'\d+', spec_str)
            if numbers:
                # 对于类似 "330ml*24" 的规格，最后一个数字通常是包装数量
                return int(numbers[-1])
                
        except (ValueError, IndexError) as e:
            logger.warning(f"解析规格'{spec_str}'时出错: {e}")
            
        return None
    
    def infer_specification_from_name(self, product_name):
        """
        从商品名称推断规格
        根据特定的命名规则匹配规格信息
        
        示例
        - 445水溶C血5入纸-> 1*15
        - 500-东方树叶-绿茶1*15-纸箱开盖活动装 -> 1*15
        - 12.9L桶装-> 12.9L*1
        - 900树叶茉莉花茶12入纸-> 1*12
        - 500茶π蜜桃乌龙15纸箱 -> 1*15
        """
        if not product_name or not isinstance(product_name, str):
            logger.warning(f"无效的商品名: {product_name}")
            return None, None
            
        product_name = product_name.strip()
        logger.info(f"从商品名称推断规格: {product_name}")
        
        # 特定商品规则匹配
        spec_rules = [
            # 445水溶C系列
            (r'445水溶C.*?(\d+)[入个]纸箱', lambda m: f"1*{m.group(1)}"),
            
            # 东方树叶系列
            (r'东方树叶.*?(\d+\*\d+).*纸箱', lambda m: m.group(1)),
            (r'东方树叶.*?纸箱.*?(\d+\*\d+)', lambda m: m.group(1)),
            
            # 桶装
            (r'(\d+\.?\d*L)桶装', lambda m: f"{m.group(1)}*1"),
            
            # 树叶茶系
            (r'树叶.*?(\d+)[入个]纸箱', lambda m: f"1*{m.group(1)}"),
            (r'(\d+)树叶.*?(\d+)[入个]纸箱', lambda m: f"1*{m.group(2)}"),
            
            # 茶m系列
            (r'茶m.*?(\d+)纸箱', lambda m: f"1*{m.group(1)}"),
            (r'(\d+)茶m.*?(\d+)纸箱', lambda m: f"1*{m.group(2)}"),
            
            # 茶π系列
            (r'茶[πΠπ].*?(\d+)纸箱', lambda m: f"1*{m.group(1)}"),
            (r'(\d+)茶[πΠπ].*?(\d+)纸箱', lambda m: f"1*{m.group(2)}"),
            
            # 通用入数匹配
            (r'.*?(\d+)[入个](?:纸箱|箱装)', lambda m: f"1*{m.group(1)}"),
            (r'.*?箱装.*?(\d+)[入个]', lambda m: f"1*{m.group(1)}"),
            
            # 通用数字+纸箱格式，如"500茶π蜜桃乌龙15纸箱"
            (r'.*?(\d+)纸箱', lambda m: f"1*{m.group(1)}")
        ]
        
        # 尝试所有规则
        for pattern, formatter in spec_rules:
            match = re.search(pattern, product_name)
            if match:
                spec = formatter(match)
                logger.info(f"根据名称 '{product_name}' 推断规格: {spec}")
                
                # 提取包装数量
                package_quantity = self.parse_specification(spec)
                if package_quantity:
                    return spec, package_quantity
        
        # 尝试直接从名称中提取数字*数字格式
        match = re.search(r'(\d+\*\d+)', product_name)
        if match:
            spec = match.group(1)
            package_quantity = self.parse_specification(spec)
            if package_quantity:
                logger.info(f"从名称中直接提取规格: {spec}, 包装数量={package_quantity}")
                return spec, package_quantity
                
        # 尝试从名称中提取末尾数字
        match = re.search(r'(\d+)[入个]$', product_name)
        if match:
            qty = match.group(1)
            spec = f"1*{qty}"
            logger.info(f"从名称末尾提取入数: {spec}")
            return spec, int(qty)
        
        # 最后尝试提取任何位置的数字，默认如果有数字15，很可能5件装
        numbers = re.findall(r'\d+', product_name)
        if numbers:
            for num in numbers:
                # 检查是否为典型的件装数(12/15/24/30)
                if num in ['12', '15', '24', '30']:
                    spec = f"1*{num}"
                    logger.info(f"从名称中提取可能的件装数: {spec}")
                    return spec, int(num)
            
        logger.warning(f"无法从商品名'{product_name}' 推断规格")
        return None, None
    
    def extract_unit_from_quantity(self, quantity_str):
        """
        从数量字符串中提取单位
        例如
        - '2' -> (2, '')
        - '5' -> (5, '')
        - '3' -> (3, '')
        - '10' -> (10, '')
        """
        if not quantity_str:
            return None, None
            
        # 如果是数字，直接返回数字和None
        if isinstance(quantity_str, (int, float)):
            return float(quantity_str), None
            
        # 转为字符串并清理
        quantity_str = str(quantity_str).strip()
        logger.info(f"从数量字符串提取单位: {quantity_str}")
        
        # 匹配数字+单位格式
        match = re.match(r'^([\d\.]+)\s*([^\d\s\.]+)$', quantity_str)
        if match:
            try:
                value = float(match.group(1))
                unit = match.group(2)
                logger.info(f"提取到数字: {value}, 单位: {unit}")
                return value, unit
            except ValueError:
                logger.warning(f"无法解析数量: {match.group(1)}")
        
        # 如果只有数字，直接返回数字
        if re.match(r'^[\d\.]+$', quantity_str):
            try:
                value = float(quantity_str)
                logger.info(f"提取到数字: {value}, 无单位")
                return value, None
            except ValueError:
                logger.warning(f"无法解析纯数字数字: {quantity_str}")
                
        # 如果只有单位，尝试查找其他可能包含数字的部分
        match = re.match(r'^([^\d\s\.]+)$', quantity_str)
        if match:
            unit = match.group(1)
            logger.info(f"仅提取到单位: {unit}, 无数值")
            return None, unit
                
        logger.warning(f"无法提取数量和单位: {quantity_str}")
        return None, None
    
    def extract_barcode(self, df: pd.DataFrame) -> List[str]:
        """从数据框中提取条码"""
        barcodes = []
        
        # 遍历数据框查找条码
        for _, row in df.iterrows():
            for col_name, value in row.items():
                # 转换为字符串并处理
                if value is not None and not pd.isna(value):
                    value_str = str(value).strip()
                    # 特殊处理特定条码
                    if value_str == "5321545613":
                        barcodes.append(value_str)
                        logger.info(f"特殊条码提取: {value_str}")
                        continue
                        
                    if self.validate_barcode(value_str):
                        # 提取数字部分
                        barcode = re.sub(r'\D', '', value_str)
                        barcodes.append(barcode)
        
        logger.info(f"提取到{len(barcodes)} 个条码")
        return barcodes
    
    def extract_product_info(self, df: pd.DataFrame) -> List[Dict]:
        """
        提取产品信息,包括条码、单价、数量、金额等
        增加识别赠品功能：金额为0或为空的产品视为赠品
        修改后的功能：当没有有效条码时，使用行号作为临时条码
        """
        logger.info(f"正在从数据帧中提取产品信息")
        product_info = []
        
        try:
            # 打印列名，用于调试
            logger.info(f"Excel文件的列名: {df.columns.tolist()}")
            
            # 检查是否有特殊表头结构（如"武侯环球乐百惠便利店3333.xlsx"）
            # 判断依据：检查第3行是否包含常见的商品表头信息
            special_header = False
            if len(df) > 3:  # 确保有足够的行
                row3 = df.iloc[3].astype(str)
                header_keywords = ['行号', '条形码', '条码', '商品名称', '规格', '单价', '数量', '金额', '单位']
                # 计算匹配的关键词数量
                matches = sum(1 for keyword in header_keywords if any(keyword in str(val) for val in row3.values))
                # 如果匹配了至少3个关键词，认为第3行是表头
                if matches >= 3:
                    logger.info(f"检测到特殊表头结构，使用第3行作为列名: {row3.values.tolist()}")
                    # 创建新的数据帧，使用第3行作为列名，数据从第4行开始
                    header_row = df.iloc[3]
                    data_rows = df.iloc[4:].reset_index(drop=True)
                    # 为每一列分配一个名称（避免重复的列名）
                    new_columns = []
                    for i, col in enumerate(header_row):
                        col_str = str(col)
                        if col_str == 'nan' or col_str == 'None' or pd.isna(col):
                            new_columns.append(f"Col_{i}")
                        else:
                            new_columns.append(col_str)
                    # 使用新列名创建新的DataFrame
                    data_rows.columns = new_columns
                    df = data_rows
                    special_header = True
                    logger.info(f"重新构建的数据帧列名: {df.columns.tolist()}")
            
            # 检查是否有商品条码
            if '商品条码' in df.columns:
                # 遍历数据框的每一行
                for index, row in df.iterrows():
                    # 打印当前行的所有值，用于调试
                    logger.info(f"处理行{index+1}: {row.to_dict()}")
                    
                    # 跳过空行
                    if row.isna().all():
                        logger.info(f"跳过空行: {index+1}")
                        continue
                    
                    # 跳过小计行
                    if any('小计' in str(val) for val in row.values if isinstance(val, str)):
                        logger.info(f"跳过小计行: {index+1}")
                        continue
                    
                    # 获取条码（直接从商品条码列获取）
                    barcode_value = row['商品条码']
                    if pd.isna(barcode_value):
                        logger.info(f"跳过无条码行: {index+1}")
                        continue
                    
                    # 处理条码
                    barcode = str(int(barcode_value)) if isinstance(barcode_value, (int, float)) else str(barcode_value)
                    if not self.validate_barcode(barcode):
                        logger.warning(f"无效条码: {barcode}")
                        continue
                    
                    # 提取其他信息
                    product = {
                        'barcode': barcode,
                        'name': row.get('商品全名', ''),
                        'specification': row.get('规格', ''),
                        'unit': row.get('单位', ''),
                        'quantity': 0,
                        'unit_price': 0,
                        'amount': 0,
                        'is_gift': False,
                        'package_quantity': 1  # 默认包装数量
                    }
                    
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
                            inferred_spec, inferred_qty = self.infer_specification_from_name(product['name'])
                            if inferred_spec:
                                product['specification'] = inferred_spec
                                product['package_quantity'] = inferred_qty
                                logger.info(f"从商品名称推断规格: {product['name']} -> {inferred_spec}, 包装数量={inferred_qty}")
                    
                    # 提取数量和可能的单位
                    if '数量' in df.columns and not pd.isna(row['数量']):
                        try:
                            # 尝试从数量中提取单位和数量
                            extracted_qty, extracted_unit = self.extract_unit_from_quantity(row['数量'])
                            
                            # 处理提取到的数量
                            if extracted_qty is not None:
                                product['quantity'] = extracted_qty
                                logger.info(f"提取数量: {product['quantity']}")
                                
                                # 处理提取到的单位
                                if extracted_unit and (not product['unit'] or product['unit'] == ''):
                                    product['unit'] = extracted_unit
                                    logger.info(f"从数量中提取单位: {extracted_unit}")
                            else:
                                # 如果没有提取到数量，使用原始方法
                                product['quantity'] = float(row['数量'])
                                logger.info(f"使用原始数量: {product['quantity']}")
                        except (ValueError, TypeError) as e:
                            logger.warning(f"无效的数量: {row['数量']}, 错误: {str(e)}")
                    
                    # 提取单位（如果还没有单位）
                    if (not product['unit'] or product['unit'] == '') and '单位' in df.columns and not pd.isna(row['单位']):
                        product['unit'] = str(row['单位'])
                        logger.info(f"从单位列提取单位: {product['unit']}")
                    
                    # 提取单价
                    if '单价' in df.columns:
                        if pd.isna(row['单价']):
                            # 单价为空，视为赠品
                            is_gift = True
                            logger.info(f"单价为空，视为赠品")
                        else:
                            try:
                                # 如果单价是字符串且不是数字，视为赠品
                                if isinstance(row['单价'], str) and not row['单价'].replace('.', '').isdigit():
                                    is_gift = True
                                    logger.info(f"单价不是有效数字({row['单价']})，视为赠品")
                                else:
                                    product['unit_price'] = float(row['单价'])
                                    logger.info(f"提取单价: {product['unit_price']}")
                            except (ValueError, TypeError):
                                is_gift = True
                                logger.warning(f"无效的单价: {row['单价']}")
                    
                    # 提取金额
                    if '金额' in df.columns:
                        if amount_col and not pd.isna(row[amount_col]):
                            try:
                                # 清理金额字符串，处理可能的范围值（如"40-44"）
                                amount_str = str(row[amount_col])
                                if '-' in amount_str:
                                    # 如果是范围，取第一个值
                                    amount_str = amount_str.split('-')[0]
                                    logger.info(f"金额为范围值({row[amount_col]})，取第一个值: {amount_str}")
                                
                                # 尝试转换为浮点数
                                product['amount'] = float(amount_str)
                                logger.info(f"提取金额: {product['amount']}")
                            except (ValueError, TypeError) as e:
                                logger.warning(f"无效的金额: {row[amount_col]}, 错误: {e}")
                                # 金额无效时，设为0
                                product['amount'] = 0
                                logger.warning(f"设置金额为0")
                        else:
                            # 如果没有金额，计算金额
                            product['amount'] = product['quantity'] * product['unit_price']
                    
                    # 判断是否为赠品
                    is_gift = False
                    
                    # 赠品识别规则，根据README要求
                    # 1. 商品单价为0或为空
                    if product['unit_price'] == 0:
                        is_gift = True
                        logger.info(f"单价为空，视为赠品")
                    
                    # 2. 商品金额为0或为空
                    if not is_gift and amount_col:
                        try:
                            if pd.isna(row[amount_col]):
                                is_gift = True
                                logger.info(f"金额为空，视为赠品")
                            else:
                                # 清理金额字符串，处理可能的范围值（如"40-44"）
                                amount_str = str(row[amount_col])
                                if '-' in amount_str:
                                    # 如果是范围，取第一个值
                                    amount_str = amount_str.split('-')[0]
                                    logger.info(f"金额为范围值({row[amount_col]})，取第一个值: {amount_str}")
                                
                                # 转换为浮点数并检查是否为0
                                amount_val = float(amount_str)
                                if amount_val == 0:
                                    is_gift = True
                                    logger.info(f"金额为0，视为赠品")
                        except (ValueError, TypeError) as e:
                            logger.warning(f"无法解析金额: {row[amount_col]}, 错误: {e}")
                            # 金额无效时，不视为赠品，继续处理
                    
                    # 从赠送量列提取赠品数量
                    gift_quantity = 0
                    if '赠送量' in df.columns and not pd.isna(row['赠送量']):
                        try:
                            gift_quantity = float(row['赠送量'])
                            if gift_quantity > 0:
                                # 如果有明确的赠送量，总是创建赠品记录
                                logger.info(f"提取赠送量: {gift_quantity}")
                        except (ValueError, TypeError):
                            logger.warning(f"无效的赠送量: {row['赠送量']}")
                    
                    # 处理单位转换
                    self.process_unit_conversion(product)
                    
                    # 如果单价为0但有金额和数量，计算单价（非赠品情况）
                    if not is_gift and product['unit_price'] == 0 and product['amount'] > 0 and product['quantity'] > 0:
                        product['unit_price'] = product['amount'] / product['quantity']
                        logger.info(f"计算单价: {product['amount']} / {product['quantity']} = {product['unit_price']}")
                    
                    # 处理产品添加逻辑
                    product['is_gift'] = is_gift
                    
                    if is_gift:
                        # 如果是赠品且数量>0，使用商品本身的数量
                        if product['quantity'] > 0:
                            logger.info(f"添加赠品商品: 条码={barcode}, 数量={product['quantity']}")
                            product_info.append(product)
                    else:
                        # 正常商品
                        if product['quantity'] > 0:
                            logger.info(f"添加正常商品: 条码={barcode}, 数量={product['quantity']}, 单价={product['unit_price']}")
                            product_info.append(product)
                        
                    # 如果有额外的赠送量，添加专门的赠品记录
                    if gift_quantity > 0 and not is_gift:
                        gift_product = product.copy()
                        gift_product['is_gift'] = True
                        gift_product['quantity'] = gift_quantity
                        gift_product['unit_price'] = 0
                        gift_product['amount'] = 0
                        product_info.append(gift_product)
                        logger.info(f"添加额外赠品: 条码={barcode}, 数量={gift_quantity}")
                
                logger.info(f"提取到{len(product_info)} 个产品信息")
                return product_info
            
            # 如果没有直接匹配的列名，尝试使用更复杂的匹配逻辑
            logger.info("未找到直接匹配的列名或未提取到产品，尝试使用更复杂的匹配逻辑")
            # 定义可能的列名
            expected_columns = {
                '序号': ['序号', '行号', 'NO', 'NO.', '行号', '行号', '行号'],
                '条码': ['条码', '条形码', '商品条码', 'barcode', '商品条形码', '条形码', '商品条码', '商品编码', '商品编号', '条形码', '基本条码'],
                '名称': ['名称', '品名', '产品名称', '商品名称', '货物名称'],
                '规格': ['规格', '包装规格', '包装', '商品规格', '规格型号'],
                '采购单价': ['单价', '价格', '采购单价', '销售价'],
                '单位': ['单位', '采购单位'],
                '数量': ['数量', '采购数量', '购买数量', '采购数量', '订单数量', '采购数量'],
                '金额': ['金额', '订单金额', '总金额', '总价金额', '小计（元）'],
                '赠送量': ['赠送量', '赠品数量', '赠送数量', '赠品'],
            }
            
            # 如果是特殊表头处理后的数据，尝试直接从列名匹配
            if special_header:
                logger.info("使用特殊表头处理后的列名进行匹配")
                direct_map = {
                    '行号': '序号',
                    '条形码': '条码',
                    '商品名称': '名称',
                    '规格': '规格',
                    '单价': '采购单价',
                    '单位': '单位',
                    '数量': '数量',
                    '金额': '金额',
                    '箱数': '箱数',  # 可能包含单位信息
                }
                
                column_mapping = {}
                for target_key, source_key in direct_map.items():
                    if target_key in df.columns:
                        column_mapping[source_key] = target_key
                        logger.info(f"特殊表头匹配: {source_key} -> {target_key}")
            
            # 如果特殊表头处理没有找到足够的列，或者不是特殊表头，使用原有的映射逻辑
            if not special_header or len(column_mapping) < 3:
                # 检查第一行的内容，尝试判断是否是特殊格式的Excel
                if len(df) > 0:  # 确保DataFrame不为空
                    first_row = df.iloc[0].astype(str)
                    # 检查是否包含"商品全名"、"基本条码"、"仓库全名"等特定字段
                    if any("商品全名" in str(val) for val in first_row.values) and any("基本条码" in str(val) for val in first_row.values):
                        logger.info("检测到特殊格式Excel，使用特定的列映射")
                        
                        # 找出各列的索引
                        name_idx = None
                        barcode_idx = None
                        spec_idx = None
                        unit_idx = None
                        qty_idx = None
                        price_idx = None
                        amount_idx = None
                        
                        for idx, val in enumerate(first_row):
                            val_str = str(val).strip()
                            if val_str == "商品全名":
                                name_idx = df.columns[idx]
                            elif val_str == "基本条码":
                                barcode_idx = df.columns[idx]
                            elif val_str == "规格":
                                spec_idx = df.columns[idx]
                            elif val_str == "数量":
                                qty_idx = df.columns[idx]
                            elif val_str == "单位":
                                unit_idx = df.columns[idx]
                            elif val_str == "单价":
                                price_idx = df.columns[idx]
                            elif val_str == "金额":
                                amount_idx = df.columns[idx]
                        
                        # 使用找到的索引创建列映射
                        if name_idx and barcode_idx:
                            column_mapping = {
                                '名称': name_idx,
                                '条码': barcode_idx
                            }
                            
                            if spec_idx:
                                column_mapping['规格'] = spec_idx
                            if unit_idx:
                                column_mapping['单位'] = unit_idx
                            if qty_idx:
                                column_mapping['数量'] = qty_idx
                            if price_idx:
                                column_mapping['采购单价'] = price_idx
                            if amount_idx:
                                column_mapping['金额'] = amount_idx
                            
                            logger.info(f"特殊格式Excel的列映射: {column_mapping}")
                            
                            # 跳过第一行（表头）
                            df = df.iloc[1:].reset_index(drop=True)
                            logger.info("已跳过第一行（表头）")
                        else:
                            logger.warning("无法在特殊格式Excel中找到必要的列")
                    else:
                        # 映射实际的列名
                        column_mapping = {}
                    
                        # 检查是否有表头
                        has_header = False
                        for col in df.columns:
                            if not str(col).startswith('Unnamed:'):
                                has_header = True
                                break
                        
                        if has_header:
                            # 有表头的情况，使用原有的映射逻辑
                            for key, patterns in expected_columns.items():
                                for col in df.columns:
                                    # 移除列名中的空白字符以进行比较
                                    clean_col = re.sub(r'\s+', '', str(col))
                                    for pattern in patterns:
                                        clean_pattern = re.sub(r'\s+', '', pattern)
                                        if clean_col == clean_pattern:
                                            column_mapping[key] = col
                                            break
                                    if key in column_mapping:
                                        break
                        else:
                            # 无表头的情况，根据列的位置进行映射
                            # 假设列的顺序是：空列、序号、条码、名称、规格、单价、单位、数量、金额
                            if len(df.columns) >= 9:
                                column_mapping = {
                                    '序号': df.columns[1],  # Unnamed: 1
                                    '条码': df.columns[2],  # Unnamed: 2
                                    '名称': df.columns[3],  # Unnamed: 3
                                    '规格': df.columns[4],  # Unnamed: 4
                                    '采购单价': df.columns[7],  # Unnamed: 7
                                    '单位': df.columns[5],  # Unnamed: 5
                                    '数量': df.columns[6],  # Unnamed: 6
                                    '金额': df.columns[8]   # Unnamed: 8
                                }
                            else:
                                logger.warning(f"列数不足，无法进行映射。当前列数: {len(df.columns)}")
                                return []
            
            logger.info(f"列映射结果: {column_mapping}")
            
            # 如果找到了必要的列，直接从DataFrame提取数据
            if '条码' in column_mapping:
                barcode_col = column_mapping['条码']
                quantity_col = column_mapping.get('数量')
                price_col = column_mapping.get('采购单价')
                amount_col = column_mapping.get('金额')
                unit_col = column_mapping.get('单位')
                spec_col = column_mapping.get('规格')
                gift_col = column_mapping.get('赠送量')
                
                # 详细打印各行的关键数据
                logger.info("逐行显示数据内容:")
                for idx, row in df.iterrows():
                    # 获取关键字段数据
                    barcode_val = row[barcode_col] if barcode_col and not pd.isna(row[barcode_col]) else ""
                    quantity_val = row[quantity_col] if quantity_col and not pd.isna(row[quantity_col]) else ""
                    unit_val = row[unit_col] if unit_col and not pd.isna(row[unit_col]) else ""
                    price_val = row[price_col] if price_col and not pd.isna(row[price_col]) else ""
                    spec_val = row[spec_col] if spec_col and not pd.isna(row[spec_col]) else ""
                    gift_val = row[gift_col] if gift_col and not pd.isna(row[gift_col]) else ""
                    
                    logger.info(f"行{idx}, 条码:{barcode_val}, 数量:{quantity_val}, 单位:{unit_val}, " +
                               f"单价:{price_val}, 规格:{spec_val}, 赠送量:{gift_val}")
                
                # 逐行处理数据
                for idx, row in df.iterrows():
                    try:
                        # 跳过表头和汇总行
                        skip_row = False
                        for col in row.index:
                            if pd.notna(row[col]) and isinstance(row[col], str):
                                # 检查是否为表头、页脚或汇总行
                                if any(keyword in str(row[col]).lower() for keyword in ['序号', '小计', '合计', '总计', '页码', '行号', '页小计']):
                                    skip_row = True
                                    logger.info(f"跳过非商品行: {row[col]}")
                                    break
                        
                        if skip_row:
                            continue
                        
                        # 检查是否有有效的数量和单价
                        has_valid_data = False
                        if quantity_col and not pd.isna(row[quantity_col]):
                            try:
                                qty = float(row[quantity_col])
                                if qty > 0:
                                    has_valid_data = True
                            except (ValueError, TypeError):
                                pass
                        
                        if not has_valid_data:
                            logger.info(f"行{idx}没有有效数量，跳过")
                            continue
                            
                        # 提取或生成条码
                        barcode_value = row[barcode_col] if not pd.isna(row[barcode_col]) else None
                        
                        # 检查条码是否有效，如果是"仓库"或无效条码，跳过该行
                        barcode = None
                        if barcode_value is not None:
                            barcode_str = str(int(barcode_value)) if isinstance(barcode_value, (int, float)) else str(barcode_value)
                            if barcode_str not in ["仓库", "仓库全名"] and self.validate_barcode(barcode_str):
                                barcode = barcode_str
                        
                        # 如果没有有效条码，跳过该行
                        if barcode is None:
                            logger.info(f"行{idx}无有效条码，跳过该行")
                            continue
                        
                        # 创建产品信息
                        product = {
                            'barcode': barcode,
                            'name': row[column_mapping['名称']] if '名称' in column_mapping and not pd.isna(row[column_mapping['名称']]) else '',
                            'specification': row[spec_col] if spec_col and not pd.isna(row[spec_col]) else '',
                            'unit': row[unit_col] if unit_col and not pd.isna(row[unit_col]) else '',
                            'quantity': 0,
                            'unit_price': 0,
                            'amount': 0,
                            'is_gift': False,
                            'package_quantity': 1  # 默认包装数量
                        }
                        
                        # 提取规格并解析包装数量
                        if spec_col and not pd.isna(row[spec_col]):
                            product['specification'] = str(row[spec_col])
                            package_quantity = self.parse_specification(product['specification'])
                            if package_quantity:
                                product['package_quantity'] = package_quantity
                                logger.info(f"解析规格: {product['specification']} -> 包装数量={package_quantity}")
                        else:
                            # 逻辑1: 如果规格为空，尝试从商品名称推断规格
                            if '名称' in column_mapping and not pd.isna(row[column_mapping['名称']]):
                                product_name = str(row[column_mapping['名称']])
                                inferred_spec, inferred_qty = self.infer_specification_from_name(product_name)
                                if inferred_spec:
                                    product['specification'] = inferred_spec
                                    product['package_quantity'] = inferred_qty
                                    logger.info(f"从商品名称推断规格: {product_name} -> {inferred_spec}, 包装数量={inferred_qty}")
                        
                        # 提取数量和可能的单位
                        if quantity_col and not pd.isna(row[quantity_col]):
                            try:
                                # 尝试从数量中提取单位和数量
                                extracted_qty, extracted_unit = self.extract_unit_from_quantity(row[quantity_col])
                                
                                # 处理提取到的数量
                                if extracted_qty is not None:
                                    product['quantity'] = extracted_qty
                                    logger.info(f"提取数量: {product['quantity']}")
                                    
                                    # 处理提取到的单位
                                    if extracted_unit and (not product['unit'] or product['unit'] == ''):
                                        product['unit'] = extracted_unit
                                        logger.info(f"从数量中提取单位: {extracted_unit}")
                                else:
                                    # 如果没有提取到数量，使用原始方法
                                    product['quantity'] = float(row[quantity_col])
                                    logger.info(f"使用原始数量: {product['quantity']}")
                            except (ValueError, TypeError) as e:
                                logger.warning(f"无效的数量: {row[quantity_col]}, 错误: {str(e)}")
                                continue  # 如果数量无效，跳过此行
                        else:
                            # 如果没有数量，跳过此行
                            logger.warning(f"行{idx}缺少数量，跳过")
                            continue
                        
                        # 提取单价
                        if price_col and not pd.isna(row[price_col]):
                            try:
                                product['unit_price'] = float(row[price_col])
                                logger.info(f"提取单价: {product['unit_price']}")
                            except (ValueError, TypeError) as e:
                                logger.warning(f"无效的单价: {row[price_col]}, 错误: {e}")
                                # 单价无效时，可能是赠品
                                is_gift = True
                        
                        # 初始化赠品标志
                        is_gift = False
                        
                        # 提取金额
                        # 忽略金额栏中可能存在的备注信息
                        if amount_col and not pd.isna(row[amount_col]):
                            amount_value = row[amount_col]
                            if isinstance(amount_value, (int, float)):
                                # 如果是数字类型，直接使用
                                product['amount'] = float(amount_value)
                                logger.info(f"提取金额: {product['amount']}")
                                if product['amount'] == 0:
                                    is_gift = True
                                    logger.info(f"金额为0，视为赠品")
                            else:
                                # 如果不是数字类型，尝试从字符串中提取数字
                                try:
                                    # 尝试转换为浮点数
                                    amount_str = str(amount_value)
                                    if amount_str.replace('.', '', 1).isdigit():
                                        product['amount'] = float(amount_str)
                                        logger.info(f"从字符串提取金额: {product['amount']}")
                                        if product['amount'] == 0:
                                            is_gift = True
                                            logger.info(f"金额为0，视为赠品")
                                    else:
                                        # 金额栏含有非数字内容，可能是备注，此时使用单价*数量计算金额
                                        logger.warning(f"金额栏包含非数字内容: {amount_value}，将被视为备注，金额计算为单价*数量")
                                        product['amount'] = product['unit_price'] * product['quantity']
                                        logger.info(f"计算金额: {product['unit_price']} * {product['quantity']} = {product['amount']}")
                                except (ValueError, TypeError) as e:
                                    logger.warning(f"无法解析金额: {amount_value}, 错误: {e}")
                                    # 计算金额
                                    product['amount'] = product['unit_price'] * product['quantity']
                                    logger.info(f"计算金额: {product['unit_price']} * {product['quantity']} = {product['amount']}")
                        else:
                            # 如果金额为空，可能是赠品，或需要计算金额
                            if product['unit_price'] > 0:
                                product['amount'] = product['unit_price'] * product['quantity']
                                logger.info(f"计算金额: {product['unit_price']} * {product['quantity']} = {product['amount']}")
                            else:
                                is_gift = True
                                logger.info(f"单价或金额为空，视为赠品")
                        
                        # 处理单位转换
                        self.process_unit_conversion(product)
                        
                        # 处理产品添加逻辑
                        product['is_gift'] = is_gift
                        
                        if is_gift:
                            # 如果是赠品且数量>0，使用商品本身的数量
                            if product['quantity'] > 0:
                                logger.info(f"添加赠品商品: 条码={barcode}, 数量={product['quantity']}")
                                product_info.append(product)
                        else:
                            # 正常商品
                            if product['quantity'] > 0:
                                logger.info(f"添加正常商品: 条码={barcode}, 数量={product['quantity']}, 单价={product['unit_price']}")
                                product_info.append(product)
                        
                        # 如果有额外的赠送量，添加专门的赠品记录
                        if gift_col and not pd.isna(row[gift_col]):
                            try:
                                gift_quantity = float(row[gift_col])
                                if gift_quantity > 0:
                                    gift_product = product.copy()
                                    gift_product['is_gift'] = True
                                    gift_product['quantity'] = gift_quantity
                                    gift_product['unit_price'] = 0
                                    gift_product['amount'] = 0
                                    product_info.append(gift_product)
                                    logger.info(f"添加额外赠品: 条码={barcode}, 数量={gift_quantity}")
                            except (ValueError, TypeError) as e:
                                logger.warning(f"无效的赠送量: {row[gift_col]}, 错误: {e}")
                    
                    except Exception as e:
                        logger.warning(f"处理行{idx}时出错: {e}")
                        continue  # 跳过有错误的行，继续处理下一行
                
                logger.info(f"提取到{len(product_info)} 个产品信息")
                return product_info
        
        except Exception as e:
            logger.error(f"提取产品信息时出错: {e}", exc_info=True)
            return []
    
    def fill_template(self, template_file_path, products, output_file_path):
        """
        填充采购单模板并保存为新文件
        按照模板格式填充（银豹采购单模板）
        - 列B(1): 条码（必填）
        - 列C(2): 采购量（必填） 对于只有赠品的商品，此列为空
        - 列D(3): 赠送量 - 同一条码的赠品数量
        - 列E(4): 采购单价（必填）- 保留4位小数
        
        特殊处理
        - 同一条码既有正常商品又有赠品时，保持正常商品的采购量不变，将赠品数量填写到赠送量栏位
        - 只有赠品没有正常商品的情况，采购量列填写0，赠送量填写赠品数量
        - 赠品的判断依据：is_gift标记为True
        """
        logger.info(f"开始填充模板: {template_file_path}")
        
        try:
            # 打开模板文件
            workbook = xlrd.open_workbook(template_file_path)
            workbook = xlcopy(workbook)
            worksheet = workbook.get_sheet(0)  # 默认第一个工作表
            
            # 从第2行开始填充数据（索引从0开始，对应Excel中的行号）
            row_index = 1  # Excel的行号从0开始，对应Excel中的行号
            
            # 先对产品按条码进行分组，识别赠品和普通商品
            barcode_groups = {}
            
            # 遍历所有产品，按条码分组
            logger.info(f"开始处理{len(products)} 个产品信息")
            for product in products:
                barcode = product.get('barcode', '')
                if not barcode:
                    logger.warning(f"跳过无条码商品")
                    continue
                
                # 使用产品中的is_gift标记来判断是否为赠品
                is_gift = product.get('is_gift', False)
                
                # 获取数量和单位
                quantity = product.get('quantity', 0)
                unit_price = product.get('unit_price', 0)
                
                logger.info(f"处理商品: 条码={barcode}, 数量={quantity}, 单价={unit_price}, 是否赠品={is_gift}")
                
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
                            'price': unit_price
                        }
                        logger.info(f"发现正常商品：条码{barcode}, 数量={quantity}, 单价={unit_price}")
                    else:
                        # 如果有多个正常商品记录，累加数量
                        barcode_groups[barcode]['normal']['quantity'] += quantity
                        logger.info(f"累加正常商品数量：条码{barcode}, 新增={quantity}, 累计={barcode_groups[barcode]['normal']['quantity']}")
                        
                        # 如果单价不同，取平均值
                        if unit_price != barcode_groups[barcode]['normal']['price']:
                            avg_price = (barcode_groups[barcode]['normal']['price'] + unit_price) / 2
                            barcode_groups[barcode]['normal']['price'] = avg_price
                            logger.info(f"调整单价(取平均值)：条码{barcode}, 原价={barcode_groups[barcode]['normal']['price']}, 新价={unit_price}, 平均={avg_price}")
            
            # 输出调试信息
            logger.info(f"分组后共{len(barcode_groups)} 个不同条码的商品")
            for barcode, group in barcode_groups.items():
                if group['normal'] is not None:
                    logger.info(f"条码 {barcode} 处理结果：正常商品数量{group['normal']['quantity']}，单价{group['normal']['price']}，赠品数量{group['gift_quantity']}")
                else:
                    logger.info(f"条码 {barcode} 处理结果：只有赠品，数量={group['gift_quantity']}")
            
            # 准备填充数据
            for barcode, group in barcode_groups.items():
                # 1. 列B(1): 条码（必填）
                worksheet.write(row_index, 1, barcode)
                
                if group['normal'] is not None:
                    # 有正常商品
                    product = group['normal']['product']
                    
                    # 2. 列C(2): 采购量（必填） 使用正常商品的采购量
                    normal_quantity = group['normal']['quantity']
                    worksheet.write(row_index, 2, normal_quantity)
                    
                    # 3. 列D(3): 赠送量 - 添加赠品数量
                    if group['gift_quantity'] > 0:
                        worksheet.write(row_index, 3, group['gift_quantity'])
                        logger.info(f"条码 {barcode} 填充：采购量={normal_quantity}，赠品数量{group['gift_quantity']}")
                    
                    # 4. 列E(4): 采购单价（必填）
                    purchase_price = group['normal']['price']
                    style = xlwt.XFStyle()
                    style.num_format_str = '0.0000'
                    worksheet.write(row_index, 4, round(purchase_price, 4), style)
                    
                elif group['gift_quantity'] > 0:
                    # 只有赠品，没有正常商品
                    logger.info(f"条码 {barcode} 只有赠品，数量{group['gift_quantity']}，采购量=0，赠送量={group['gift_quantity']}")
                    
                    # 2. 列C(2): 采购量（必填） 对于只有赠品的条目，采购量填写为0
                    worksheet.write(row_index, 2, 0)
                    
                    # 3. 列D(3): 赠送量 - 填写赠品数量
                    worksheet.write(row_index, 3, group['gift_quantity'])
                    
                    # 4. 列E(4): 采购单价（必填） - 对于只有赠品的条目，采购单价为0
                    style = xlwt.XFStyle()
                    style.num_format_str = '0.0000'
                    worksheet.write(row_index, 4, 0, style)
                
                row_index += 1
            
            # 保存文件
            workbook.save(output_file_path)
            logger.info(f"采购单已保存: {output_file_path}")
            return True
            
        except Exception as e:
            logger.error(f"填充模板时出错: {str(e)}", exc_info=True)
            return False
    
    def create_new_xls(self, input_file_path, products):
        """
        根据输入的Excel文件创建新的采购单
        """
        try:
            # 获取输入文件的文件名（不带扩展名）
            input_filename = os.path.basename(input_file_path)
            name_without_ext = os.path.splitext(input_filename)[0]
            
            # 创建基本输出文件路径
            base_output_path = os.path.join("output", f"采购单_{name_without_ext}.xls")
            
            # 如果文件已存在，自动添加时间戳避免覆盖
            output_file_path = base_output_path
            if os.path.exists(base_output_path):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                name_parts = os.path.splitext(base_output_path)
                output_file_path = f"{name_parts[0]}_{timestamp}{name_parts[1]}"
                logger.info(f"文件 {base_output_path} 已存在，重命名为 {output_file_path}")
            
            # 填充模板
            result = self.fill_template(self.template_path, products, output_file_path)
            
            if result:
                logger.info(f"成功创建采购单: {output_file_path}")
                return output_file_path
            else:
                logger.error("创建采购单失败")
                return None
            
        except Exception as e:
            logger.error(f"创建采购单时出错: {str(e)}")
            return None

    def process_specific_file(self, file_path):
        """
        处理指定的Excel文件
        """
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return False

        # 检查文件是否已处理
        file_stat = os.stat(file_path)
        file_key = f"{os.path.basename(file_path)}_{file_stat.st_size}_{file_stat.st_mtime}"
        
        if file_key in self.processed_files:
            output_file = self.processed_files[file_key]
            if os.path.exists(output_file):
                logger.info(f"文件已处理过，采购单文件: {output_file}")
                return True
        
        logger.info(f"开始处理Excel文件: {file_path}")
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 删除行号列（如果存在）
            if '行号' in df.columns:
                df = df.drop('行号', axis=1)
                logger.info("已删除行号列")
            
            # 提取商品信息
            products = self.extract_product_info(df)
            
            if not products:
                logger.warning("未从Excel文件中提取到有效的商品信息")
                return False
            
            # 获取文件名（不含扩展名）
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 基本输出文件路径
            base_output_file = os.path.join(self.output_dir, f"采购单_{file_name}.xls")
            
            # 如果文件已存在，自动添加时间戳避免覆盖
            output_file = base_output_file
            if os.path.exists(base_output_file):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                name_parts = os.path.splitext(base_output_file)
                output_file = f"{name_parts[0]}_{timestamp}{name_parts[1]}"
                logger.info(f"文件 {base_output_file} 已存在，重命名为 {output_file}")
            
            # 填充模板
            result = self.fill_template(self.template_path, products, output_file)
            
            if result:
                # 记录已处理文件
                self.processed_files[file_key] = output_file
                self._save_processed_files()
                
                logger.info(f"Excel处理成功，采购单已保存至: {output_file}")
                return True
            else:
                logger.error("填充模板失败")
                return False
            
        except Exception as e:
            logger.error(f"处理Excel文件时出错: {str(e)}")
            return False
    
    def process_latest_file(self):
        """
        处理最新的Excel文件
        """
        latest_file = self.get_latest_excel()
        if not latest_file:
            logger.error("未找到可处理的Excel文件")
            return False
            
        return self.process_specific_file(latest_file)

    def process(self):
        """
        处理最新的Excel文件
        """
        return self.process_latest_file()

    def process_unit_conversion(self, product):
        """
        处理单位转换
        """
        if product['unit'] in ['提', '盒']:
            # 检查是否是特殊条码
            if product['barcode'] in self.special_barcodes:
                special_config = self.special_barcodes[product['barcode']]
                # 特殊条码处理
                actual_quantity = product['quantity'] * special_config['multiplier']
                logger.info(f"特殊条码处理: {product['quantity']}{product['unit']} -> {actual_quantity}{special_config['target_unit']}")
                
                # 更新产品信息
                product['original_quantity'] = product['quantity']
                product['quantity'] = actual_quantity
                product['original_unit'] = product['unit']
                product['unit'] = special_config['target_unit']
                
                # 如果有单价，计算转换后的单价
                if product['unit_price'] > 0:
                    product['original_unit_price'] = product['unit_price']
                    product['unit_price'] = product['unit_price'] / special_config['multiplier']
                    logger.info(f"单价转换: {product['original_unit_price']}/{product['original_unit']} -> {product['unit_price']}/{special_config['target_unit']}")
            else:
                # 提取规格中的数字
                spec_parts = re.findall(r'\d+', product['specification'])
                
                # 检查是否是1*5*12这样的三级格式
                if len(spec_parts) >= 3:
                    # 三级规格：按件处理
                    actual_quantity = product['quantity'] * product['package_quantity']
                    logger.info(f"{product['unit']}单位三级规格转换: {product['quantity']}{product['unit']} -> {actual_quantity}瓶")
                    
                    # 更新产品信息
                    product['original_quantity'] = product['quantity']
                    product['quantity'] = actual_quantity
                    product['original_unit'] = product['unit']
                    product['unit'] = '瓶'
                    
                    # 如果有单价，计算转换后的单价
                    if product['unit_price'] > 0:
                        product['original_unit_price'] = product['unit_price']
                        product['unit_price'] = product['unit_price'] / product['package_quantity']
                        logger.info(f"单价转换: {product['original_unit_price']}/{product['original_unit']} -> {product['unit_price']}/瓶")
                else:
                    # 二级规格：保持原数量不变
                    logger.info(f"{product['unit']}单位二级规格保持原数量: {product['quantity']}{product['unit']}")
        # 对于"件"单位或其他特殊条码的处理
        elif product['barcode'] in self.special_barcodes:
            special_config = self.special_barcodes[product['barcode']]
            # 特殊条码处理
            actual_quantity = product['quantity'] * special_config['multiplier']
            logger.info(f"特殊条码处理: {product['quantity']}{product['unit']} -> {actual_quantity}{special_config['target_unit']}")
            
            # 更新产品信息
            product['original_quantity'] = product['quantity']
            product['quantity'] = actual_quantity
            product['original_unit'] = product['unit']
            product['unit'] = special_config['target_unit']
            
            # 如果有单价，计算转换后的单价
            if product['unit_price'] > 0:
                product['original_unit_price'] = product['unit_price']
                product['unit_price'] = product['unit_price'] / special_config['multiplier']
                logger.info(f"单价转换: {product['original_unit_price']}/{product['original_unit']} -> {product['unit_price']}/{special_config['target_unit']}")
        elif product['unit'] == '件':
            # 标准件处理：数量×包装数量
            if product['package_quantity'] and product['package_quantity'] > 1:
                actual_quantity = product['quantity'] * product['package_quantity']
                logger.info(f"件单位转换: {product['quantity']}件 -> {actual_quantity}瓶")
                
                # 更新产品信息
                product['original_quantity'] = product['quantity']
                product['quantity'] = actual_quantity
                product['original_unit'] = product['unit']
                product['unit'] = '瓶'
                
                # 如果有单价，计算转换后的单价
                if product['unit_price'] > 0:
                    product['original_unit_price'] = product['unit_price']
                    product['unit_price'] = product['unit_price'] / product['package_quantity']
                    logger.info(f"单价转换: {product['original_unit_price']}/件 -> {product['unit_price']}/瓶")

def main():
    """主程序"""
    import argparse
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='Excel处理程序 - 第二步')
    parser.add_argument('--input', type=str, help='指定输入Excel文件路径，默认使用output目录中最新的Excel文件')
    parser.add_argument('--output', type=str, help='指定输出文件路径，默认使用模板文件路径加时间')
    args = parser.parse_args()
    
    processor = ExcelProcessorStep2()
    
    # 处理Excel文件
    try:
        # 根据是否指定输入文件选择处理方式
        if args.input:
            # 使用指定文件处理
            result = processor.process_specific_file(args.input)
        else:
            # 使用默认处理流程（查找最新文件）
            result = processor.process()
        
        if result:
            print("处理成功！已将数据填充并保存")
        else:
            print("处理失败！请查看日志了解详细信息")
    except Exception as e:
        logger.error(f"处理过程中发生错误: {e}", exc_info=True)
        print(f"处理过程中发生错误: {e}")
        print("请查看日志文件了解详细信息")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"程序执行过程中发生错误: {e}", exc_info=True)
        sys.exit(1) 
