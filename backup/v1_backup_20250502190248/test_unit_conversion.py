#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
单位转换处理规则测试
-------------------
这个脚本用于演示excel_processor_step2.py中的单位转换处理规则，
包括件、提、盒单位的处理，以及特殊条码的处理。
"""

import os
import sys
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def test_unit_conversion(barcode, unit, quantity, specification, unit_price):
    """
    测试单位转换处理逻辑
    """
    logger.info(f"测试条码: {barcode}, 单位: {unit}, 数量: {quantity}, 规格: {specification}, 单价: {unit_price}")
    
    # 特殊条码处理
    special_barcodes = {
        '6925019900087': {
            'multiplier': 10,  # 数量乘以10
            'target_unit': '瓶',  # 目标单位
            'description': '特殊处理：数量*10，单位转换为瓶'
        }
    }
    
    # 解析规格
    package_quantity = None
    is_tertiary_spec = False
    
    if specification:
        import re
        # 三级规格，如1*5*12
        match = re.search(r'(\d+)[\*xX×](\d+)[\*xX×](\d+)', specification)
        if match:
            package_quantity = int(match.group(3))
            is_tertiary_spec = True
        else:
            # 二级规格，如1*15
            match = re.search(r'(\d+)[\*xX×](\d+)', specification)
            if match:
                package_quantity = int(match.group(2))
    
    # 初始化结果
    result_quantity = quantity
    result_unit = unit
    result_unit_price = unit_price
    
    # 处理单位转换
    if barcode in special_barcodes:
        # 特殊条码处理
        special_config = special_barcodes[barcode]
        result_quantity = quantity * special_config['multiplier']
        result_unit = special_config['target_unit']
        
        if unit_price:
            result_unit_price = unit_price / special_config['multiplier']
        
        logger.info(f"特殊条码处理: {quantity}{unit} -> {result_quantity}{result_unit}")
        if unit_price:
            logger.info(f"单价转换: {unit_price}/{unit} -> {result_unit_price}/{result_unit}")
    
    elif unit in ['提', '盒']:
        # 提和盒单位特殊处理
        if is_tertiary_spec and package_quantity:
            # 三级规格：按照件的计算方式处理
            result_quantity = quantity * package_quantity
            result_unit = '瓶'
            
            if unit_price:
                result_unit_price = unit_price / package_quantity
            
            logger.info(f"{unit}单位三级规格转换: {quantity}{unit} -> {result_quantity}瓶")
            if unit_price:
                logger.info(f"单价转换: {unit_price}/{unit} -> {result_unit_price}/瓶")
        else:
            # 二级规格或无规格：保持原数量不变
            logger.info(f"{unit}单位二级规格保持原数量: {quantity}{unit}")
    
    elif unit == '件' and package_quantity:
        # 件单位处理：数量×包装数量
        result_quantity = quantity * package_quantity
        result_unit = '瓶'
        
        if unit_price:
            result_unit_price = unit_price / package_quantity
        
        logger.info(f"件单位转换: {quantity}件 -> {result_quantity}瓶")
        if unit_price:
            logger.info(f"单价转换: {unit_price}/件 -> {result_unit_price}/瓶")
    
    else:
        # 其他单位保持不变
        logger.info(f"保持原单位不变: {quantity}{unit}")
    
    # 输出处理结果
    logger.info(f"处理结果 => 数量: {result_quantity}, 单位: {result_unit}, 单价: {result_unit_price}")
    logger.info("-" * 50)
    
    return result_quantity, result_unit, result_unit_price

def run_tests():
    """运行一系列测试用例"""
    
    # 标准件单位测试
    test_unit_conversion("1234567890123", "件", 1, "1*12", 108)
    test_unit_conversion("1234567890124", "件", 2, "1*24", 120)
    
    # 提和盒单位测试 - 二级规格
    test_unit_conversion("1234567890125", "提", 3, "1*16", 50)
    test_unit_conversion("1234567890126", "盒", 5, "1*20", 60)
    
    # 提和盒单位测试 - 三级规格
    test_unit_conversion("1234567890127", "提", 2, "1*5*12", 100)
    test_unit_conversion("1234567890128", "盒", 3, "1*6*8", 120)
    
    # 特殊条码测试
    test_unit_conversion("6925019900087", "副", 2, "1*10", 50)
    test_unit_conversion("6925019900087", "提", 1, "1*16", 30)
    
    # 其他单位测试
    test_unit_conversion("1234567890129", "包", 4, "1*24", 12)
    test_unit_conversion("1234567890130", "瓶", 10, "", 5)

if __name__ == "__main__":
    logger.info("开始测试单位转换处理规则")
    run_tests()
    logger.info("单位转换处理规则测试完成") 