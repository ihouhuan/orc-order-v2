"""
单位转换模块
----------
提供单位转换功能，支持规格推断和单位自动提取。
"""

import re
import logging
from typing import Dict, Tuple, Optional, Any, List, Union

from ..utils.log_utils import get_logger

logger = get_logger(__name__)

class UnitConverter:
    """
    单位转换器：处理不同单位之间的转换，支持从商品名称推断规格
    """
    
    def __init__(self):
        """
        初始化单位转换器
        """
        # 特殊条码配置
        self.special_barcodes = {
            '6925019900087': {
                'multiplier': 10,  # 数量乘以10
                'target_unit': '瓶',  # 目标单位
                'description': '特殊处理：数量*10，单位转换为瓶'
            }
            # 可以添加更多特殊条码的配置
        }
        
        # 规格推断的正则表达式模式
        self.spec_patterns = [
            # 1*6、1x12、1X20等格式
            (r'(\d+)[*xX×](\d+)', r'\1*\2'),
            # 1*5*12和1x5x12等三级格式
            (r'(\d+)[*xX×](\d+)[*xX×](\d+)', r'\1*\2*\3'),
            # "xx入"格式，如"12入"、"24入"
            (r'(\d+)入', r'1*\1'),
            # "xxL*1"或"xx升*1"格式
            (r'([\d\.]+)[L升][*xX×]?(\d+)?', r'\1L*\2' if r'\2' else r'\1L*1'),
            # "xxkg*1"或"xx公斤*1"格式
            (r'([\d\.]+)(?:kg|公斤)[*xX×]?(\d+)?', r'\1kg*\2' if r'\2' else r'\1kg*1'),
            # "xxg*1"或"xx克*1"格式
            (r'([\d\.]+)(?:g|克)[*xX×]?(\d+)?', r'\1g*\2' if r'\2' else r'\1g*1'),
            # "xxmL*1"或"xx毫升*1"格式
            (r'([\d\.]+)(?:mL|毫升)[*xX×]?(\d+)?', r'\1mL*\2' if r'\2' else r'\1mL*1'),
        ]
    
    def extract_unit_from_quantity(self, quantity_str: str) -> Tuple[Optional[float], Optional[str]]:
        """
        从数量字符串中提取单位
        
        Args:
            quantity_str: 数量字符串，如"2箱"、"5件"
            
        Returns:
            (数量, 单位)的元组，如果无法提取则返回(None, None)
        """
        if not quantity_str or not isinstance(quantity_str, str):
            return None, None
        
        # 匹配数字+单位格式
        match = re.match(r'^([\d\.]+)\s*([^\d\s\.]+)$', quantity_str.strip())
        if match:
            try:
                num = float(match.group(1))
                unit = match.group(2)
                logger.info(f"从数量提取单位: {quantity_str} -> 数量={num}, 单位={unit}")
                return num, unit
            except ValueError:
                pass
        
        return None, None
    
    def extract_specification(self, text: str) -> Optional[str]:
        """
        从文本中提取规格信息
        
        Args:
            text: 文本字符串
            
        Returns:
            提取的规格字符串，如果无法提取则返回None
        """
        if not text or not isinstance(text, str):
            return None
        
        # 尝试所有模式
        for pattern, replacement in self.spec_patterns:
            match = re.search(pattern, text)
            if match:
                # 特殊处理三级格式，确保正确显示为1*5*12
                if '*' in replacement and replacement.count('*') == 1 and len(match.groups()) >= 2:
                    result = f"{match.group(1)}*{match.group(2)}"
                    logger.info(f"提取规格: {text} -> {result}")
                    return result
                # 特殊处理三级规格格式
                elif '*' in replacement and replacement.count('*') == 2 and len(match.groups()) >= 3:
                    result = f"{match.group(1)}*{match.group(2)}*{match.group(3)}"
                    logger.info(f"提取三级规格: {text} -> {result}")
                    return result
                # 一般情况
                else:
                    result = re.sub(pattern, replacement, text)
                    logger.info(f"提取规格: {text} -> {result}")
                    return result
                
        # 没有匹配任何模式
        return None
    
    def infer_specification_from_name(self, name: str) -> Optional[str]:
        """
        从商品名称中推断规格
        
        Args:
            name: 商品名称
            
        Returns:
            推断的规格，如果无法推断则返回None
        """
        if not name or not isinstance(name, str):
            return None
        
        # 特殊模式的名称处理
        # 如"445水溶C血橙15入纸箱" -> "1*15"
        pattern1 = r'.*(\d+)入'
        match = re.match(pattern1, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(入): {name} -> {inferred_spec}")
            return inferred_spec
        
        # 如"500-东方树叶-绿茶1*15-纸箱装" -> "1*15"
        pattern2 = r'.*(\d+)[*xX×](\d+).*'
        match = re.match(pattern2, name)
        if match:
            inferred_spec = f"{match.group(1)}*{match.group(2)}"
            logger.info(f"从名称推断规格(直接): {name} -> {inferred_spec}")
            return inferred_spec
        
        # 如"12.9L桶装水" -> "12.9L*1"
        pattern3 = r'.*?([\d\.]+)L.*'
        match = re.match(pattern3, name)
        if match:
            inferred_spec = f"{match.group(1)}L*1"
            logger.info(f"从名称推断规格(L): {name} -> {inferred_spec}")
            return inferred_spec
        
        # 从名称中提取规格
        spec = self.extract_specification(name)
        if spec:
            return spec
            
        return None
        
    def parse_specification(self, spec: str) -> Tuple[int, int, Optional[int]]:
        """
        解析规格字符串，支持1*12和1*5*12等格式
        
        Args:
            spec: 规格字符串
            
        Returns:
            (一级包装, 二级包装, 三级包装)元组，如果是二级包装，第三个值为None
        """
        if not spec or not isinstance(spec, str):
            return 1, 1, None
            
        # 处理三级包装，如1*5*12
        three_level_match = re.match(r'(\d+)[*xX×](\d+)[*xX×](\d+)', spec)
        if three_level_match:
            try:
                level1 = int(three_level_match.group(1))
                level2 = int(three_level_match.group(2))
                level3 = int(three_level_match.group(3))
                logger.info(f"解析三级规格: {spec} -> {level1}*{level2}*{level3}")
                return level1, level2, level3
            except ValueError:
                pass
                
        # 处理二级包装，如1*12
        two_level_match = re.match(r'(\d+)[*xX×](\d+)', spec)
        if two_level_match:
            try:
                level1 = int(two_level_match.group(1))
                level2 = int(two_level_match.group(2))
                logger.info(f"解析二级规格: {spec} -> {level1}*{level2}")
                return level1, level2, None
            except ValueError:
                pass
                
        # 特殊处理L/升为单位的规格，如12.5L*1
        volume_match = re.match(r'([\d\.]+)[L升][*xX×](\d+)', spec)
        if volume_match:
            try:
                volume = float(volume_match.group(1))
                quantity = int(volume_match.group(2))
                logger.info(f"解析容量规格: {spec} -> {volume}L*{quantity}")
                return 1, quantity, None
            except ValueError:
                pass
        
        # 默认值
        logger.warning(f"无法解析规格: {spec}，使用默认值1*1")
        return 1, 1, None
        
    def process_unit_conversion(self, product: Dict) -> Dict:
        """
        处理单位转换，按照以下规则：
        1. 特殊条码: 优先处理特殊条码
        2. "件"单位: 数量×包装数量, 单价÷包装数量, 单位转为"瓶"
        3. "箱"单位: 数量×包装数量, 单价÷包装数量, 单位转为"瓶"
        4. "提"和"盒"单位: 如果是三级规格, 按件处理; 如果是二级规格, 保持不变
        5. 其他单位: 保持不变
        
        Args:
            product: 商品信息字典
            
        Returns:
            处理后的商品信息字典
        """
        # 复制原始数据，避免修改原始字典
        result = product.copy()
        
        barcode = result.get('barcode', '')
        unit = result.get('unit', '')
        quantity = result.get('quantity', 0)
        price = result.get('price', 0)
        specification = result.get('specification', '')
        
        # 跳过无效数据
        if not barcode or not quantity:
            return result
        
        # 特殊条码处理
        if barcode in self.special_barcodes:
            special_config = self.special_barcodes[barcode]
            multiplier = special_config.get('multiplier', 1)
            target_unit = special_config.get('target_unit', '瓶')
            
            # 数量乘以倍数
            new_quantity = quantity * multiplier
            
            # 如果有单价，单价除以倍数
            new_price = price / multiplier if price else 0
            
            logger.info(f"特殊条码处理: {barcode}, 数量: {quantity} -> {new_quantity}, 单价: {price} -> {new_price}, 单位: {unit} -> {target_unit}")
            
            result['quantity'] = new_quantity
            result['price'] = new_price
            result['unit'] = target_unit
            return result
        
        # 没有规格信息，无法进行单位转换
        if not specification:
            return result
            
        # 解析规格信息
        level1, level2, level3 = self.parse_specification(specification)
        
        # "件"单位处理
        if unit in ['件']:
            # 计算包装数量（二级*三级，如果无三级则仅二级）
            packaging_count = level2 * (level3 or 1)
            
            # 数量×包装数量
            new_quantity = quantity * packaging_count
            
            # 单价÷包装数量
            new_price = price / packaging_count if price else 0
            
            logger.info(f"件单位处理: 数量: {quantity} -> {new_quantity}, 单价: {price} -> {new_price}, 单位: 件 -> 瓶")
            
            result['quantity'] = new_quantity
            result['price'] = new_price
            result['unit'] = '瓶'
            return result
            
        # "箱"单位处理 - 与"件"单位处理相同
        if unit in ['箱']:
            # 计算包装数量
            packaging_count = level2 * (level3 or 1)
            
            # 数量×包装数量
            new_quantity = quantity * packaging_count
            
            # 单价÷包装数量
            new_price = price / packaging_count if price else 0
            
            logger.info(f"箱单位处理: 数量: {quantity} -> {new_quantity}, 单价: {price} -> {new_price}, 单位: 箱 -> 瓶")
            
            result['quantity'] = new_quantity
            result['price'] = new_price
            result['unit'] = '瓶'
            return result
            
        # "提"和"盒"单位处理
        if unit in ['提', '盒']:
            # 如果是三级规格，按件处理
            if level3 is not None:
                # 计算包装数量 - 只乘以最后一级数量
                packaging_count = level3
                
                # 数量×包装数量
                new_quantity = quantity * packaging_count
                
                # 单价÷包装数量
                new_price = price / packaging_count if price else 0
                
                logger.info(f"提/盒单位(三级规格)处理: 数量: {quantity} -> {new_quantity}, 单价: {price} -> {new_price}, 单位: {unit} -> 瓶")
                
                result['quantity'] = new_quantity
                result['price'] = new_price
                result['unit'] = '瓶'
            else:
                # 如果是二级规格，保持不变
                logger.info(f"提/盒单位(二级规格)处理: 保持原样 数量: {quantity}, 单价: {price}, 单位: {unit}")
            
            return result
        
        # 其他单位保持不变
        logger.info(f"其他单位处理: 保持原样 数量: {quantity}, 单价: {price}, 单位: {unit}")
        return result 