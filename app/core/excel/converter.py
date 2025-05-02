"""
单位转换处理模块
-------------
提供规格和单位的处理和转换功能。
"""

import re
from typing import Dict, List, Optional, Tuple, Any

from ..utils.log_utils import get_logger
from ..utils.string_utils import (
    clean_string,
    extract_number,
    extract_unit,
    extract_number_and_unit,
    parse_specification
)

logger = get_logger(__name__)

class UnitConverter:
    """
    单位转换器：处理商品规格和单位转换
    """
    
    def __init__(self):
        """初始化单位转换器"""
        # 特殊条码配置
        self.special_barcodes = {
            '6925019900087': {
                'multiplier': 10,  # 数量乘以10
                'target_unit': '瓶',  # 目标单位
                'description': '特殊处理：数量*10，单位转换为瓶'
            }
            # 可以在这里添加更多特殊条码的配置
        }
        
        # 有效的单位列表
        self.valid_units = ['件', '箱', '包', '提', '盒', '瓶', '个', '支', '袋', '副', '桶', '罐', 'L', 'l', '升']
        
        # 需要特殊处理的单位
        self.special_units = ['件', '箱', '提', '盒']
        
        logger.info("单位转换器初始化完成")
    
    def add_special_barcode(self, barcode: str, multiplier: int, target_unit: str, description: str = "") -> None:
        """
        添加特殊条码处理配置
        
        Args:
            barcode: 条码
            multiplier: 数量乘数
            target_unit: 目标单位
            description: 处理描述
        """
        self.special_barcodes[barcode] = {
            'multiplier': multiplier,
            'target_unit': target_unit,
            'description': description or f'特殊处理：数量*{multiplier}，单位转换为{target_unit}'
        }
        logger.info(f"添加特殊条码配置: {barcode}, {description}")
    
    def infer_specification_from_name(self, product_name: str) -> Optional[str]:
        """
        从商品名称推断规格
        
        Args:
            product_name: 商品名称
            
        Returns:
            推断的规格，如果无法推断则返回None
        """
        if not product_name or not isinstance(product_name, str):
            return None
            
        try:
            # 清理商品名称
            name = clean_string(product_name)
            
            # 1. 匹配 XX入纸箱 格式
            match = re.search(r'(\d+)入纸箱', name)
            if match:
                return f"1*{match.group(1)}"
                
            # 2. 匹配 绿茶1*15-纸箱装 格式
            match = re.search(r'(\d+)[*×xX](\d+)[-\s]?纸箱', name)
            if match:
                return f"{match.group(1)}*{match.group(2)}"
                
            # 3. 匹配 12.9L桶装水 格式
            match = re.search(r'([\d\.]+)[Ll升](?!.*[*×xX])', name)
            if match:
                return f"{match.group(1)}L*1"
                
            # 4. 匹配 商品12入纸箱 格式（数字在中间）
            match = re.search(r'\D(\d+)入\w*箱', name)
            if match:
                return f"1*{match.group(1)}"
                
            # 5. 匹配 商品15纸箱 格式（数字在中间）
            match = re.search(r'\D(\d+)\w*箱', name)
            if match:
                return f"1*{match.group(1)}"
                
            # 6. 匹配 商品1*30 格式
            match = re.search(r'(\d+)[*×xX](\d+)', name)
            if match:
                return f"{match.group(1)}*{match.group(2)}"
                
            logger.debug(f"无法从商品名称推断规格: {name}")
            return None
            
        except Exception as e:
            logger.error(f"从商品名称推断规格时出错: {e}")
            return None
    
    def extract_unit_from_quantity(self, quantity_str: str) -> Tuple[Optional[float], Optional[str]]:
        """
        从数量字符串提取单位
        
        Args:
            quantity_str: 数量字符串
            
        Returns:
            (数量, 单位)元组
        """
        if not quantity_str or not isinstance(quantity_str, str):
            return None, None
            
        try:
            # 清理数量字符串
            quantity_str = clean_string(quantity_str)
            
            # 提取数字和单位
            return extract_number_and_unit(quantity_str)
            
        except Exception as e:
            logger.error(f"从数量字符串提取单位时出错: {quantity_str}, 错误: {e}")
            return None, None
    
    def process_unit_conversion(self, product: Dict[str, Any]) -> Dict[str, Any]:
        """
        处理单位转换，根据单位和规格转换数量和单价
        
        Args:
            product: 商品字典，包含条码、单位、规格、数量和单价等字段
            
        Returns:
            处理后的商品字典
        """
        # 复制商品信息，避免修改原始数据
        result = product.copy()
        
        try:
            # 获取条码、单位、规格、数量和单价
            barcode = product.get('barcode', '')
            unit = product.get('unit', '')
            specification = product.get('specification', '')
            quantity = product.get('quantity', 0)
            price = product.get('price', 0)
            
            # 如果缺少关键信息，无法进行转换
            if not barcode or quantity == 0:
                return result
                
            # 1. 首先检查是否是特殊条码
            if barcode in self.special_barcodes:
                special_config = self.special_barcodes[barcode]
                logger.info(f"应用特殊条码配置: {barcode}, {special_config['description']}")
                
                # 应用乘数和单位转换
                result['quantity'] = quantity * special_config['multiplier']
                result['unit'] = special_config['target_unit']
                
                # 如果有单价，进行单价转换
                if price != 0:
                    result['price'] = price / special_config['multiplier']
                
                return result
            
            # 2. 提取规格包装数量
            package_quantity = None
            if specification:
                package_quantity = parse_specification(specification)
            
            # 3. 处理单位转换
            if unit and unit in self.special_units and package_quantity:
                # 判断是否是三级规格（1*5*12格式）
                is_three_level = bool(re.search(r'\d+[\*xX×]\d+[\*xX×]\d+', str(specification)))
                
                # 对于"提"和"盒"单位的特殊处理
                if (unit in ['提', '盒']) and not is_three_level:
                    # 二级规格：保持原数量不变
                    logger.info(f"二级规格的提/盒单位，保持原状: {unit}, 规格={specification}")
                    return result
                
                # 标准处理：数量×包装数量，单价÷包装数量
                logger.info(f"标准单位转换: {unit}->瓶, 规格={specification}, 包装数量={package_quantity}")
                result['quantity'] = quantity * package_quantity
                result['unit'] = '瓶'
                
                if price != 0:
                    result['price'] = price / package_quantity
                
                return result
            
            # 4. 默认返回原始数据
            return result
        
        except Exception as e:
            logger.error(f"单位转换处理出错: {e}")
            # 发生错误时，返回原始数据
            return result 