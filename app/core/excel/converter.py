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
            },
            '6921168593804': {
                'multiplier': 30,  # 数量乘以30
                'target_unit': '瓶',  # 目标单位
                'description': 'NFC产品特殊处理：每箱30瓶'
            },
            '6901826888138': {
                'multiplier': 30,  # 数量乘以30
                'target_unit': '瓶',  # 目标单位
                'fixed_price': 112/30,  # 固定单价为112/30
                'specification': '1*30',  # 固定规格
                'description': '特殊处理: 规格1*30，数量*30，单价=112/30'
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
        
        支持的格式:
        1. "2箱" -> (2, "箱")
        2. "3件" -> (3, "件")
        3. "1.5提" -> (1.5, "提")
        4. "数量: 5盒" -> (5, "盒")
        5. "× 2瓶" -> (2, "瓶")
        
        Args:
            quantity_str: 数量字符串，如"2箱"、"5件"
            
        Returns:
            (数量, 单位)的元组，如果无法提取则返回(None, None)
        """
        if not quantity_str or not isinstance(quantity_str, str):
            return None, None
        
        # 清理字符串，移除前后空白和一些常见前缀
        cleaned_str = quantity_str.strip()
        for prefix in ['数量:', '数量：', '×', 'x', 'X', '*']:
            cleaned_str = cleaned_str.replace(prefix, '').strip()
        
        # 匹配数字+单位格式 (基本格式)
        basic_match = re.match(r'^([\d\.]+)\s*([^\d\s\.]+)$', cleaned_str)
        if basic_match:
            try:
                num = float(basic_match.group(1))
                unit = basic_match.group(2)
                logger.info(f"从数量提取单位(基本格式): {quantity_str} -> 数量={num}, 单位={unit}")
                return num, unit
            except ValueError:
                pass
        
        # 匹配更复杂的格式，如包含其他文本的情况
        complex_match = re.search(r'([\d\.]+)\s*([箱|件|瓶|提|盒|袋|桶|包|kg|g|升|毫升|L|ml|个])', cleaned_str)
        if complex_match:
            try:
                num = float(complex_match.group(1))
                unit = complex_match.group(2)
                logger.info(f"从数量提取单位(复杂格式): {quantity_str} -> 数量={num}, 单位={unit}")
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
        
        # 处理XX入白膜格式，如"550纯净水24入白膜"
        match = re.search(r'.*?(\d+)入白膜', text)
        if match:
            result = f"1*{match.group(1)}"
            logger.info(f"提取规格(入白膜): {text} -> {result}")
            return result
            
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
        
        规则:
        1. "xx入纸箱" -> 1*xx (如"15入纸箱" -> 1*15)
        2. 直接包含规格 "1*15" -> 1*15
        3. "xx纸箱" -> 1*xx (如"15纸箱" -> 1*15)
        4. "xx白膜" -> 1*xx (如"12白膜" -> 1*12)
        5. "xxL" 容量单位特殊处理
        6. "xx(g|ml|毫升|克)*数字" -> 1*数字 (如"450g*15" -> 1*15)
        
        Args:
            name: 商品名称
            
        Returns:
            推断的规格，如果无法推断则返回None
        """
        if not name or not isinstance(name, str):
            return None
        
        # 记录原始商品名称，用于日志
        original_name = name
        
        # 新增模式: 处理重量/容量*数字格式，如"450g*15", "450ml*15"
        # 忽略重量/容量值，只提取后面的数量作为规格
        weight_volume_pattern = r'.*?\d+(?:g|ml|毫升|克)[*xX×](\d+)'
        match = re.search(weight_volume_pattern, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(重量/容量*数量): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式1.1: "xx入白膜" 格式，如"550纯净水24入白膜" -> "1*24"
        pattern1_1 = r'.*?(\d+)入白膜'
        match = re.search(pattern1_1, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(入白膜): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式1: "xx入纸箱" 格式，如"445水溶C血橙15入纸箱" -> "1*15"
        pattern1 = r'.*?(\d+)入纸箱'
        match = re.search(pattern1, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(入纸箱): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式2: 直接包含规格，如"500-东方树叶-乌龙茶1*15-纸箱装" -> "1*15"
        pattern2 = r'.*?(\d+)[*xX×](\d+).*'
        match = re.search(pattern2, name)
        if match:
            inferred_spec = f"{match.group(1)}*{match.group(2)}"
            logger.info(f"从名称推断规格(直接格式): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式3: "xx纸箱" 格式，如"500茶π蜜桃乌龙15纸箱" -> "1*15"
        pattern3 = r'.*?(\d+)纸箱'
        match = re.search(pattern3, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(纸箱): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式4: "xx白膜" 格式，如"1.5L水12白膜" 或 "550水24白膜" -> "1*12" 或 "1*24"
        pattern4 = r'.*?(\d+)白膜'
        match = re.search(pattern4, name)
        if match:
            inferred_spec = f"1*{match.group(1)}"
            logger.info(f"从名称推断规格(白膜): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 特殊模式5: 容量单位带数量格式 "1.8L*8瓶" -> "1.8L*8"
        volume_count_pattern = r'.*?([\d\.]+)[Ll升][*×xX](\d+).*'
        match = re.search(volume_count_pattern, name)
        if match:
            volume = match.group(1)
            count = match.group(2)
            inferred_spec = f"{volume}L*{count}"
            logger.info(f"从名称推断规格(容量*数量): {original_name} -> {inferred_spec}")
            return inferred_spec
            
        # 特殊模式6: 简单容量单位如"12.9L桶装水" -> "12.9L*1"
        simple_volume_pattern = r'.*?([\d\.]+)[Ll升].*'
        match = re.search(simple_volume_pattern, name)
        if match:
            inferred_spec = f"{match.group(1)}L*1"
            logger.info(f"从名称推断规格(简单容量): {original_name} -> {inferred_spec}")
            return inferred_spec
        
        # 尝试通用模式匹配
        spec = self.extract_specification(name)
        if spec:
            logger.info(f"从名称推断规格(通用模式): {original_name} -> {spec}")
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
            
        try:
            # 清理规格字符串，确保格式统一
            spec = re.sub(r'\s+', '', spec)  # 移除所有空白
            spec = re.sub(r'[xX×]', '*', spec)  # 统一分隔符为*
            
            logger.debug(f"解析规格: {spec}")
            
            # 处理三级包装，如1*5*12
            three_level_match = re.match(r'(\d+)[*](\d+)[*](\d+)', spec)
            if three_level_match:
                try:
                    level1 = int(three_level_match.group(1))
                    level2 = int(three_level_match.group(2))
                    level3 = int(three_level_match.group(3))
                    logger.info(f"解析三级规格: {spec} -> {level1}*{level2}*{level3}")
                    return level1, level2, level3
                except ValueError:
                    pass
            
            # 处理带容量单位的规格，如500ml*15, 1L*12等
            ml_match = re.match(r'(\d+)(?:ml|毫升)[*](\d+)', spec, re.IGNORECASE)
            if ml_match:
                try:
                    # 对于ml单位，使用1作为一级包装，后面的数字作为二级包装
                    level2 = int(ml_match.group(2))
                    logger.info(f"解析容量(ml)规格: {spec} -> 1*{level2}")
                    return 1, level2, None
                except ValueError:
                    pass
            
            # 处理带L单位的规格，如1L*12等
            l_match = re.match(r'(\d+(?:\.\d+)?)[Ll升][*](\d+)', spec)
            if l_match:
                try:
                    # 对于L单位，正确提取第二部分作为包装数量
                    level2 = int(l_match.group(2))
                    logger.info(f"解析容量(L)规格: {spec} -> 1*{level2}")
                    return 1, level2, None
                except ValueError:
                    pass
            
            # 处理二级包装，如1*12
            two_level_match = re.match(r'(\d+)[*](\d+)', spec)
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
        except Exception as e:
            logger.error(f"解析规格时出错: {e}")
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
            
            # 如果有固定单价，优先使用
            if 'fixed_price' in special_config:
                new_price = special_config['fixed_price']
                logger.info(f"特殊条码({barcode})使用固定单价: {new_price}")
            
            # 如果有固定规格，设置规格
            if 'specification' in special_config:
                result['specification'] = special_config['specification']
                # 解析规格以获取包装数量
                package_quantity = self.parse_specification(special_config['specification'])
                if package_quantity:
                    result['package_quantity'] = package_quantity
                logger.info(f"特殊条码({barcode})使用固定规格: {special_config['specification']}, 包装数量={package_quantity}")
            
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