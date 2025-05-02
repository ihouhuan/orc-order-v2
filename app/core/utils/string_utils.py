"""
字符串处理工具模块
---------------
提供字符串处理、正则表达式匹配等功能。
"""

import re
from typing import Dict, List, Optional, Tuple, Any, Match, Pattern

def clean_string(text: str) -> str:
    """
    清理字符串，移除多余空白
    
    Args:
        text: 源字符串
        
    Returns:
        清理后的字符串
    """
    if not isinstance(text, str):
        return ""
    
    # 移除首尾空白
    text = text.strip()
    # 移除多余空白
    text = re.sub(r'\s+', ' ', text)
    return text

def remove_non_digits(text: str) -> str:
    """
    移除字符串中的非数字字符
    
    Args:
        text: 源字符串
        
    Returns:
        只包含数字的字符串
    """
    if not isinstance(text, str):
        return ""
        
    return re.sub(r'\D', '', text)

def extract_number(text: str) -> Optional[float]:
    """
    从字符串中提取数字
    
    Args:
        text: 源字符串
        
    Returns:
        提取的数字，如果没有则返回None
    """
    if not isinstance(text, str):
        return None
        
    # 匹配数字（可以包含小数点和负号）
    match = re.search(r'-?\d+(\.\d+)?', text)
    if match:
        return float(match.group())
    return None

def extract_unit(text: str, units: List[str] = None) -> Optional[str]:
    """
    从字符串中提取单位
    
    Args:
        text: 源字符串
        units: 有效单位列表，如果为None则自动识别
        
    Returns:
        提取的单位，如果没有则返回None
    """
    if not isinstance(text, str):
        return None
        
    # 如果提供了单位列表，检查字符串中是否包含
    if units:
        for unit in units:
            if unit in text:
                return unit
        return None
        
    # 否则，尝试自动识别常见单位
    # 正则表达式：匹配数字后面的非数字部分作为单位
    match = re.search(r'\d+\s*([^\d\s]+)', text)
    if match:
        return match.group(1)
    return None

def extract_number_and_unit(text: str) -> Tuple[Optional[float], Optional[str]]:
    """
    从字符串中同时提取数字和单位
    
    Args:
        text: 源字符串
        
    Returns:
        (数字, 单位)元组，如果没有则对应返回None
    """
    if not isinstance(text, str):
        return None, None
        
    # 匹配数字和单位的组合
    match = re.search(r'(-?\d+(?:\.\d+)?)\s*([^\d\s]+)?', text)
    if match:
        number = float(match.group(1))
        unit = match.group(2) if match.group(2) else None
        return number, unit
    return None, None

def parse_specification(spec_str: str) -> Optional[int]:
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
            
    except Exception:
        pass
        
    return None

def clean_barcode(barcode: Any) -> str:
    """
    清理条码格式
    
    Args:
        barcode: 条码（可以是字符串、整数或浮点数）
        
    Returns:
        清理后的条码字符串
    """
    if isinstance(barcode, (int, float)):
        barcode = f"{barcode:.0f}"
        
    # 清理条码格式，移除可能的非数字字符（包括小数点）
    barcode_clean = re.sub(r'\.0+$', '', str(barcode))  # 移除末尾0
    barcode_clean = re.sub(r'\D', '', barcode_clean)  # 只保留数字
    
    return barcode_clean

def is_scientific_notation(value: str) -> bool:
    """
    检查字符串是否是科学计数法表示
    
    Args:
        value: 字符串值
        
    Returns:
        是否是科学计数法
    """
    return bool(re.match(r'^-?\d+(\.\d+)?[eE][+-]?\d+$', str(value)))

def format_barcode(barcode: Any) -> str:
    """
    格式化条码，处理科学计数法
    
    Args:
        barcode: 条码值
        
    Returns:
        格式化后的条码字符串
    """
    if isinstance(barcode, (int, float)) or is_scientific_notation(str(barcode)):
        try:
            # 转换为整数并格式化为字符串
            return f"{int(float(barcode))}"
        except (ValueError, TypeError):
            pass
    
    # 如果不是数字或转换失败，返回原始字符串
    return str(barcode) 