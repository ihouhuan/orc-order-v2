"""
文件操作工具模块
--------------
提供文件处理、查找和管理功能。
"""

import os
import sys
import shutil
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Union, Any

from .log_utils import get_logger

logger = get_logger(__name__)

def ensure_dir(directory: str) -> bool:
    """
    确保目录存在，如果不存在则创建
    
    Args:
        directory: 目录路径
        
    Returns:
        是否成功创建或目录已存在
    """
    try:
        os.makedirs(directory, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"创建目录失败: {directory}, 错误: {e}")
        return False

def get_file_extension(file_path: str) -> str:
    """
    获取文件扩展名（小写）
    
    Args:
        file_path: 文件路径
        
    Returns:
        文件扩展名，包含点（例如 .jpg）
    """
    return os.path.splitext(file_path)[1].lower()

def is_valid_extension(file_path: str, allowed_extensions: List[str]) -> bool:
    """
    检查文件扩展名是否在允许的列表中
    
    Args:
        file_path: 文件路径
        allowed_extensions: 允许的扩展名列表（例如 ['.jpg', '.png']）
        
    Returns:
        文件扩展名是否有效
    """
    ext = get_file_extension(file_path)
    return ext in allowed_extensions

def get_files_by_extensions(directory: str, extensions: List[str], exclude_patterns: List[str] = None) -> List[str]:
    """
    获取指定目录下所有符合扩展名的文件路径
    
    Args:
        directory: 目录路径
        extensions: 扩展名列表（例如 ['.jpg', '.png']）
        exclude_patterns: 排除的文件名模式（例如 ['~$', '.tmp']）
        
    Returns:
        文件路径列表
    """
    if exclude_patterns is None:
        exclude_patterns = ['~$', '.tmp']
        
    files = []
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        
        # 检查是否是文件
        if not os.path.isfile(file_path):
            continue
            
        # 检查扩展名
        if not is_valid_extension(file_path, extensions):
            continue
            
        # 检查排除模式
        exclude = False
        for pattern in exclude_patterns:
            if pattern in file:
                exclude = True
                break
                
        if not exclude:
            files.append(file_path)
            
    return files

def get_latest_file(directory: str, pattern: str = "", extensions: List[str] = None) -> Optional[str]:
    """
    获取指定目录下最新的文件
    
    Args:
        directory: 目录路径
        pattern: 文件名包含的字符串模式
        extensions: 限制的文件扩展名列表
        
    Returns:
        最新文件的路径，如果没有找到则返回None
    """
    if not os.path.exists(directory):
        logger.warning(f"目录不存在: {directory}")
        return None
        
    files = []
    for file in os.listdir(directory):
        # 检查模式和扩展名
        if (pattern and pattern not in file) or \
           (extensions and not is_valid_extension(file, extensions)):
            continue
            
        file_path = os.path.join(directory, file)
        if os.path.isfile(file_path):
            files.append((file_path, os.path.getmtime(file_path)))
    
    if not files:
        logger.warning(f"未在目录 {directory} 中找到符合条件的文件")
        return None
    
    # 按修改时间排序，返回最新的
    sorted_files = sorted(files, key=lambda x: x[1], reverse=True)
    return sorted_files[0][0]

def generate_timestamp_filename(original_path: str) -> str:
    """
    生成基于时间戳的文件名
    
    Args:
        original_path: 原始文件路径
        
    Returns:
        带时间戳的新文件路径
    """
    dir_path = os.path.dirname(original_path)
    ext = os.path.splitext(original_path)[1]
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    return os.path.join(dir_path, f"{timestamp}{ext}")

def rename_file(source_path: str, target_path: str) -> bool:
    """
    重命名文件
    
    Args:
        source_path: 源文件路径
        target_path: 目标文件路径
        
    Returns:
        是否成功重命名
    """
    try:
        # 确保目标目录存在
        target_dir = os.path.dirname(target_path)
        ensure_dir(target_dir)
        
        # 重命名文件
        os.rename(source_path, target_path)
        logger.info(f"文件已重命名: {os.path.basename(source_path)} -> {os.path.basename(target_path)}")
        return True
    except Exception as e:
        logger.error(f"重命名文件失败: {e}")
        return False

def load_json(file_path: str, default: Any = None) -> Any:
    """
    加载JSON文件
    
    Args:
        file_path: JSON文件路径
        default: 如果文件不存在或加载失败时返回的默认值
        
    Returns:
        JSON内容，或者默认值
    """
    if not os.path.exists(file_path):
        return default
        
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"加载JSON文件失败: {file_path}, 错误: {e}")
        return default

def save_json(data: Any, file_path: str, ensure_ascii: bool = False, indent: int = 2) -> bool:
    """
    保存数据到JSON文件
    
    Args:
        data: 要保存的数据
        file_path: JSON文件路径
        ensure_ascii: 是否确保ASCII编码
        indent: 缩进空格数
        
    Returns:
        是否成功保存
    """
    try:
        # 确保目录存在
        directory = os.path.dirname(file_path)
        ensure_dir(directory)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)
        logger.debug(f"JSON数据已保存到: {file_path}")
        return True
    except Exception as e:
        logger.error(f"保存JSON文件失败: {file_path}, 错误: {e}")
        return False

def get_file_size(file_path: str) -> int:
    """
    获取文件大小（字节）
    
    Args:
        file_path: 文件路径
        
    Returns:
        文件大小（字节）
    """
    try:
        return os.path.getsize(file_path)
    except Exception as e:
        logger.error(f"获取文件大小失败: {file_path}, 错误: {e}")
        return 0

def is_file_size_valid(file_path: str, max_size_mb: float) -> bool:
    """
    检查文件大小是否在允许范围内
    
    Args:
        file_path: 文件路径
        max_size_mb: 最大允许大小（MB）
        
    Returns:
        文件大小是否有效
    """
    size_bytes = get_file_size(file_path)
    max_size_bytes = max_size_mb * 1024 * 1024
    return size_bytes <= max_size_bytes 