"""
日志工具模块
----------
提供统一的日志配置和管理功能。
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict

# 日志处理器字典，用于跟踪已创建的处理器
_handlers: Dict[str, logging.Handler] = {}

def setup_logger(name: str, 
                log_file: Optional[str] = None, 
                level=logging.INFO, 
                console_output: bool = True,
                file_output: bool = True,
                log_format: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s') -> logging.Logger:
    """
    配置并返回日志记录器
    
    Args:
        name: 日志记录器的名称
        log_file: 日志文件路径，如果为None则使用默认路径
        level: 日志级别
        console_output: 是否输出到控制台
        file_output: 是否输出到文件
        log_format: 日志格式
        
    Returns:
        配置好的日志记录器
    """
    # 获取或创建日志记录器
    logger = logging.getLogger(name)
    
    # 如果已经配置过处理器，不重复配置
    if logger.handlers:
        return logger
    
    # 设置日志级别
    logger.setLevel(level)
    
    # 创建格式化器
    formatter = logging.Formatter(log_format)
    
    # 如果需要输出到文件
    if file_output:
        # 如果没有指定日志文件，使用默认路径
        if log_file is None:
            log_dir = os.path.abspath('logs')
            # 确保日志目录存在
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"{name}.log")
        
        # 创建文件处理器
        try:
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            file_handler.setLevel(level)
            logger.addHandler(file_handler)
            _handlers[f"{name}_file"] = file_handler
            
            # 记录活跃标记，避免被日志清理工具删除
            active_marker = os.path.join(os.path.dirname(log_file), f"{name}.active")
            with open(active_marker, 'w', encoding='utf-8') as f:
                f.write(f"Active since: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as e:
            print(f"无法创建日志文件处理器: {e}")
    
    # 如果需要输出到控制台
    if console_output:
        # 创建控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        console_handler.setLevel(level)
        logger.addHandler(console_handler)
        _handlers[f"{name}_console"] = console_handler
    
    return logger

def get_logger(name: str) -> logging.Logger:
    """
    获取已配置的日志记录器，如果不存在则创建一个新的
    
    Args:
        name: 日志记录器的名称
        
    Returns:
        日志记录器
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        return setup_logger(name)
    return logger

def close_logger(name: str) -> None:
    """
    关闭日志记录器的所有处理器
    
    Args:
        name: 日志记录器的名称
    """
    logger = logging.getLogger(name)
    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)
    
    # 清除处理器缓存
    _handlers.pop(f"{name}_file", None)
    _handlers.pop(f"{name}_console", None)

def cleanup_active_marker(name: str) -> None:
    """
    清理日志活跃标记
    
    Args:
        name: 日志记录器的名称
    """
    try:
        log_dir = os.path.abspath('logs')
        active_marker = os.path.join(log_dir, f"{name}.active")
        if os.path.exists(active_marker):
            os.remove(active_marker)
    except Exception as e:
        print(f"无法清理日志活跃标记: {e}") 