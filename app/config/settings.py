"""
配置管理模块
-----------
提供统一的配置加载、访问和保存功能。
"""

import os
import configparser
import logging
from typing import Dict, List, Optional, Any

from .defaults import DEFAULT_CONFIG

logger = logging.getLogger(__name__)

class ConfigManager:
    """
    配置管理类，负责加载和保存配置
    单例模式确保全局只有一个配置实例
    """
    _instance = None
    
    def __new__(cls, config_file=None):
        """单例模式实现"""
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
            cls._instance._init(config_file)
        return cls._instance
    
    def _init(self, config_file):
        """初始化配置管理器"""
        self.config_file = config_file or 'config.ini'
        self.config = configparser.ConfigParser()
        self.load_config()
    
    def load_config(self) -> None:
        """
        加载配置文件，如果不存在则创建默认配置
        """
        if not os.path.exists(self.config_file):
            self.create_default_config()
        
        try:
            self.config.read(self.config_file, encoding='utf-8')
            logger.info(f"已加载配置文件: {self.config_file}")
        except Exception as e:
            logger.error(f"加载配置文件时出错: {e}")
            logger.info("使用默认配置")
            self.create_default_config(save=False)
    
    def create_default_config(self, save: bool = True) -> None:
        """创建默认配置"""
        for section, options in DEFAULT_CONFIG.items():
            if not self.config.has_section(section):
                self.config.add_section(section)
            
            for option, value in options.items():
                self.config.set(section, option, value)
        
        if save:
            self.save_config()
            logger.info(f"已创建默认配置文件: {self.config_file}")
    
    def save_config(self) -> None:
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            logger.error(f"保存配置文件时出错: {e}")
    
    def get(self, section: str, option: str, fallback: Any = None) -> Any:
        """获取配置值"""
        return self.config.get(section, option, fallback=fallback)
    
    def getint(self, section: str, option: str, fallback: int = 0) -> int:
        """获取整数配置值"""
        return self.config.getint(section, option, fallback=fallback)
    
    def getfloat(self, section: str, option: str, fallback: float = 0.0) -> float:
        """获取浮点数配置值"""
        return self.config.getfloat(section, option, fallback=fallback)
    
    def getboolean(self, section: str, option: str, fallback: bool = False) -> bool:
        """获取布尔配置值"""
        return self.config.getboolean(section, option, fallback=fallback)
    
    def get_list(self, section: str, option: str, fallback: str = "", delimiter: str = ",") -> List[str]:
        """获取列表配置值（逗号分隔的字符串转为列表）"""
        value = self.get(section, option, fallback)
        return [item.strip() for item in value.split(delimiter) if item.strip()]

    def update(self, section: str, option: str, value: Any) -> None:
        """更新配置选项"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        
        self.config.set(section, option, str(value))
        logger.debug(f"更新配置: [{section}] {option} = {value}")
    
    def get_path(self, section: str, option: str, fallback: str = "", create: bool = False) -> str:
        """
        获取路径配置并确保它是一个有效的绝对路径
        如果create为True，则自动创建该目录
        """
        path = self.get(section, option, fallback)
        
        if not os.path.isabs(path):
            # 相对路径，转为绝对路径
            path = os.path.abspath(path)
        
        if create and not os.path.exists(path):
            try:
                # 如果是文件路径，创建其父目录
                if '.' in os.path.basename(path):
                    directory = os.path.dirname(path)
                    if directory and not os.path.exists(directory):
                        os.makedirs(directory, exist_ok=True)
                        logger.info(f"已创建目录: {directory}")
                else:
                    # 否则认为是目录路径
                    os.makedirs(path, exist_ok=True)
                    logger.info(f"已创建目录: {path}")
            except Exception as e:
                logger.error(f"创建目录失败: {path}, 错误: {e}")
        
        return path 