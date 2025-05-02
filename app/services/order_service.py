"""
订单服务模块
---------
提供订单处理服务，协调Excel处理和订单合并流程。
"""

from typing import Dict, List, Optional, Tuple, Union, Any

from ..config.settings import ConfigManager
from ..core.utils.log_utils import get_logger
from ..core.excel.processor import ExcelProcessor
from ..core.excel.merger import PurchaseOrderMerger

logger = get_logger(__name__)

class OrderService:
    """
    订单服务：协调Excel处理和订单合并流程
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化订单服务
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        logger.info("初始化OrderService")
        self.config = config or ConfigManager()
        
        # 创建Excel处理器和采购单合并器
        self.excel_processor = ExcelProcessor(self.config)
        self.order_merger = PurchaseOrderMerger(self.config)
        
        logger.info("OrderService初始化完成")
    
    def get_latest_excel(self) -> Optional[str]:
        """
        获取最新的Excel文件
        
        Returns:
            最新Excel文件路径，如果未找到则返回None
        """
        return self.excel_processor.get_latest_excel()
    
    def process_excel(self, file_path: Optional[str] = None) -> Optional[str]:
        """
        处理Excel文件，生成采购单
        
        Args:
            file_path: Excel文件路径，如果为None则处理最新的文件
            
        Returns:
            输出采购单文件路径，如果处理失败则返回None
        """
        if file_path:
            logger.info(f"OrderService开始处理指定Excel文件: {file_path}")
            return self.excel_processor.process_specific_file(file_path)
        else:
            logger.info("OrderService开始处理最新Excel文件")
            return self.excel_processor.process_latest_file()
    
    def get_purchase_orders(self) -> List[str]:
        """
        获取采购单文件列表
        
        Returns:
            采购单文件路径列表
        """
        return self.order_merger.get_purchase_orders()
    
    def merge_orders(self, file_paths: Optional[List[str]] = None) -> Optional[str]:
        """
        合并采购单
        
        Args:
            file_paths: 采购单文件路径列表，如果为None则处理所有采购单
            
        Returns:
            合并后的采购单文件路径，如果合并失败则返回None
        """
        if file_paths:
            logger.info(f"OrderService开始合并指定采购单: {file_paths}")
        else:
            logger.info("OrderService开始合并所有采购单")
        
        return self.order_merger.process(file_paths) 