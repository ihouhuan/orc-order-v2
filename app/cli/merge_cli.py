"""
订单合并命令行工具
--------------
提供订单合并相关的命令行接口。
"""

import os
import sys
import argparse
from typing import List, Optional

from ..config.settings import ConfigManager
from ..core.utils.log_utils import get_logger, close_logger
from ..services.order_service import OrderService

logger = get_logger(__name__)

def create_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    
    Returns:
        参数解析器
    """
    parser = argparse.ArgumentParser(description='订单合并工具')
    
    # 通用选项
    parser.add_argument('--config', type=str, help='配置文件路径')
    
    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='子命令')
    
    # 合并命令
    merge_parser = subparsers.add_parser('merge', help='合并采购单')
    merge_parser.add_argument('--input', type=str, help='输入采购单文件路径列表，以逗号分隔，如果不指定则合并所有采购单')
    
    # 列出采购单命令
    list_parser = subparsers.add_parser('list', help='列出采购单文件')
    
    return parser

def merge_orders(order_service: OrderService, input_files: Optional[str] = None) -> bool:
    """
    合并采购单
    
    Args:
        order_service: 订单服务
        input_files: 输入文件路径列表，以逗号分隔，如果为None则合并所有采购单
        
    Returns:
        合并是否成功
    """
    if input_files:
        # 分割输入文件列表
        file_paths = [path.strip() for path in input_files.split(',')]
        
        # 检查文件是否存在
        for path in file_paths:
            if not os.path.exists(path):
                logger.error(f"输入文件不存在: {path}")
                return False
                
        result = order_service.merge_orders(file_paths)
    else:
        # 获取所有采购单文件
        file_paths = order_service.get_purchase_orders()
        if not file_paths:
            logger.warning("未找到采购单文件")
            return False
            
        logger.info(f"合并 {len(file_paths)} 个采购单文件")
        result = order_service.merge_orders()
    
    if result:
        logger.info(f"合并成功，输出文件: {result}")
        return True
    else:
        logger.error("合并失败")
        return False

def list_purchase_orders(order_service: OrderService) -> bool:
    """
    列出采购单文件
    
    Args:
        order_service: 订单服务
        
    Returns:
        是否有采购单文件
    """
    files = order_service.get_purchase_orders()
    
    if not files:
        logger.info("未找到采购单文件")
        return False
        
    logger.info(f"采购单文件 ({len(files)}):")
    for file in files:
        logger.info(f"  {file}")
    
    return True

def main(args: Optional[List[str]] = None) -> int:
    """
    订单合并命令行主函数
    
    Args:
        args: 命令行参数，如果为None则使用sys.argv
        
    Returns:
        退出状态码
    """
    parser = create_parser()
    parsed_args = parser.parse_args(args)
    
    if parsed_args.command is None:
        parser.print_help()
        return 1
    
    try:
        # 创建配置管理器
        config = ConfigManager(parsed_args.config) if parsed_args.config else ConfigManager()
        
        # 创建订单服务
        order_service = OrderService(config)
        
        # 根据命令执行不同功能
        if parsed_args.command == 'merge':
            success = merge_orders(order_service, parsed_args.input)
        elif parsed_args.command == 'list':
            success = list_purchase_orders(order_service)
        else:
            parser.print_help()
            return 1
            
        return 0 if success else 1
        
    except Exception as e:
        logger.error(f"执行过程中发生错误: {e}")
        return 1
        
    finally:
        # 关闭日志
        close_logger(__name__)

if __name__ == '__main__':
    sys.exit(main()) 