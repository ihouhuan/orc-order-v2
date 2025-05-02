"""
Excel处理命令行工具
---------------
提供Excel处理相关的命令行接口。
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
    parser = argparse.ArgumentParser(description='Excel处理工具')
    
    # 通用选项
    parser.add_argument('--config', type=str, help='配置文件路径')
    
    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='子命令')
    
    # 处理Excel命令
    process_parser = subparsers.add_parser('process', help='处理Excel文件')
    process_parser.add_argument('--input', type=str, help='输入Excel文件路径，如果不指定则处理最新的文件')
    
    # 查看命令
    list_parser = subparsers.add_parser('list', help='获取最新的Excel文件')
    
    return parser

def process_excel(order_service: OrderService, input_file: Optional[str] = None) -> bool:
    """
    处理Excel文件
    
    Args:
        order_service: 订单服务
        input_file: 输入文件路径，如果为None则处理最新的文件
        
    Returns:
        处理是否成功
    """
    if input_file:
        if not os.path.exists(input_file):
            logger.error(f"输入文件不存在: {input_file}")
            return False
            
        result = order_service.process_excel(input_file)
    else:
        latest_file = order_service.get_latest_excel()
        if not latest_file:
            logger.warning("未找到可处理的Excel文件")
            return False
            
        logger.info(f"处理最新的Excel文件: {latest_file}")
        result = order_service.process_excel(latest_file)
    
    if result:
        logger.info(f"处理成功，输出文件: {result}")
        return True
    else:
        logger.error("处理失败")
        return False

def list_latest_excel(order_service: OrderService) -> bool:
    """
    获取最新的Excel文件
    
    Args:
        order_service: 订单服务
        
    Returns:
        是否找到Excel文件
    """
    latest_file = order_service.get_latest_excel()
    
    if latest_file:
        logger.info(f"最新的Excel文件: {latest_file}")
        return True
    else:
        logger.info("未找到Excel文件")
        return False

def main(args: Optional[List[str]] = None) -> int:
    """
    Excel处理命令行主函数
    
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
        if parsed_args.command == 'process':
            success = process_excel(order_service, parsed_args.input)
        elif parsed_args.command == 'list':
            success = list_latest_excel(order_service)
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