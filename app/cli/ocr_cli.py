"""
OCR命令行工具
----------
提供OCR识别相关的命令行接口。
"""

import os
import sys
import argparse
from typing import List, Optional

from ..config.settings import ConfigManager
from ..core.utils.log_utils import get_logger, close_logger
from ..services.ocr_service import OCRService

logger = get_logger(__name__)

def create_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    
    Returns:
        参数解析器
    """
    parser = argparse.ArgumentParser(description='OCR识别工具')
    
    # 通用选项
    parser.add_argument('--config', type=str, help='配置文件路径')
    
    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='子命令')
    
    # 单文件处理命令
    process_parser = subparsers.add_parser('process', help='处理单个文件')
    process_parser.add_argument('--input', type=str, required=True, help='输入图片文件路径')
    
    # 批量处理命令
    batch_parser = subparsers.add_parser('batch', help='批量处理文件')
    batch_parser.add_argument('--batch-size', type=int, help='批处理大小')
    batch_parser.add_argument('--max-workers', type=int, help='最大线程数')
    
    # 查看未处理文件命令
    list_parser = subparsers.add_parser('list', help='列出未处理的文件')
    
    return parser

def process_file(ocr_service: OCRService, input_file: str) -> bool:
    """
    处理单个文件
    
    Args:
        ocr_service: OCR服务
        input_file: 输入文件路径
        
    Returns:
        处理是否成功
    """
    if not os.path.exists(input_file):
        logger.error(f"输入文件不存在: {input_file}")
        return False
        
    if not ocr_service.validate_image(input_file):
        logger.error(f"输入文件无效: {input_file}")
        return False
        
    result = ocr_service.process_image(input_file)
    
    if result:
        logger.info(f"处理成功，输出文件: {result}")
        return True
    else:
        logger.error("处理失败")
        return False

def process_batch(ocr_service: OCRService, batch_size: Optional[int] = None, max_workers: Optional[int] = None) -> bool:
    """
    批量处理文件
    
    Args:
        ocr_service: OCR服务
        batch_size: 批处理大小
        max_workers: 最大线程数
        
    Returns:
        处理是否成功
    """
    total, success = ocr_service.process_images_batch(batch_size, max_workers)
    
    if total == 0:
        logger.warning("没有找到需要处理的文件")
        return False
        
    logger.info(f"批量处理完成，总计: {total}，成功: {success}")
    return success > 0

def list_unprocessed(ocr_service: OCRService) -> bool:
    """
    列出未处理的文件
    
    Args:
        ocr_service: OCR服务
        
    Returns:
        是否有未处理的文件
    """
    files = ocr_service.get_unprocessed_images()
    
    if not files:
        logger.info("没有未处理的文件")
        return False
        
    logger.info(f"未处理的文件 ({len(files)}):")
    for file in files:
        logger.info(f"  {file}")
    
    return True

def main(args: Optional[List[str]] = None) -> int:
    """
    OCR命令行主函数
    
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
        
        # 创建OCR服务
        ocr_service = OCRService(config)
        
        # 根据命令执行不同功能
        if parsed_args.command == 'process':
            success = process_file(ocr_service, parsed_args.input)
        elif parsed_args.command == 'batch':
            success = process_batch(ocr_service, parsed_args.batch_size, parsed_args.max_workers)
        elif parsed_args.command == 'list':
            success = list_unprocessed(ocr_service)
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