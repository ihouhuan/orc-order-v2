#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
OCR订单处理系统 - 主入口
---------------------
提供命令行接口，整合OCR识别、Excel处理和订单合并功能。
"""

import os
import sys
import argparse
from typing import List, Optional

from app.config.settings import ConfigManager
from app.core.utils.log_utils import get_logger, close_logger
from app.services.ocr_service import OCRService
from app.services.order_service import OrderService

logger = get_logger(__name__)

def create_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    
    Returns:
        参数解析器
    """
    parser = argparse.ArgumentParser(description='OCR订单处理系统')
    
    # 通用选项
    parser.add_argument('--config', type=str, help='配置文件路径')
    
    # 子命令
    subparsers = parser.add_subparsers(dest='command', help='子命令')
    
    # OCR识别命令
    ocr_parser = subparsers.add_parser('ocr', help='OCR识别')
    ocr_parser.add_argument('--input', type=str, help='输入图片文件路径')
    ocr_parser.add_argument('--batch', action='store_true', help='批量处理模式')
    ocr_parser.add_argument('--batch-size', type=int, help='批处理大小')
    ocr_parser.add_argument('--max-workers', type=int, help='最大线程数')
    
    # Excel处理命令
    excel_parser = subparsers.add_parser('excel', help='Excel处理')
    excel_parser.add_argument('--input', type=str, help='输入Excel文件路径，如果不指定则处理最新的文件')
    
    # 订单合并命令
    merge_parser = subparsers.add_parser('merge', help='订单合并')
    merge_parser.add_argument('--input', type=str, help='输入采购单文件路径列表，以逗号分隔，如果不指定则合并所有采购单')
    
    # 完整流程命令
    pipeline_parser = subparsers.add_parser('pipeline', help='完整流程')
    pipeline_parser.add_argument('--input', type=str, help='输入图片文件路径，如果不指定则处理所有图片')
    
    return parser

def run_ocr(ocr_service: OCRService, args) -> bool:
    """
    运行OCR识别
    
    Args:
        ocr_service: OCR服务
        args: 命令行参数
        
    Returns:
        处理是否成功
    """
    if args.input:
        if not os.path.exists(args.input):
            logger.error(f"输入文件不存在: {args.input}")
            return False
            
        if not ocr_service.validate_image(args.input):
            logger.error(f"输入文件无效: {args.input}")
            return False
            
        logger.info(f"处理单个图片: {args.input}")
        result = ocr_service.process_image(args.input)
        
        if result:
            logger.info(f"OCR处理成功，输出文件: {result}")
            return True
        else:
            logger.error("OCR处理失败")
            return False
    elif args.batch:
        logger.info("批量处理模式")
        total, success = ocr_service.process_images_batch(args.batch_size, args.max_workers)
        
        if total == 0:
            logger.warning("没有找到需要处理的文件")
            return False
            
        logger.info(f"批量处理完成，总计: {total}，成功: {success}")
        return success > 0
    else:
        # 列出未处理的文件
        files = ocr_service.get_unprocessed_images()
        
        if not files:
            logger.info("没有未处理的文件")
            return True
            
        logger.info(f"未处理的文件 ({len(files)}):")
        for file in files:
            logger.info(f"  {file}")
        
        return True

def run_excel(order_service: OrderService, args) -> bool:
    """
    运行Excel处理
    
    Args:
        order_service: 订单服务
        args: 命令行参数
        
    Returns:
        处理是否成功
    """
    if args.input:
        if not os.path.exists(args.input):
            logger.error(f"输入文件不存在: {args.input}")
            return False
            
        logger.info(f"处理Excel文件: {args.input}")
        result = order_service.process_excel(args.input)
    else:
        latest_file = order_service.get_latest_excel()
        if not latest_file:
            logger.warning("未找到可处理的Excel文件")
            return False
            
        logger.info(f"处理最新的Excel文件: {latest_file}")
        result = order_service.process_excel(latest_file)
    
    if result:
        logger.info(f"Excel处理成功，输出文件: {result}")
        return True
    else:
        logger.error("Excel处理失败")
        return False

def run_merge(order_service: OrderService, args) -> bool:
    """
    运行订单合并
    
    Args:
        order_service: 订单服务
        args: 命令行参数
        
    Returns:
        处理是否成功
    """
    if args.input:
        # 分割输入文件列表
        file_paths = [path.strip() for path in args.input.split(',')]
        
        # 检查文件是否存在
        for path in file_paths:
            if not os.path.exists(path):
                logger.error(f"输入文件不存在: {path}")
                return False
                
        logger.info(f"合并指定的采购单文件: {file_paths}")
        result = order_service.merge_orders(file_paths)
    else:
        # 获取所有采购单文件
        file_paths = order_service.get_purchase_orders()
        if not file_paths:
            logger.warning("未找到采购单文件")
            return False
            
        logger.info(f"合并所有采购单文件: {len(file_paths)} 个")
        result = order_service.merge_orders()
    
    if result:
        logger.info(f"订单合并成功，输出文件: {result}")
        return True
    else:
        logger.error("订单合并失败")
        return False

def run_pipeline(ocr_service: OCRService, order_service: OrderService, args) -> bool:
    """
    运行完整流程
    
    Args:
        ocr_service: OCR服务
        order_service: 订单服务
        args: 命令行参数
        
    Returns:
        处理是否成功
    """
    # 1. OCR识别
    logger.info("=== 流程步骤 1: OCR识别 ===")
    
    if args.input:
        if not os.path.exists(args.input):
            logger.error(f"输入文件不存在: {args.input}")
            return False
            
        if not ocr_service.validate_image(args.input):
            logger.error(f"输入文件无效: {args.input}")
            return False
            
        logger.info(f"处理单个图片: {args.input}")
        ocr_result = ocr_service.process_image(args.input)
        
        if not ocr_result:
            logger.error("OCR处理失败")
            return False
            
        logger.info(f"OCR处理成功，输出文件: {ocr_result}")
    else:
        # 批量处理所有图片
        logger.info("批量处理所有图片")
        total, success = ocr_service.process_images_batch()
        
        if total == 0:
            logger.warning("没有找到需要处理的图片")
            # 继续下一步，因为可能已经有处理好的Excel文件
        elif success == 0:
            logger.error("OCR处理失败，没有成功处理的图片")
            return False
        else:
            logger.info(f"OCR处理完成，总计: {total}，成功: {success}")
    
    # 2. Excel处理
    logger.info("=== 流程步骤 2: Excel处理 ===")
    
    latest_file = order_service.get_latest_excel()
    if not latest_file:
        logger.warning("未找到可处理的Excel文件")
        return False
        
    logger.info(f"处理最新的Excel文件: {latest_file}")
    excel_result = order_service.process_excel(latest_file)
    
    if not excel_result:
        logger.error("Excel处理失败")
        return False
        
    logger.info(f"Excel处理成功，输出文件: {excel_result}")
    
    # 3. 订单合并
    logger.info("=== 流程步骤 3: 订单合并 ===")
    
    # 获取所有采购单文件
    file_paths = order_service.get_purchase_orders()
    if not file_paths:
        logger.warning("未找到采购单文件，跳过合并步骤")
        logger.info("=== 完整流程处理成功（未执行合并步骤）===")
        # 非错误状态，继续执行
        return True
        
    # 有文件需要合并
    logger.info(f"发现 {len(file_paths)} 个采购单文件")
    
    if len(file_paths) == 1:
        logger.warning(f"只有1个采购单文件 {file_paths[0]}，无需合并")
        logger.info("=== 完整流程处理成功（只有一个文件，跳过合并）===")
        return True
        
    logger.info(f"合并所有采购单文件: {len(file_paths)} 个")
    merge_result = order_service.merge_orders()
    
    if not merge_result:
        logger.error("订单合并失败")
        return False
        
    logger.info(f"订单合并成功，输出文件: {merge_result}")
    
    logger.info("=== 完整流程处理成功 ===")
    return True

def main(args: Optional[List[str]] = None) -> int:
    """
    主函数
    
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
        
        # 创建服务
        ocr_service = OCRService(config)
        order_service = OrderService(config)
        
        # 根据命令执行不同功能
        if parsed_args.command == 'ocr':
            success = run_ocr(ocr_service, parsed_args)
        elif parsed_args.command == 'excel':
            success = run_excel(order_service, parsed_args)
        elif parsed_args.command == 'merge':
            success = run_merge(order_service, parsed_args)
        elif parsed_args.command == 'pipeline':
            success = run_pipeline(ocr_service, order_service, parsed_args)
        else:
            parser.print_help()
            return 1
            
        return 0 if success else 1
        
    except Exception as e:
        logger.error(f"执行过程中发生错误: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1
        
    finally:
        # 关闭日志
        close_logger(__name__)

if __name__ == '__main__':
    sys.exit(main()) 