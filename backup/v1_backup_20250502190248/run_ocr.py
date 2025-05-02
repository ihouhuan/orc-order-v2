#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
OCR流程运行脚本
-------------
整合百度OCR和Excel处理功能的便捷脚本
"""

import os
import sys
import argparse
import logging
import configparser
from pathlib import Path
from datetime import datetime

# 确保logs目录存在
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# 设置日志文件路径
log_file = os.path.join(log_dir, 'ocr_processor.log')

# 配置日志
logger = logging.getLogger('ocr_processor')
if not logger.handlers:
    # 创建文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # 设置格式
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器到日志器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    # 设置日志级别
    logger.setLevel(logging.INFO)

logger.info("OCR处理器初始化")

# 标记该日志文件为活跃，避免被清理工具删除
try:
    # 创建一个标记文件，表示该日志文件正在使用中
    active_marker = os.path.join(log_dir, 'ocr_processor.active')
    with open(active_marker, 'w') as f:
        f.write(f"Active since: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
except Exception as e:
    logger.warning(f"无法创建日志活跃标记: {e}")

def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='OCR流程运行脚本')
    parser.add_argument('--step', type=int, default=0, help='运行步骤: 1-OCR识别, 2-Excel处理, 0-全部运行 (默认)')
    parser.add_argument('--config', type=str, default='config.ini', help='配置文件路径')
    parser.add_argument('--force', action='store_true', help='强制处理所有文件，包括已处理的文件')
    parser.add_argument('--input', type=str, help='指定输入文件（仅用于单文件处理）')
    parser.add_argument('--output', type=str, help='指定输出文件（仅用于单文件处理）')
    return parser.parse_args()

def check_env():
    """检查配置是否有效"""
    try:
        # 尝试读取配置文件
        config = configparser.ConfigParser()
        if not config.read('config.ini', encoding='utf-8'):
            logger.warning("未找到配置文件config.ini或文件为空")
            return
            
        # 检查API密钥是否已配置
        if not config.has_section('API'):
            logger.warning("配置文件中缺少[API]部分")
            return
            
        api_key = config.get('API', 'api_key', fallback='')
        secret_key = config.get('API', 'secret_key', fallback='')
        
        if not api_key or not secret_key:
            logger.warning("API密钥未设置或为空，请在config.ini中配置API密钥")
        
    except Exception as e:
        logger.error(f"检查配置时出错: {e}")

def run_ocr(args):
    """运行OCR识别过程"""
    logger.info("开始OCR识别过程...")
    
    # 导入模块
    try:
        from baidu_table_ocr import OCRProcessor, ConfigManager
        
        # 创建配置管理器
        config_manager = ConfigManager(args.config)
        
        # 创建处理器
        processor = OCRProcessor(config_manager)
        
        # 检查输入目录中是否有图片
        input_files = processor.get_unprocessed_images()
        if not input_files and not args.input:
            logger.warning(f"在{processor.input_folder}目录中没有找到未处理的图片文件")
            return False
        
        # 单文件处理或批量处理
        if args.input:
            if not os.path.exists(args.input):
                logger.error(f"输入文件不存在: {args.input}")
                return False
            
            logger.info(f"处理单个文件: {args.input}")
            output_file = processor.process_image(args.input)
            if output_file:
                logger.info(f"OCR识别成功，输出文件: {output_file}")
                return True
            else:
                logger.error("OCR识别失败")
                return False
        else:
            # 批量处理
            batch_size = processor.batch_size
            max_workers = processor.max_workers
            
            # 如果需要强制处理，先设置skip_existing为False
            if args.force:
                processor.skip_existing = False
            
            logger.info(f"批量处理文件，批量大小: {batch_size}, 最大线程数: {max_workers}")
            total, success = processor.process_images_batch(
                batch_size=batch_size,
                max_workers=max_workers
            )
            
            logger.info(f"OCR识别完成，总计处理: {total}，成功: {success}")
            return success > 0
    
    except ImportError as e:
        logger.error(f"导入OCR模块失败: {e}")
        return False
    except Exception as e:
        logger.error(f"OCR识别过程出错: {e}")
        return False

def run_excel_processing(args):
    """运行Excel处理过程"""
    logger.info("开始Excel处理过程...")
    
    # 导入模块
    try:
        from excel_processor_step2 import ExcelProcessorStep2
        
        # 创建处理器
        processor = ExcelProcessorStep2()
        
        # 单文件处理或批量处理
        if args.input:
            if not os.path.exists(args.input):
                logger.error(f"输入文件不存在: {args.input}")
                return False
            
            logger.info(f"处理单个Excel文件: {args.input}")
            result = processor.process_specific_file(args.input)
            if result:
                logger.info(f"Excel处理成功")
                return True
            else:
                logger.error("Excel处理失败，请查看日志了解详细信息")
                return False
        else:
            # 检查output目录中最新的Excel文件
            latest_file = processor.get_latest_excel()
            if not latest_file:
                logger.error("未找到可处理的Excel文件，无法进行处理")
                return False
                
            # 处理最新的Excel文件
            logger.info(f"处理最新的Excel文件: {latest_file}")
            result = processor.process_latest_file()
            
            if result:
                logger.info("Excel处理成功")
                return True
            else:
                logger.error("Excel处理失败，请查看日志了解详细信息")
                return False
    
    except ImportError as e:
        logger.error(f"导入Excel处理模块失败: {e}")
        return False
    except Exception as e:
        logger.error(f"Excel处理过程出错: {e}")
        return False

def main():
    """主函数"""
    # 解析命令行参数
    args = parse_args()
    
    # 检查环境变量
    check_env()
    
    # 根据步骤运行相应的处理
    ocr_success = False
    
    if args.step == 0 or args.step == 1:
        ocr_success = run_ocr(args)
        if not ocr_success:
            if args.step == 1:
                logger.error("OCR识别失败，请检查input目录是否有图片或检查API配置")
                sys.exit(1)
            else:
                logger.warning("OCR识别未处理任何文件，跳过Excel处理步骤")
                return
    else:
        # 如果只运行第二步，假设OCR已成功完成
        ocr_success = True
    
    # 只有当OCR成功或只运行第二步时才执行Excel处理
    if ocr_success and (args.step == 0 or args.step == 2):
        excel_result = run_excel_processing(args)
        if not excel_result and args.step == 2:
            logger.error("Excel处理失败")
            sys.exit(1)
    
    logger.info("处理完成")

if __name__ == "__main__":
    main() 