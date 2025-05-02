"""
表格OCR处理模块
-------------
处理图片并提取表格内容，保存为Excel文件。
"""

import os
import sys
import time
import json
import base64
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional, Tuple, Union, Any

from ...config.settings import ConfigManager
from ..utils.log_utils import get_logger
from ..utils.file_utils import (
    ensure_dir, 
    get_file_extension, 
    get_files_by_extensions, 
    generate_timestamp_filename,
    is_file_size_valid,
    load_json,
    save_json
)
from .baidu_ocr import BaiduOCRClient

logger = get_logger(__name__)

class ProcessedRecordManager:
    """处理记录管理器，用于跟踪已处理的文件"""
    
    def __init__(self, record_file: str):
        """
        初始化处理记录管理器
        
        Args:
            record_file: 记录文件路径
        """
        self.record_file = record_file
        self.processed_files = self._load_record()
    
    def _load_record(self) -> Dict[str, str]:
        """
        加载处理记录
        
        Returns:
            处理记录字典，键为输入文件路径，值为输出文件路径
        """
        return load_json(self.record_file, {})
    
    def save_record(self) -> None:
        """保存处理记录"""
        save_json(self.processed_files, self.record_file)
    
    def is_processed(self, image_file: str) -> bool:
        """
        检查图片是否已处理
        
        Args:
            image_file: 图片文件路径
            
        Returns:
            是否已处理
        """
        return image_file in self.processed_files
    
    def mark_as_processed(self, image_file: str, output_file: str) -> None:
        """
        标记图片为已处理
        
        Args:
            image_file: 图片文件路径
            output_file: 输出文件路径
        """
        self.processed_files[image_file] = output_file
        self.save_record()
    
    def get_output_file(self, image_file: str) -> Optional[str]:
        """
        获取图片的输出文件路径
        
        Args:
            image_file: 图片文件路径
            
        Returns:
            输出文件路径，如果不存在则返回None
        """
        return self.processed_files.get(image_file)
    
    def get_unprocessed_files(self, files: List[str]) -> List[str]:
        """
        获取未处理的文件列表
        
        Args:
            files: 文件列表
            
        Returns:
            未处理的文件列表
        """
        return [file for file in files if not self.is_processed(file)]

class OCRProcessor:
    """
    OCR处理器，用于表格识别与处理
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化OCR处理器
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        self.config = config or ConfigManager()
        
        # 创建百度OCR客户端
        self.ocr_client = BaiduOCRClient(self.config)
        
        # 获取配置
        self.input_folder = self.config.get_path('Paths', 'input_folder', 'data/input', create=True)
        self.output_folder = self.config.get_path('Paths', 'output_folder', 'data/output', create=True)
        self.temp_folder = self.config.get_path('Paths', 'temp_folder', 'data/temp', create=True)
        
        # 确保目录结构正确
        for folder in [self.input_folder, self.output_folder, self.temp_folder]:
            if not os.path.exists(folder):
                os.makedirs(folder, exist_ok=True)
                logger.info(f"创建目录: {folder}")
        
        # 记录实际路径
        logger.info(f"使用输入目录: {os.path.abspath(self.input_folder)}")
        logger.info(f"使用输出目录: {os.path.abspath(self.output_folder)}")
        logger.info(f"使用临时目录: {os.path.abspath(self.temp_folder)}")
        
        self.allowed_extensions = self.config.get_list('File', 'allowed_extensions', '.jpg,.jpeg,.png,.bmp')
        self.max_file_size_mb = self.config.getfloat('File', 'max_file_size_mb', 4.0)
        self.excel_extension = self.config.get('File', 'excel_extension', '.xlsx')
        
        # 处理性能配置
        self.max_workers = self.config.getint('Performance', 'max_workers', 4)
        self.batch_size = self.config.getint('Performance', 'batch_size', 5)
        self.skip_existing = self.config.getboolean('Performance', 'skip_existing', True)
        
        # 初始化处理记录管理器
        record_file = self.config.get('Paths', 'processed_record', 'data/processed_files.json')
        self.record_manager = ProcessedRecordManager(record_file)
        
        logger.info(f"OCR处理器初始化完成，输入目录: {self.input_folder}, 输出目录: {self.output_folder}")
    
    def get_unprocessed_images(self) -> List[str]:
        """
        获取未处理的图片列表
        
        Returns:
            未处理的图片文件路径列表
        """
        # 获取所有图片文件
        image_files = get_files_by_extensions(self.input_folder, self.allowed_extensions)
        
        # 如果需要跳过已存在的文件
        if self.skip_existing:
            # 过滤已处理的文件
            unprocessed_files = self.record_manager.get_unprocessed_files(image_files)
            logger.info(f"找到 {len(image_files)} 个图片文件，其中 {len(unprocessed_files)} 个未处理")
            return unprocessed_files
        
        logger.info(f"找到 {len(image_files)} 个图片文件（不跳过已处理的文件）")
        return image_files
    
    def validate_image(self, image_path: str) -> bool:
        """
        验证图片是否有效
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            图片是否有效
        """
        # 检查文件是否存在
        if not os.path.exists(image_path):
            logger.warning(f"图片文件不存在: {image_path}")
            return False
        
        # 检查文件扩展名
        ext = get_file_extension(image_path)
        if ext not in self.allowed_extensions:
            logger.warning(f"不支持的文件类型: {ext}, 文件: {image_path}")
            return False
        
        # 检查文件大小
        if not is_file_size_valid(image_path, self.max_file_size_mb):
            logger.warning(f"文件大小超过限制 ({self.max_file_size_mb}MB): {image_path}")
            return False
        
        return True
    
    def process_image(self, image_path: str) -> Optional[str]:
        """
        处理单个图片
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            输出Excel文件路径，如果处理失败则返回None
        """
        # 验证图片
        if not self.validate_image(image_path):
            return None
        
        # 如果需要跳过已处理的文件
        if self.skip_existing and self.record_manager.is_processed(image_path):
            output_file = self.record_manager.get_output_file(image_path)
            logger.info(f"图片已处理，跳过: {image_path}, 输出文件: {output_file}")
            return output_file
        
        logger.info(f"开始处理图片: {image_path}")
        
        try:
            # 生成输出文件路径
            file_name = os.path.splitext(os.path.basename(image_path))[0]
            output_file = os.path.join(self.output_folder, f"{file_name}{self.excel_extension}")
            
            # 检查是否已存在对应的Excel文件
            if os.path.exists(output_file) and self.skip_existing:
                logger.info(f"已存在对应的Excel文件，跳过处理: {os.path.basename(image_path)} -> {os.path.basename(output_file)}")
                # 记录处理结果
                self.record_manager.mark_as_processed(image_path, output_file)
                return output_file
            
            # 进行OCR识别
            ocr_result = self.ocr_client.recognize_table(image_path)
            if not ocr_result:
                logger.error(f"OCR识别失败: {image_path}")
                return None
                
            # 保存Excel文件 - 按照v1版本逻辑提取Excel数据
            excel_base64 = None
            
            # 从不同可能的字段中尝试获取Excel数据
            if 'excel_file' in ocr_result:
                excel_base64 = ocr_result['excel_file']
                logger.debug("从excel_file字段获取Excel数据")
            elif 'result' in ocr_result:
                if 'result_data' in ocr_result['result']:
                    excel_base64 = ocr_result['result']['result_data']
                    logger.debug("从result.result_data字段获取Excel数据")
                elif 'excel_file' in ocr_result['result']:
                    excel_base64 = ocr_result['result']['excel_file']
                    logger.debug("从result.excel_file字段获取Excel数据")
                elif 'tables_result' in ocr_result['result'] and ocr_result['result']['tables_result']:
                    for table in ocr_result['result']['tables_result']:
                        if 'excel_file' in table:
                            excel_base64 = table['excel_file']
                            logger.debug("从tables_result中获取Excel数据")
                            break
                    
            # 如果还是没有找到Excel数据，尝试通过get_excel_result获取
            if not excel_base64:
                logger.info("无法从直接返回中获取Excel数据，尝试通过API获取...")
                excel_data = self.ocr_client.get_excel_result(ocr_result)
                if not excel_data:
                    logger.error(f"获取Excel结果失败: {image_path}")
                    return None
                    
                # 保存Excel文件
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
                with open(output_file, 'wb') as f:
                    f.write(excel_data)
            else:
                # 解码并保存Excel文件
                try:
                    excel_data = base64.b64decode(excel_base64)
                    os.makedirs(os.path.dirname(output_file), exist_ok=True)
                    with open(output_file, 'wb') as f:
                        f.write(excel_data)
                except Exception as e:
                    logger.error(f"解码或保存Excel数据时出错: {e}")
                    return None
            
            logger.info(f"图片处理成功: {image_path}, 输出文件: {output_file}")
            
            # 标记为已处理
            self.record_manager.mark_as_processed(image_path, output_file)
            
            return output_file
            
        except Exception as e:
            logger.error(f"处理图片时出错: {image_path}, 错误: {e}")
            return None
    
    def process_images_batch(self, batch_size: int = None, max_workers: int = None) -> Tuple[int, int]:
        """
        批量处理图片
        
        Args:
            batch_size: 批处理大小，如果为None则使用配置值
            max_workers: 最大线程数，如果为None则使用配置值
            
        Returns:
            (总处理数, 成功处理数)元组
        """
        # 使用配置值或参数值
        batch_size = batch_size or self.batch_size
        max_workers = max_workers or self.max_workers
        
        # 获取未处理的图片
        unprocessed_images = self.get_unprocessed_images()
        if not unprocessed_images:
            logger.warning("没有需要处理的图片")
            return 0, 0
        
        total = len(unprocessed_images)
        success = 0
        
        # 按批次处理
        for i in range(0, total, batch_size):
            batch = unprocessed_images[i:i + batch_size]
            logger.info(f"处理批次 {i//batch_size + 1}/{(total-1)//batch_size + 1}, 大小: {len(batch)}")
            
            # 使用线程池并行处理
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                results = list(executor.map(self.process_image, batch))
            
            # 统计成功数
            success += sum(1 for result in results if result is not None)
            
            logger.info(f"批次处理完成, 成功: {sum(1 for result in results if result is not None)}/{len(batch)}")
        
        logger.info(f"所有图片处理完成, 总计: {total}, 成功: {success}")
        return total, success 