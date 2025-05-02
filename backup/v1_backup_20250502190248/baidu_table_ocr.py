#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
百度表格OCR识别工具
-----------------
用于将图片中的表格转换为Excel文件的工具。
使用百度云OCR API进行识别，支持批量处理。
"""

import os
import sys
import requests
import base64
import json
import time
import logging
import datetime
import configparser
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Any
from concurrent.futures import ThreadPoolExecutor

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ocr_processor.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# 默认配置
DEFAULT_CONFIG = {
    'API': {
        'api_key': '',  # 将从配置文件中读取
        'secret_key': '',  # 将从配置文件中读取
        'timeout': '30',
        'max_retries': '3',
        'retry_delay': '2',
        'api_url': 'https://aip.baidubce.com/rest/2.0/ocr/v1/table'
    },
    'Paths': {
        'input_folder': 'input',
        'output_folder': 'output',
        'temp_folder': 'temp',
        'processed_record': 'processed_files.json'
    },
    'Performance': {
        'max_workers': '4',
        'batch_size': '5',
        'skip_existing': 'true'
    },
    'File': {
        'allowed_extensions': '.jpg,.jpeg,.png,.bmp',
        'excel_extension': '.xlsx',
        'max_file_size_mb': '4'
    }
}

class ConfigManager:
    """配置管理类，负责加载和保存配置"""
    
    def __init__(self, config_file: str = 'config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.load_config()
    
    def load_config(self) -> None:
        """加载配置文件，如果不存在则创建默认配置"""
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
        """获取列表配置值"""
        value = self.get(section, option, fallback)
        return [item.strip() for item in value.split(delimiter) if item.strip()]

class TokenManager:
    """令牌管理类，负责获取和刷新百度API访问令牌"""
    
    def __init__(self, api_key: str, secret_key: str, max_retries: int = 3, retry_delay: int = 2):
        self.api_key = api_key
        self.secret_key = secret_key
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.access_token = None
        self.token_expiry = 0
    
    def get_token(self) -> Optional[str]:
        """获取访问令牌，如果令牌已过期则刷新"""
        if self.is_token_valid():
            return self.access_token
        
        return self.refresh_token()
    
    def is_token_valid(self) -> bool:
        """检查令牌是否有效"""
        return (
            self.access_token is not None and 
            self.token_expiry > time.time() + 60  # 提前1分钟刷新
        )
    
    def refresh_token(self) -> Optional[str]:
        """刷新访问令牌"""
        url = "https://aip.baidubce.com/oauth/2.0/token"
        params = {
            "grant_type": "client_credentials",
            "client_id": self.api_key,
            "client_secret": self.secret_key
        }
        
        for attempt in range(self.max_retries):
            try:
                response = requests.post(url, params=params, timeout=10)
                if response.status_code == 200:
                    result = response.json()
                    if "access_token" in result:
                        self.access_token = result["access_token"]
                        # 设置令牌过期时间（默认30天，提前1小时过期以确保安全）
                        self.token_expiry = time.time() + result.get("expires_in", 2592000) - 3600
                        logger.info("成功获取访问令牌")
                        return self.access_token
                
                logger.warning(f"获取访问令牌失败 (尝试 {attempt+1}/{self.max_retries}): {response.text}")
                
            except Exception as e:
                logger.warning(f"获取访问令牌时发生错误 (尝试 {attempt+1}/{self.max_retries}): {e}")
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < self.max_retries - 1:
                time.sleep(self.retry_delay * (attempt + 1))  # 指数退避
        
        logger.error("无法获取访问令牌")
        return None

class ProcessedRecordManager:
    """处理记录管理器，用于跟踪已处理的文件"""
    
    def __init__(self, record_file: str):
        self.record_file = record_file
        self.processed_files = self._load_record()
    
    def _load_record(self) -> Dict[str, str]:
        """加载处理记录"""
        if os.path.exists(self.record_file):
            try:
                with open(self.record_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"加载处理记录时出错: {e}")
        
        return {}
    
    def save_record(self) -> None:
        """保存处理记录"""
        try:
            with open(self.record_file, 'w', encoding='utf-8') as f:
                json.dump(self.processed_files, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存处理记录时出错: {e}")
    
    def is_processed(self, image_file: str) -> bool:
        """检查文件是否已处理"""
        return image_file in self.processed_files
    
    def mark_as_processed(self, image_file: str, output_file: str) -> None:
        """标记文件为已处理"""
        self.processed_files[image_file] = output_file
        self.save_record()
    
    def get_output_file(self, image_file: str) -> Optional[str]:
        """获取已处理文件对应的输出文件"""
        return self.processed_files.get(image_file)

class OCRProcessor:
    """OCR处理器核心类，用于识别表格并保存为Excel"""
    
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        
        # 路径配置
        self.input_folder = self.config.get('Paths', 'input_folder')
        self.output_folder = self.config.get('Paths', 'output_folder')
        self.temp_folder = self.config.get('Paths', 'temp_folder')
        self.processed_record_file = os.path.join(
            self.config.get('Paths', 'output_folder'), 
            self.config.get('Paths', 'processed_record')
        )
        
        # API配置
        self.api_url = self.config.get('API', 'api_url')
        self.timeout = self.config.getint('API', 'timeout')
        self.max_retries = self.config.getint('API', 'max_retries')
        self.retry_delay = self.config.getint('API', 'retry_delay')
        
        # 文件配置
        self.allowed_extensions = self.config.get_list('File', 'allowed_extensions')
        self.excel_extension = self.config.get('File', 'excel_extension')
        self.max_file_size_mb = self.config.getfloat('File', 'max_file_size_mb')
        
        # 性能配置
        self.max_workers = self.config.getint('Performance', 'max_workers')
        self.batch_size = self.config.getint('Performance', 'batch_size')
        self.skip_existing = self.config.getboolean('Performance', 'skip_existing')
        
        # 初始化其他组件
        self.token_manager = TokenManager(
            self.config.get('API', 'api_key'),
            self.config.get('API', 'secret_key'),
            self.max_retries,
            self.retry_delay
        )
        self.record_manager = ProcessedRecordManager(self.processed_record_file)
        
        # 确保文件夹存在
        for folder in [self.input_folder, self.output_folder, self.temp_folder]:
            os.makedirs(folder, exist_ok=True)
            logger.info(f"已确保文件夹存在: {folder}")
    
    def get_unprocessed_images(self) -> List[str]:
        """获取待处理的图像文件列表"""
        all_files = []
        for ext in self.allowed_extensions:
            all_files.extend(Path(self.input_folder).glob(f"*{ext}"))
        
        # 转换为字符串路径
        file_paths = [str(file_path) for file_path in all_files]
        
        if self.skip_existing:
            # 过滤掉已处理的文件
            return [
                file_path for file_path in file_paths
                if not self.record_manager.is_processed(os.path.basename(file_path))
            ]
        
        return file_paths
    
    def validate_image(self, image_path: str) -> bool:
        """验证图像文件是否有效且符合大小限制"""
        # 检查文件是否存在
        if not os.path.exists(image_path):
            logger.error(f"文件不存在: {image_path}")
            return False
        
        # 检查是否是文件
        if not os.path.isfile(image_path):
            logger.error(f"路径不是文件: {image_path}")
            return False
        
        # 检查文件大小
        file_size_mb = os.path.getsize(image_path) / (1024 * 1024)
        if file_size_mb > self.max_file_size_mb:
            logger.error(f"文件过大 ({file_size_mb:.2f}MB > {self.max_file_size_mb}MB): {image_path}")
            return False
        
        # 检查文件扩展名
        _, ext = os.path.splitext(image_path)
        if ext.lower() not in self.allowed_extensions:
            logger.error(f"不支持的文件格式 {ext}: {image_path}")
            return False
        
        return True
    
    def rename_image_to_timestamp(self, image_path: str) -> str:
        """将图像文件重命名为时间戳格式"""
        try:
            # 获取目录和文件扩展名
            dir_name = os.path.dirname(image_path)
            file_ext = os.path.splitext(image_path)[1]
            
            # 生成时间戳文件名
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            new_filename = f"{timestamp}{file_ext}"
            
            # 构建新路径
            new_path = os.path.join(dir_name, new_filename)
            
            # 如果目标文件已存在，添加毫秒级别的后缀
            if os.path.exists(new_path):
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
                new_filename = f"{timestamp}{file_ext}"
                new_path = os.path.join(dir_name, new_filename)
            
            # 重命名文件
            os.rename(image_path, new_path)
            logger.info(f"文件已重命名: {os.path.basename(image_path)} -> {new_filename}")
            return new_path
            
        except Exception as e:
            logger.error(f"重命名文件时出错: {e}")
            return image_path
    
    def recognize_table(self, image_path: str) -> Optional[Dict]:
        """使用百度表格OCR API识别图像中的表格"""
        # 获取访问令牌
        access_token = self.token_manager.get_token()
        if not access_token:
            logger.error("无法获取访问令牌")
            return None
        
        url = f"{self.api_url}?access_token={access_token}"
        
        for attempt in range(self.max_retries):
            try:
                # 读取图像文件并进行base64编码
                with open(image_path, 'rb') as f:
                    image_data = f.read()
                
                image_base64 = base64.b64encode(image_data).decode('utf-8')
                
                # 设置请求头和请求参数
                headers = {'Content-Type': 'application/x-www-form-urlencoded'}
                params = {
                    'image': image_base64,
                    'return_excel': 'true'  # 返回Excel文件编码
                }
                
                # 发送请求
                response = requests.post(
                    url, 
                    data=params, 
                    headers=headers, 
                    timeout=self.timeout
                )
                
                # 检查响应状态
                if response.status_code == 200:
                    result = response.json()
                    if 'error_code' in result:
                        error_msg = result.get('error_msg', '未知错误')
                        logger.error(f"表格识别失败: {error_msg}")
                        
                        # 如果是授权错误，尝试刷新令牌
                        if result.get('error_code') in [110, 111]:  # 授权相关错误码
                            self.token_manager.refresh_token()
                    else:
                        return result
                else:
                    logger.error(f"表格识别失败: {response.status_code} - {response.text}")
                
            except Exception as e:
                logger.error(f"表格识别过程中发生错误 (尝试 {attempt+1}/{self.max_retries}): {e}")
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < self.max_retries - 1:
                wait_time = self.retry_delay * (2 ** attempt)  # 指数退避
                logger.info(f"将在 {wait_time} 秒后重试...")
                time.sleep(wait_time)
        
        return None
    
    def save_to_excel(self, ocr_result: Dict, output_path: str) -> bool:
        """将表格识别结果保存为Excel文件"""
        try:
            # 检查结果中是否包含表格数据和Excel文件
            if not ocr_result:
                logger.error("无法保存结果: 识别结果为空")
                return False
            
            # 直接从excel_file字段获取Excel文件的base64编码
            excel_base64 = None
            
            if 'excel_file' in ocr_result:
                excel_base64 = ocr_result['excel_file']
            elif 'tables_result' in ocr_result and ocr_result['tables_result']:
                for table in ocr_result['tables_result']:
                    if 'excel_file' in table:
                        excel_base64 = table['excel_file']
                        break
            
            if not excel_base64:
                logger.error("无法获取Excel文件编码")
                logger.debug(f"API返回结果: {json.dumps(ocr_result, ensure_ascii=False, indent=2)}")
                return False
            
            # 解码base64并保存Excel文件
            try:
                excel_data = base64.b64decode(excel_base64)
                
                # 确保输出目录存在
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                with open(output_path, 'wb') as f:
                    f.write(excel_data)
                
                logger.info(f"成功保存表格数据到: {output_path}")
                return True
                
            except Exception as e:
                logger.error(f"解码Excel数据时出错: {e}")
                return False
            
        except Exception as e:
            logger.error(f"保存Excel文件时发生错误: {e}")
            return False
    
    def process_image(self, image_path: str) -> Optional[str]:
        """处理单个图像文件：验证、重命名、识别和保存"""
        try:
            # 获取原始图片文件名（不含扩展名）
            image_basename = os.path.basename(image_path)
            image_name_without_ext = os.path.splitext(image_basename)[0]
            
            # 检查是否已存在对应的Excel文件
            excel_filename = f"{image_name_without_ext}{self.excel_extension}"
            excel_path = os.path.join(self.output_folder, excel_filename)
            
            if os.path.exists(excel_path):
                logger.info(f"已存在对应的Excel文件，跳过处理: {image_basename} -> {excel_filename}")
                # 记录处理结果（虽然跳过了处理，但仍标记为已处理）
                self.record_manager.mark_as_processed(image_basename, excel_path)
                return excel_path
            
            # 检查文件是否已经处理过
            if self.skip_existing and self.record_manager.is_processed(image_basename):
                output_file = self.record_manager.get_output_file(image_basename)
                logger.info(f"文件已处理过，跳过: {image_basename} -> {output_file}")
                return output_file
            
            # 验证图像文件
            if not self.validate_image(image_path):
                logger.warning(f"图像验证失败: {image_path}")
                return None
            
            # 识别表格（不再重命名图片）
            logger.info(f"正在识别表格: {image_basename}")
            ocr_result = self.recognize_table(image_path)
            
            if not ocr_result:
                logger.error(f"表格识别失败: {image_basename}")
                return None
            
            # 保存结果到Excel，使用原始图片名
            if self.save_to_excel(ocr_result, excel_path):
                # 记录处理结果
                self.record_manager.mark_as_processed(image_basename, excel_path)
                return excel_path
            
            return None
            
        except Exception as e:
            logger.error(f"处理图像时发生错误: {e}")
            return None
    
    def process_images_batch(self, batch_size: int = None, max_workers: int = None) -> Tuple[int, int]:
        """批量处理图像文件"""
        if batch_size is None:
            batch_size = self.batch_size
            
        if max_workers is None:
            max_workers = self.max_workers
        
        # 获取待处理的图像文件
        image_files = self.get_unprocessed_images()
        total_files = len(image_files)
        
        if total_files == 0:
            logger.info("没有需要处理的图像文件")
            return 0, 0
        
        logger.info(f"找到 {total_files} 个待处理图像文件")
        
        # 处理所有文件
        processed_count = 0
        success_count = 0
        
        # 如果文件数量很少，直接顺序处理
        if total_files <= 2 or max_workers <= 1:
            for image_path in image_files:
                processed_count += 1
                
                logger.info(f"处理文件 ({processed_count}/{total_files}): {os.path.basename(image_path)}")
                output_path = self.process_image(image_path)
                
                if output_path:
                    success_count += 1
                    logger.info(f"处理成功 ({success_count}/{processed_count}): {os.path.basename(output_path)}")
                else:
                    logger.warning(f"处理失败: {os.path.basename(image_path)}")
        else:
            # 使用线程池并行处理
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                for i in range(0, total_files, batch_size):
                    batch = image_files[i:i+batch_size]
                    batch_results = list(executor.map(self.process_image, batch))
                    
                    for j, result in enumerate(batch_results):
                        processed_count += 1
                        if result:
                            success_count += 1
                            logger.info(f"处理成功 ({success_count}/{processed_count}): {os.path.basename(result)}")
                        else:
                            logger.warning(f"处理失败: {os.path.basename(batch[j])}")
                    
                    logger.info(f"已处理 {processed_count}/{total_files} 个文件，成功率: {success_count/processed_count*100:.1f}%")
        
        logger.info(f"处理完成。总共处理 {processed_count} 个文件，成功 {success_count} 个，成功率: {success_count/max(processed_count,1)*100:.1f}%")
        return processed_count, success_count
    
    def check_processed_status(self) -> Dict[str, List[str]]:
        """检查处理状态，返回已处理和未处理的文件列表"""
        # 获取输入文件夹中的所有支持格式的图像文件
        all_images = []
        for ext in self.allowed_extensions:
            all_images.extend([str(file) for file in Path(self.input_folder).glob(f"*{ext}")])
        
        # 获取已处理的文件列表
        processed_files = list(self.record_manager.processed_files.keys())
        
        # 对路径进行规范化以便比较
        all_image_basenames = [os.path.basename(img) for img in all_images]
        
        # 找出未处理的文件
        unprocessed_files = [
            img for img, basename in zip(all_images, all_image_basenames)
            if basename not in processed_files
        ]
        
        # 找出已处理的文件及其对应的输出文件
        processed_with_output = {
            img: self.record_manager.get_output_file(basename)
            for img, basename in zip(all_images, all_image_basenames)
            if basename in processed_files
        }
        
        return {
            'all': all_images,
            'unprocessed': unprocessed_files,
            'processed': processed_with_output
        }

def main():
    """主函数: 解析命令行参数并执行相应操作"""
    import argparse
    
    parser = argparse.ArgumentParser(description='百度表格OCR识别工具')
    parser.add_argument('--config', type=str, default='config.ini', help='配置文件路径')
    parser.add_argument('--batch-size', type=int, help='批处理大小')
    parser.add_argument('--max-workers', type=int, help='最大工作线程数')
    parser.add_argument('--force', action='store_true', help='强制处理所有文件，包括已处理的文件')
    parser.add_argument('--check', action='store_true', help='检查处理状态而不执行处理')
    
    args = parser.parse_args()
    
    # 加载配置
    config_manager = ConfigManager(args.config)
    
    # 创建处理器
    processor = OCRProcessor(config_manager)
    
    # 根据命令行参数调整配置
    if args.force:
        processor.skip_existing = False
    
    if args.check:
        # 检查处理状态
        status = processor.check_processed_status()
        
        print("\n=== 处理状态 ===")
        print(f"总共 {len(status['all'])} 个图像文件")
        print(f"已处理: {len(status['processed'])} 个")
        print(f"未处理: {len(status['unprocessed'])} 个")
        
        if status['processed']:
            print("\n已处理文件:")
            for img, output in status['processed'].items():
                print(f"  {os.path.basename(img)} -> {os.path.basename(output)}")
        
        if status['unprocessed']:
            print("\n未处理文件:")
            for img in status['unprocessed']:
                print(f"  {os.path.basename(img)}")
        
        return
    
    # 处理图像
    batch_size = args.batch_size if args.batch_size is not None else processor.batch_size
    max_workers = args.max_workers if args.max_workers is not None else processor.max_workers
    
    processor.process_images_batch(batch_size, max_workers)

if __name__ == "__main__":
    try:
        start_time = time.time()
        logger.info("开始百度表格OCR识别程序")
        main()
        elapsed_time = time.time() - start_time
        logger.info(f"百度表格OCR识别程序已完成，耗时: {elapsed_time:.2f}秒")
    except Exception as e:
        logger.error(f"程序执行过程中发生错误: {e}", exc_info=True)
        sys.exit(1) 