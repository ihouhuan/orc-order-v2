#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
百度表格OCR识别工具
-----------------
用于将图片中的表格转换为Excel文件的工具。
使用百度云OCR API进行识别。
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
from typing import Dict, List, Optional, Any

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
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
        'retry_delay': '2'
    },
    'Paths': {
        'input_folder': 'input',
        'output_folder': 'output',
        'temp_folder': 'temp'
    },
    'File': {
        'allowed_extensions': '.jpg,.jpeg,.png,.bmp',
        'excel_extension': '.xlsx'
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
    
    def getboolean(self, section: str, option: str, fallback: bool = False) -> bool:
        """获取布尔配置值"""
        return self.config.getboolean(section, option, fallback=fallback)
    
    def get_list(self, section: str, option: str, fallback: str = "", delimiter: str = ",") -> List[str]:
        """获取列表配置值"""
        value = self.get(section, option, fallback)
        return [item.strip() for item in value.split(delimiter) if item.strip()]

class OCRProcessor:
    """OCR处理器，用于表格识别"""
    
    def __init__(self, config_file: str = 'config.ini'):
        """
        初始化OCR处理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_manager = ConfigManager(config_file)
        
        # 获取配置
        self.api_key = self.config_manager.get('API', 'api_key')
        self.secret_key = self.config_manager.get('API', 'secret_key')
        self.timeout = self.config_manager.getint('API', 'timeout', 30)
        self.max_retries = self.config_manager.getint('API', 'max_retries', 3)
        self.retry_delay = self.config_manager.getint('API', 'retry_delay', 2)
        
        # 设置路径
        self.input_folder = self.config_manager.get('Paths', 'input_folder', 'input')
        self.output_folder = self.config_manager.get('Paths', 'output_folder', 'output')
        self.temp_folder = self.config_manager.get('Paths', 'temp_folder', 'temp')
        
        # 确保目录存在
        for dir_path in [self.input_folder, self.output_folder, self.temp_folder]:
            os.makedirs(dir_path, exist_ok=True)
        
        # 设置允许的文件扩展名
        self.allowed_extensions = self.config_manager.get_list('File', 'allowed_extensions')
        
        # 验证API配置
        if not self.api_key or not self.secret_key:
            logger.warning("API密钥未设置，请在配置文件中设置API密钥")
    
    def get_access_token(self) -> Optional[str]:
        """获取百度API访问令牌"""
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
                        return result["access_token"]
                
                logger.warning(f"获取访问令牌失败 (尝试 {attempt+1}/{self.max_retries}): {response.text}")
                
            except Exception as e:
                logger.warning(f"获取访问令牌时发生错误 (尝试 {attempt+1}/{self.max_retries}): {e}")
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < self.max_retries - 1:
                time.sleep(self.retry_delay * (attempt + 1))
        
        logger.error("无法获取访问令牌")
        return None
    
    def rename_image_to_timestamp(self, image_path: str) -> str:
        """将图片重命名为时间戳格式（如果需要）"""
        try:
            # 获取当前时间戳
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y%m%d%H%M%S")
            
            # 构造新文件名
            dir_path = os.path.dirname(image_path)
            ext = os.path.splitext(image_path)[1]
            new_path = os.path.join(dir_path, f"{timestamp}{ext}")
            
            # 如果文件名不同，则重命名
            if image_path != new_path:
                os.rename(image_path, new_path)
                logger.info(f"已将图片重命名为: {os.path.basename(new_path)}")
                return new_path
            
            return image_path
        except Exception as e:
            logger.error(f"重命名图片时出错: {e}")
            return image_path
    
    def recognize_table(self, image_path: str) -> Optional[Dict]:
        """
        识别图片中的表格
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            Dict: 识别结果，失败返回None
        """
        try:
            # 获取access_token
            access_token = self.get_access_token()
            if not access_token:
                return None
            
            # 请求URL
            url = f"https://aip.baidubce.com/rest/2.0/solution/v1/form_ocr/request?access_token={access_token}"
            
            # 读取图片内容
            with open(image_path, 'rb') as f:
                image_data = f.read()
            
            # Base64编码
            image_base64 = base64.b64encode(image_data).decode('utf-8')
            
            # 请求参数
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            data = {
                'image': image_base64,
                'is_sync': 'true',
                'request_type': 'excel'
            }
            
            # 发送请求
            response = requests.post(url, headers=headers, data=data, timeout=self.timeout)
            response.raise_for_status()
            
            # 解析结果
            result = response.json()
            
            # 检查错误码
            if 'error_code' in result:
                logger.error(f"识别表格失败: {result.get('error_msg', '未知错误')}")
                return None
            
            # 返回识别结果
            return result
            
        except Exception as e:
            logger.error(f"识别表格时出错: {e}")
            return None
    
    def get_excel_result(self, request_id: str, access_token: str) -> Optional[bytes]:
        """
        获取Excel结果
        
        Args:
            request_id: 请求ID
            access_token: 访问令牌
            
        Returns:
            bytes: Excel文件内容，失败返回None
        """
        try:
            # 请求URL
            url = f"https://aip.baidubce.com/rest/2.0/solution/v1/form_ocr/get_request_result?access_token={access_token}"
            
            # 请求参数
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            data = {
                'request_id': request_id,
                'result_type': 'excel'
            }
            
            # 最大重试次数
            max_retries = 10
            
            # 循环获取结果
            for i in range(max_retries):
                # 发送请求
                response = requests.post(url, headers=headers, data=data, timeout=self.timeout)
                response.raise_for_status()
                
                # 解析结果
                result = response.json()
                
                # 检查错误码
                if 'error_code' in result:
                    logger.error(f"获取Excel结果失败: {result.get('error_msg', '未知错误')}")
                    return None
                
                # 检查处理状态
                result_data = result.get('result', {})
                status = result_data.get('ret_code')
                
                if status == 3:  # 处理完成
                    # 获取Excel文件URL
                    excel_url = result_data.get('result_data')
                    if not excel_url:
                        logger.error("未获取到Excel结果URL")
                        return None
                    
                    # 下载Excel文件
                    excel_response = requests.get(excel_url)
                    excel_response.raise_for_status()
                    
                    # 返回Excel文件内容
                    return excel_response.content
                    
                elif status == 1:  # 排队中
                    logger.info(f"请求排队中 ({i+1}/{max_retries})，等待后重试...")
                elif status == 2:  # 处理中
                    logger.info(f"正在处理 ({i+1}/{max_retries})，等待后重试...")
                else:
                    logger.error(f"未知状态码: {status}")
                    return None
                
                # 等待后重试
                time.sleep(2)
            
            logger.error(f"获取Excel结果超时，请稍后再试")
            return None
            
        except Exception as e:
            logger.error(f"获取Excel结果时出错: {e}")
            return None
    
    def process_image(self, image_path: str) -> Optional[str]:
        """
        处理单个图片
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            str: 生成的Excel文件路径，失败返回None
        """
        try:
            logger.info(f"开始处理图片: {image_path}")
            
            # 验证文件扩展名
            ext = os.path.splitext(image_path)[1].lower()
            if self.allowed_extensions and ext not in self.allowed_extensions:
                logger.error(f"不支持的文件类型: {ext}，支持的类型: {', '.join(self.allowed_extensions)}")
                return None
            
            # 重命名图片（可选）
            renamed_path = self.rename_image_to_timestamp(image_path)
            
            # 获取文件名（不含扩展名）
            basename = os.path.basename(renamed_path)
            name_without_ext = os.path.splitext(basename)[0]
            
            # 获取access_token
            access_token = self.get_access_token()
            if not access_token:
                return None
            
            # 识别表格
            ocr_result = self.recognize_table(renamed_path)
            if not ocr_result:
                return None
            
            # 获取请求ID
            request_id = ocr_result.get('result', {}).get('request_id')
            if not request_id:
                logger.error("未获取到请求ID")
                return None
            
            # 获取Excel结果
            excel_content = self.get_excel_result(request_id, access_token)
            if not excel_content:
                return None
            
            # 保存Excel文件
            output_path = os.path.join(self.output_folder, f"{name_without_ext}.xlsx")
            with open(output_path, 'wb') as f:
                f.write(excel_content)
            
            logger.info(f"已保存Excel文件: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"处理图片时出错: {e}")
            return None
    
    def process_directory(self) -> List[str]:
        """
        处理输入目录中的所有图片
        
        Returns:
            List[str]: 生成的Excel文件路径列表
        """
        results = []
        
        try:
            # 获取输入目录中的所有图片文件
            image_files = []
            for ext in self.allowed_extensions:
                image_files.extend(list(Path(self.input_folder).glob(f"*{ext}")))
                image_files.extend(list(Path(self.input_folder).glob(f"*{ext.upper()}")))
            
            if not image_files:
                logger.warning(f"输入目录 {self.input_folder} 中没有找到图片文件")
                return []
            
            logger.info(f"在 {self.input_folder} 中找到 {len(image_files)} 个图片文件")
            
            # 处理每个图片
            for image_file in image_files:
                result = self.process_image(str(image_file))
                if result:
                    results.append(result)
            
            logger.info(f"处理完成，成功生成 {len(results)} 个Excel文件")
            return results
            
        except Exception as e:
            logger.error(f"处理目录时出错: {e}")
            return results

def main():
    """主函数"""
    import argparse
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='百度表格OCR识别工具')
    parser.add_argument('--config', type=str, default='config.ini', help='配置文件路径')
    parser.add_argument('--input', type=str, help='输入图片路径')
    parser.add_argument('--debug', action='store_true', help='启用调试模式')
    args = parser.parse_args()
    
    # 设置日志级别
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # 创建OCR处理器
    processor = OCRProcessor(args.config)
    
    # 处理单个图片或目录
    if args.input:
        if os.path.isfile(args.input):
            result = processor.process_image(args.input)
            if result:
                print(f"处理成功: {result}")
                return 0
            else:
                print("处理失败")
                return 1
        elif os.path.isdir(args.input):
            results = processor.process_directory()
            print(f"处理完成，成功生成 {len(results)} 个Excel文件")
            return 0
        else:
            print(f"输入路径不存在: {args.input}")
            return 1
    else:
        # 处理默认输入目录
        results = processor.process_directory()
        print(f"处理完成，成功生成 {len(results)} 个Excel文件")
        return 0

if __name__ == "__main__":
    sys.exit(main()) 