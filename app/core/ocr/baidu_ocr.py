"""
百度OCR客户端模块
---------------
提供百度OCR API的访问和调用功能。
"""

import os
import time
import base64
import requests
import logging
from typing import Dict, Optional, Any, Union

from ...config.settings import ConfigManager
from ..utils.log_utils import get_logger

logger = get_logger(__name__)

class TokenManager:
    """
    令牌管理类，负责获取和刷新百度API访问令牌
    """
    
    def __init__(self, api_key: str, secret_key: str, max_retries: int = 3, retry_delay: int = 2):
        """
        初始化令牌管理器
        
        Args:
            api_key: 百度API Key
            secret_key: 百度Secret Key
            max_retries: 最大重试次数
            retry_delay: 重试延迟（秒）
        """
        self.api_key = api_key
        self.secret_key = secret_key
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.access_token = None
        self.token_expiry = 0
    
    def get_token(self) -> Optional[str]:
        """
        获取访问令牌，如果令牌已过期则刷新
        
        Returns:
            访问令牌，如果获取失败则返回None
        """
        if self.is_token_valid():
            return self.access_token
        
        return self.refresh_token()
    
    def is_token_valid(self) -> bool:
        """
        检查令牌是否有效
        
        Returns:
            令牌是否有效
        """
        return (
            self.access_token is not None and 
            self.token_expiry > time.time() + 60  # 提前1分钟刷新
        )
    
    def refresh_token(self) -> Optional[str]:
        """
        刷新访问令牌
        
        Returns:
            新的访问令牌，如果获取失败则返回None
        """
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

class BaiduOCRClient:
    """
    百度OCR API客户端
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化百度OCR客户端
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        self.config = config or ConfigManager()
        
        # 获取配置
        self.api_key = self.config.get('API', 'api_key')
        self.secret_key = self.config.get('API', 'secret_key')
        self.timeout = self.config.getint('API', 'timeout', 30)
        self.max_retries = self.config.getint('API', 'max_retries', 3)
        self.retry_delay = self.config.getint('API', 'retry_delay', 2)
        self.api_url = self.config.get('API', 'api_url', 'https://aip.baidubce.com/rest/2.0/ocr/v1/table')
        
        # 创建令牌管理器
        self.token_manager = TokenManager(
            self.api_key, 
            self.secret_key, 
            self.max_retries, 
            self.retry_delay
        )
        
        # 验证API配置
        if not self.api_key or not self.secret_key:
            logger.warning("API密钥未设置，请在配置文件中设置API密钥")
    
    def read_image(self, image_path: str) -> Optional[bytes]:
        """
        读取图片文件为二进制数据
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            图片二进制数据，如果读取失败则返回None
        """
        try:
            with open(image_path, 'rb') as f:
                return f.read()
        except Exception as e:
            logger.error(f"读取图片文件失败: {image_path}, 错误: {e}")
            return None
    
    def recognize_table(self, image_data: Union[str, bytes]) -> Optional[Dict]:
        """
        识别表格
        
        Args:
            image_data: 图片数据，可以是文件路径或二进制数据
            
        Returns:
            识别结果字典，如果识别失败则返回None
        """
        # 获取访问令牌
        access_token = self.token_manager.get_token()
        if not access_token:
            logger.error("无法获取访问令牌，无法进行表格识别")
            return None
        
        # 如果是文件路径，读取图片数据
        if isinstance(image_data, str):
            image_data = self.read_image(image_data)
            if image_data is None:
                return None
        
        # 准备请求参数
        url = f"{self.api_url}?access_token={access_token}"
        image_base64 = base64.b64encode(image_data).decode('utf-8')
        
        # 请求参数 - 添加return_excel参数，与v1版本保持一致
        payload = {
            'image': image_base64,
            'is_sync': 'true',  # 同步请求
            'request_type': 'excel',  # 输出为Excel
            'return_excel': 'true'  # 直接返回Excel数据
        }
        
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json'
        }
        
        # 发送请求
        for attempt in range(self.max_retries):
            try:
                response = requests.post(
                    url, 
                    data=payload, 
                    headers=headers, 
                    timeout=self.timeout
                )
                
                if response.status_code == 200:
                    result = response.json()
                    # 打印返回结果以便调试
                    logger.debug(f"百度OCR API返回结果: {result}")
                    
                    if 'error_code' in result:
                        error_msg = result.get('error_msg', '未知错误')
                        logger.error(f"百度OCR API错误: {error_msg}")
                        # 如果是授权错误，尝试刷新令牌
                        if result.get('error_code') in [110, 111]:  # 授权相关错误码
                            logger.info("尝试刷新访问令牌...")
                            self.token_manager.refresh_token()
                        return None
                    
                    # 兼容不同的返回结构
                    # 这是最关键的修改部分: 直接返回整个结果，不强制要求特定结构
                    return result
                else:
                    logger.warning(f"表格识别请求失败 (尝试 {attempt+1}/{self.max_retries}): {response.text}")
            
            except Exception as e:
                logger.warning(f"表格识别时发生错误 (尝试 {attempt+1}/{self.max_retries}): {e}")
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < self.max_retries - 1:
                wait_time = self.retry_delay * (2 ** attempt)  # 指数退避
                logger.info(f"将在 {wait_time} 秒后重试...")
                time.sleep(wait_time)
        
        logger.error("表格识别失败")
        return None
    
    def get_excel_result(self, request_id_or_result: Union[str, Dict]) -> Optional[bytes]:
        """
        获取Excel结果
        
        Args:
            request_id_or_result: 请求ID或完整的识别结果
            
        Returns:
            Excel二进制数据，如果获取失败则返回None
        """
        # 获取访问令牌
        access_token = self.token_manager.get_token()
        if not access_token:
            logger.error("无法获取访问令牌，无法获取Excel结果")
            return None
        
        # 处理直接传入结果对象的情况
        request_id = request_id_or_result
        if isinstance(request_id_or_result, dict):
            # v1版本兼容处理：如果结果中直接包含Excel数据
            if 'result' in request_id_or_result:
                # 如果是同步返回的Excel结果（某些API版本会直接返回）
                if 'result_data' in request_id_or_result['result']:
                    excel_content = request_id_or_result['result']['result_data']
                    if excel_content:
                        try:
                            return base64.b64decode(excel_content)
                        except Exception as e:
                            logger.error(f"解析Excel数据失败: {e}")
                
                # 提取request_id
                if 'request_id' in request_id_or_result['result']:
                    request_id = request_id_or_result['result']['request_id']
                    logger.debug(f"从result子对象中提取request_id: {request_id}")
                elif 'tables_result' in request_id_or_result['result'] and len(request_id_or_result['result']['tables_result']) > 0:
                    # 某些版本API可能直接返回表格内容，此时可能没有request_id
                    logger.info("检测到API直接返回了表格内容，但没有request_id")
                    return None
            # 有些版本可能request_id在顶层
            elif 'request_id' in request_id_or_result:
                request_id = request_id_or_result['request_id']
                logger.debug(f"从顶层对象中提取request_id: {request_id}")
        
        # 如果没有有效的request_id，无法获取结果
        if not isinstance(request_id, str):
            logger.error(f"无法从结果中提取有效的request_id: {request_id_or_result}")
            return None
            
        url = f"https://aip.baidubce.com/rest/2.0/solution/v1/form_ocr/get_request_result?access_token={access_token}"
        
        payload = {
            'request_id': request_id,
            'result_type': 'excel'
        }
        
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json'
        }
        
        for attempt in range(self.max_retries):
            try:
                response = requests.post(
                    url, 
                    data=payload, 
                    headers=headers, 
                    timeout=self.timeout
                )
                
                if response.status_code == 200:
                    try:
                        result = response.json()
                        logger.debug(f"获取Excel结果返回: {result}")
                        
                        # 检查是否还在处理中
                        if result.get('result', {}).get('ret_code') == 3:
                            logger.info(f"Excel结果正在处理中，等待后重试 (尝试 {attempt+1}/{self.max_retries})")
                            time.sleep(2)
                            continue
                        
                        # 检查是否有错误
                        if 'error_code' in result or result.get('result', {}).get('ret_code') != 0:
                            error_msg = result.get('error_msg') or result.get('result', {}).get('ret_msg', '未知错误')
                            logger.error(f"获取Excel结果失败: {error_msg}")
                            return None
                        
                        # 获取Excel内容
                        excel_content = result.get('result', {}).get('result_data')
                        if excel_content:
                            return base64.b64decode(excel_content)
                        else:
                            logger.error("Excel结果为空")
                            return None
                    
                    except Exception as e:
                        logger.error(f"解析Excel结果时出错: {e}")
                        return None
                
                else:
                    logger.warning(f"获取Excel结果请求失败 (尝试 {attempt+1}/{self.max_retries}): {response.text}")
            
            except Exception as e:
                logger.warning(f"获取Excel结果时发生错误 (尝试 {attempt+1}/{self.max_retries}): {e}")
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < self.max_retries - 1:
                time.sleep(self.retry_delay * (attempt + 1))
        
        logger.error("获取Excel结果失败")
        return None 