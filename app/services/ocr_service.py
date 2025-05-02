"""
OCR服务模块
---------
提供OCR识别服务，协调OCR流程。
"""

from typing import Dict, List, Optional, Tuple, Union, Any

from ..config.settings import ConfigManager
from ..core.utils.log_utils import get_logger
from ..core.ocr.table_ocr import OCRProcessor

logger = get_logger(__name__)

class OCRService:
    """
    OCR识别服务：协调OCR流程
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化OCR服务
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        logger.info("初始化OCRService")
        self.config = config or ConfigManager()
        
        # 创建OCR处理器
        self.ocr_processor = OCRProcessor(self.config)
        
        logger.info("OCRService初始化完成")
    
    def get_unprocessed_images(self) -> List[str]:
        """
        获取待处理的图片列表
        
        Returns:
            待处理图片路径列表
        """
        return self.ocr_processor.get_unprocessed_images()
    
    def process_image(self, image_path: str) -> Optional[str]:
        """
        处理单张图片
        
        Args:
            image_path: 图片路径
            
        Returns:
            输出Excel文件路径，如果处理失败则返回None
        """
        logger.info(f"OCRService开始处理图片: {image_path}")
        result = self.ocr_processor.process_image(image_path)
        
        if result:
            logger.info(f"OCRService处理图片成功: {image_path} -> {result}")
        else:
            logger.error(f"OCRService处理图片失败: {image_path}")
        
        return result
    
    def process_images_batch(self, batch_size: int = None, max_workers: int = None) -> Tuple[int, int]:
        """
        批量处理图片
        
        Args:
            batch_size: 批处理大小
            max_workers: 最大线程数
            
        Returns:
            (总处理数, 成功处理数)元组
        """
        logger.info(f"OCRService开始批量处理图片, batch_size={batch_size}, max_workers={max_workers}")
        return self.ocr_processor.process_images_batch(batch_size, max_workers)
    
    def validate_image(self, image_path: str) -> bool:
        """
        验证图片是否有效
        
        Args:
            image_path: 图片路径
            
        Returns:
            图片是否有效
        """
        return self.ocr_processor.validate_image(image_path) 