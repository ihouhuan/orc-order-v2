#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
测试OCR处理器日志文件创建
"""

import os
import sys
import logging
from datetime import datetime

# 确保logs目录存在
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
print(f"日志目录: {log_dir}")

# 设置日志文件路径
log_file = os.path.join(log_dir, 'ocr_processor.log')
print(f"日志文件路径: {log_file}")

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

# 写入测试日志
logger.info("这是一条测试日志消息")
logger.info(f"测试时间: {datetime.now()}")

# 标记该日志文件为活跃，避免被清理工具删除
try:
    # 创建一个标记文件，表示该日志文件正在使用中
    active_marker = os.path.join(log_dir, 'ocr_processor.active')
    with open(active_marker, 'w') as f:
        f.write(f"Active since: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"活跃标记文件: {active_marker}")
except Exception as e:
    print(f"无法创建日志活跃标记: {e}")

# 检查文件是否已创建
if os.path.exists(log_file):
    print(f"日志文件已成功创建: {log_file}")
    print(f"文件大小: {os.path.getsize(log_file)} 字节")
else:
    print(f"错误: 日志文件创建失败: {log_file}")

print("测试完成") 