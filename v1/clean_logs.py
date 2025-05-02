#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
日志清理脚本
-----------
用于清理和管理日志文件，包括：
1. 清理指定天数之前的日志文件
2. 保留最新的N个日志文件
3. 清理过大的日志文件
4. 支持压缩旧日志文件
"""

import os
import sys
import time
import shutil
import logging
import argparse
from datetime import datetime, timedelta
import gzip
from pathlib import Path
import glob
import re

# 配置日志
logger = logging.getLogger(__name__)
if not logger.handlers:
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs', 'clean_logs.log')
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logger = logging.getLogger(__name__)
    
    # 标记该日志文件为活跃
    active_marker = os.path.join(os.path.dirname(log_file), 'clean_logs.active')
    with open(active_marker, 'w') as f:
        f.write(f"Active since: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

def is_log_active(log_file):
    """检查日志文件是否处于活跃状态（正在被使用）"""
    # 检查对应的活跃标记文件是否存在
    log_name = os.path.basename(log_file)
    base_name = os.path.splitext(log_name)[0]
    active_marker = os.path.join(os.path.dirname(log_file), f"{base_name}.active")
    
    # 如果活跃标记文件存在，说明日志文件正在被使用
    if os.path.exists(active_marker):
        logger.info(f"日志文件 {log_name} 正在使用中，不会被删除")
        return True
    
    # 检查是否是当前脚本正在使用的日志文件
    if log_name == os.path.basename(log_file):
        logger.info(f"当前脚本正在使用 {log_name}，不会被删除")
        return True
    
    return False

def clean_logs(log_dir="logs", max_days=7, max_files=10, max_size=100, force=False):
    """
    清理日志文件
    
    参数:
        log_dir: 日志目录
        max_days: 保留的最大天数
        max_files: 保留的最大文件数
        max_size: 日志文件大小上限(MB)
        force: 是否强制清理
    """
    logger.info(f"开始清理日志目录: {log_dir}")
    
    # 确保日志目录存在
    if not os.path.exists(log_dir):
        logger.warning(f"日志目录不存在: {log_dir}")
        return
        
    # 获取所有日志文件
    log_files = []
    for ext in ['*.log', '*.log.*']:
        log_files.extend(glob.glob(os.path.join(log_dir, ext)))
    
    if not log_files:
        logger.info(f"没有找到日志文件")
        return
        
    logger.info(f"找到 {len(log_files)} 个日志文件")
    
    # 按修改时间排序
    log_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    # 处理大文件
    for log_file in log_files:
        # 跳过活跃的日志文件
        if is_log_active(log_file):
            continue
            
        # 检查文件大小
        file_size_mb = os.path.getsize(log_file) / (1024 * 1024)
        if file_size_mb > max_size:
            logger.info(f"日志文件 {os.path.basename(log_file)} 大小为 {file_size_mb:.2f}MB，超过限制 {max_size}MB")
            
            # 压缩并重命名大文件
            compressed_file = f"{log_file}.{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            try:
                shutil.make_archive(os.path.splitext(compressed_file)[0], 'zip', log_dir, os.path.basename(log_file))
                logger.info(f"已压缩日志文件: {compressed_file}")
                
                # 清空原文件内容
                if not force:
                    confirm = input(f"是否清空日志文件 {os.path.basename(log_file)}? (y/n): ")
                    if confirm.lower() != 'y':
                        logger.info("已取消清空操作")
                        continue
                        
                with open(log_file, 'w') as f:
                    f.write(f"日志已于 {datetime.now()} 清空并压缩\n")
                logger.info(f"已清空日志文件: {os.path.basename(log_file)}")
            except Exception as e:
                logger.error(f"压缩日志文件时出错: {e}")
    
    # 清理过期的文件
    cutoff_date = datetime.now() - timedelta(days=max_days)
    files_to_delete = []
    
    for log_file in log_files[max_files:]:
        # 跳过活跃的日志文件
        if is_log_active(log_file):
            continue
            
        mtime = datetime.fromtimestamp(os.path.getmtime(log_file))
        if mtime < cutoff_date:
            files_to_delete.append(log_file)
    
    if not files_to_delete:
        logger.info("没有需要删除的过期日志文件")
        return
        
    logger.info(f"找到 {len(files_to_delete)} 个过期日志文件")
    
    # 确认删除
    if not force:
        print(f"以下 {len(files_to_delete)} 个文件将被删除:")
        for file in files_to_delete:
            print(f"  - {os.path.basename(file)}")
        confirm = input("确认删除? (y/n): ")
        if confirm.lower() != 'y':
            logger.info("已取消删除操作")
            return
    
    # 删除文件
    deleted_count = 0
    for file in files_to_delete:
        try:
            os.remove(file)
            logger.info(f"已删除日志文件: {os.path.basename(file)}")
            deleted_count += 1
        except Exception as e:
            logger.error(f"删除日志文件时出错: {e}")
    
    logger.info(f"成功删除 {deleted_count} 个日志文件")

def show_stats(log_dir="logs"):
    """显示日志文件统计信息"""
    if not os.path.exists(log_dir):
        print(f"日志目录不存在: {log_dir}")
        return
        
    log_files = []
    for ext in ['*.log', '*.log.*']:
        log_files.extend(glob.glob(os.path.join(log_dir, ext)))
        
    if not log_files:
        print("没有找到日志文件")
        return
        
    print(f"\n找到 {len(log_files)} 个日志文件:")
    print("=" * 80)
    print(f"{'文件名':<30} {'大小':<10} {'最后修改时间':<20} {'状态':<10}")
    print("-" * 80)
    
    total_size = 0
    for file in sorted(log_files, key=lambda x: os.path.getmtime(x), reverse=True):
        size = os.path.getsize(file)
        total_size += size
        
        mtime = datetime.fromtimestamp(os.path.getmtime(file))
        size_str = f"{size / 1024:.1f} KB" if size < 1024*1024 else f"{size / (1024*1024):.1f} MB"
        
        # 检查是否是活跃日志
        status = "活跃" if is_log_active(file) else ""
        
        print(f"{os.path.basename(file):<30} {size_str:<10} {mtime.strftime('%Y-%m-%d %H:%M:%S'):<20} {status:<10}")
    
    print("-" * 80)
    total_size_str = f"{total_size / 1024:.1f} KB" if total_size < 1024*1024 else f"{total_size / (1024*1024):.1f} MB"
    print(f"总大小: {total_size_str}")
    print("=" * 80)

def main():
    parser = argparse.ArgumentParser(description="日志文件清理工具")
    parser.add_argument("--max-days", type=int, default=7, help="日志保留的最大天数")
    parser.add_argument("--max-files", type=int, default=10, help="保留的最大文件数")
    parser.add_argument("--max-size", type=float, default=100, help="日志文件大小上限(MB)")
    parser.add_argument("--force", action="store_true", help="强制清理，不提示确认")
    parser.add_argument("--stats", action="store_true", help="显示日志统计信息")
    parser.add_argument("--log-dir", type=str, default="logs", help="日志目录")
    
    args = parser.parse_args()
    
    if args.stats:
        show_stats(args.log_dir)
    else:
        clean_logs(args.log_dir, args.max_days, args.max_files, args.max_size, args.force)

if __name__ == "__main__":
    main() 