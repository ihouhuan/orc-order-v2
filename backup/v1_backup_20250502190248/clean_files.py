#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件清理工具
-----------
用于清理输入/输出目录中的旧文件，支持按天数和文件名模式进行清理。
默认情况下会清理input目录下的所有图片文件和output目录下的Excel文件。
"""

import os
import re
import sys
import logging
import argparse
from datetime import datetime, timedelta
from pathlib import Path
import time
import glob

# 配置日志
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, 'clean_files.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class FileCleaner:
    """文件清理工具类"""
    
    def __init__(self, input_dir="input", output_dir="output"):
        """初始化清理工具"""
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
        
        # 确保目录存在
        for directory in [self.input_dir, self.output_dir, self.logs_dir]:
            os.makedirs(directory, exist_ok=True)
            logger.info(f"确保目录存在: {directory}")
    
    def get_file_stats(self, directory):
        """获取目录的文件统计信息"""
        if not os.path.exists(directory):
            logger.warning(f"目录不存在: {directory}")
            return {}
        
        stats = {
            'total_files': 0,
            'total_size': 0,
            'oldest_file': None,
            'newest_file': None,
            'file_types': {},
            'files_by_age': {
                '1_day': 0,
                '7_days': 0,
                '30_days': 0,
                'older': 0
            }
        }
        
        now = datetime.now()
        one_day_ago = now - timedelta(days=1)
        seven_days_ago = now - timedelta(days=7)
        thirty_days_ago = now - timedelta(days=30)
        
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                
                # 跳过临时文件
                if file.startswith('~$') or file.startswith('.'):
                    continue
                
                # 文件信息
                try:
                    file_stats = os.stat(file_path)
                    file_size = file_stats.st_size
                    mod_time = datetime.fromtimestamp(file_stats.st_mtime)
                    
                    # 更新统计信息
                    stats['total_files'] += 1
                    stats['total_size'] += file_size
                    
                    # 更新最旧和最新文件
                    if stats['oldest_file'] is None or mod_time < stats['oldest_file'][1]:
                        stats['oldest_file'] = (file_path, mod_time)
                    
                    if stats['newest_file'] is None or mod_time > stats['newest_file'][1]:
                        stats['newest_file'] = (file_path, mod_time)
                    
                    # 按文件类型统计
                    ext = os.path.splitext(file)[1].lower()
                    if ext in stats['file_types']:
                        stats['file_types'][ext]['count'] += 1
                        stats['file_types'][ext]['size'] += file_size
                    else:
                        stats['file_types'][ext] = {'count': 1, 'size': file_size}
                    
                    # 按年龄统计
                    if mod_time > one_day_ago:
                        stats['files_by_age']['1_day'] += 1
                    elif mod_time > seven_days_ago:
                        stats['files_by_age']['7_days'] += 1
                    elif mod_time > thirty_days_ago:
                        stats['files_by_age']['30_days'] += 1
                    else:
                        stats['files_by_age']['older'] += 1
                        
                except Exception as e:
                    logger.error(f"处理文件时出错 {file_path}: {e}")
        
        return stats
    
    def print_stats(self):
        """打印文件统计信息"""
        # 输入目录统计
        input_stats = self.get_file_stats(self.input_dir)
        output_stats = self.get_file_stats(self.output_dir)
        
        print("\n===== 文件统计信息 =====")
        
        # 打印输入目录统计
        if input_stats:
            print(f"\n输入目录 ({self.input_dir}):")
            print(f"  总文件数: {input_stats['total_files']}")
            print(f"  总大小: {self._format_size(input_stats['total_size'])}")
            
            if input_stats['oldest_file']:
                oldest = input_stats['oldest_file']
                print(f"  最旧文件: {os.path.basename(oldest[0])} ({oldest[1].strftime('%Y-%m-%d %H:%M:%S')})")
            
            if input_stats['newest_file']:
                newest = input_stats['newest_file']
                print(f"  最新文件: {os.path.basename(newest[0])} ({newest[1].strftime('%Y-%m-%d %H:%M:%S')})")
            
            print("  文件年龄分布:")
            print(f"    1天内: {input_stats['files_by_age']['1_day']}个文件")
            print(f"    7天内(不含1天内): {input_stats['files_by_age']['7_days']}个文件")
            print(f"    30天内(不含7天内): {input_stats['files_by_age']['30_days']}个文件")
            print(f"    更旧: {input_stats['files_by_age']['older']}个文件")
            
            print("  文件类型分布:")
            for ext, data in sorted(input_stats['file_types'].items(), key=lambda x: x[1]['count'], reverse=True):
                print(f"    {ext or '无扩展名'}: {data['count']}个文件, {self._format_size(data['size'])}")
        
        # 打印输出目录统计
        if output_stats:
            print(f"\n输出目录 ({self.output_dir}):")
            print(f"  总文件数: {output_stats['total_files']}")
            print(f"  总大小: {self._format_size(output_stats['total_size'])}")
            
            if output_stats['oldest_file']:
                oldest = output_stats['oldest_file']
                print(f"  最旧文件: {os.path.basename(oldest[0])} ({oldest[1].strftime('%Y-%m-%d %H:%M:%S')})")
            
            if output_stats['newest_file']:
                newest = output_stats['newest_file']
                print(f"  最新文件: {os.path.basename(newest[0])} ({newest[1].strftime('%Y-%m-%d %H:%M:%S')})")
            
            print("  文件年龄分布:")
            print(f"    1天内: {output_stats['files_by_age']['1_day']}个文件")
            print(f"    7天内(不含1天内): {output_stats['files_by_age']['7_days']}个文件")
            print(f"    30天内(不含7天内): {output_stats['files_by_age']['30_days']}个文件")
            print(f"    更旧: {output_stats['files_by_age']['older']}个文件")
    
    def _format_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} 字节"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes/1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes/(1024*1024):.2f} MB"
        else:
            return f"{size_bytes/(1024*1024*1024):.2f} GB"
    
    def clean_files(self, directory, days=None, pattern=None, extensions=None, exclude_patterns=None, force=False, test_mode=False):
        """
        清理指定目录中的文件
        
        参数:
            directory (str): 要清理的目录
            days (int): 保留的天数，超过这个天数的文件将被清理，None表示不考虑时间
            pattern (str): 文件名匹配模式（正则表达式）
            extensions (list): 要删除的文件扩展名列表，如['.jpg', '.xlsx']
            exclude_patterns (list): 要排除的文件名模式列表
            force (bool): 是否强制清理，不显示确认提示
            test_mode (bool): 测试模式，只显示要删除的文件而不实际删除
        
        返回:
            tuple: (cleaned_count, cleaned_size) 清理的文件数量和总大小
        """
        if not os.path.exists(directory):
            logger.warning(f"目录不存在: {directory}")
            return 0, 0
        
        cutoff_date = None
        if days is not None:
            cutoff_date = datetime.now() - timedelta(days=days)
        
        pattern_regex = re.compile(pattern) if pattern else None
        
        files_to_clean = []
        
        logger.info(f"扫描目录: {directory}")
        
        # 查找需要清理的文件
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                
                # 跳过临时文件
                if file.startswith('~$') or file.startswith('.'):
                    continue
                
                # 检查是否在排除列表中
                if exclude_patterns and any(pattern in file for pattern in exclude_patterns):
                    logger.info(f"跳过文件: {file}")
                    continue
                
                # 检查文件扩展名
                if extensions and not any(file.lower().endswith(ext.lower()) for ext in extensions):
                    continue
                
                # 检查修改时间
                if cutoff_date:
                    try:
                        mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                        if mod_time >= cutoff_date:
                            logger.debug(f"文件未超过保留天数: {file} - {mod_time.strftime('%Y-%m-%d %H:%M:%S')}")
                            continue
                    except Exception as e:
                        logger.error(f"检查文件时间时出错 {file_path}: {e}")
                        continue
                
                # 检查是否匹配模式
                if pattern_regex and not pattern_regex.search(file):
                    continue
                
                try:
                    file_size = os.path.getsize(file_path)
                    files_to_clean.append((file_path, file_size))
                    logger.info(f"找到要清理的文件: {file_path}")
                except Exception as e:
                    logger.error(f"获取文件大小时出错 {file_path}: {e}")
        
        if not files_to_clean:
            logger.info(f"没有找到需要清理的文件: {directory}")
            return 0, 0
        
        # 显示要清理的文件
        total_size = sum(f[1] for f in files_to_clean)
        print(f"\n找到 {len(files_to_clean)} 个文件要清理，总大小: {self._format_size(total_size)}")
        
        if len(files_to_clean) > 10:
            print("前10个文件:")
            for file_path, size in files_to_clean[:10]:
                print(f"  {os.path.basename(file_path)} ({self._format_size(size)})")
            print(f"  ...以及其他 {len(files_to_clean) - 10} 个文件")
        else:
            for file_path, size in files_to_clean:
                print(f"  {os.path.basename(file_path)} ({self._format_size(size)})")
        
        # 如果是测试模式，就不实际删除
        if test_mode:
            print("\n测试模式：不会实际删除文件。")
            return len(files_to_clean), total_size
        
        # 确认清理
        if not force:
            confirm = input(f"\n确定要清理这些文件吗？[y/N] ")
            if confirm.lower() != 'y':
                print("清理操作已取消。")
                return 0, 0
        
        # 执行清理
        cleaned_count = 0
        cleaned_size = 0
        
        for file_path, size in files_to_clean:
            try:
                # 删除文件
                try:
                    # 尝试检查文件是否被其他进程占用
                    if os.path.exists(file_path):
                        # 在Windows系统上，可能需要先关闭可能打开的文件句柄
                        if sys.platform == 'win32':
                            try:
                                # 尝试重命名文件，如果被占用通常会失败
                                temp_path = file_path + '.temp'
                                os.rename(file_path, temp_path)
                                os.rename(temp_path, file_path)
                            except Exception as e:
                                logger.warning(f"文件可能被占用: {file_path}, 错误: {e}")
                                # 尝试关闭文件句柄（仅Windows）
                                try:
                                    import ctypes
                                    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
                                    handle = kernel32.CreateFileW(file_path, 0x80000000, 0, None, 3, 0x80, None)
                                    if handle != -1:
                                        kernel32.CloseHandle(handle)
                                except Exception:
                                    pass
                        
                        # 使用Path对象删除文件
                        try:
                            Path(file_path).unlink(missing_ok=True)
                            logger.info(f"已删除文件: {file_path}")
                            
                            cleaned_count += 1
                            cleaned_size += size
                        except Exception as e1:
                            # 如果Path.unlink失败，尝试使用os.remove
                            try:
                                os.remove(file_path)
                                logger.info(f"使用os.remove删除文件: {file_path}")
                                
                                cleaned_count += 1
                                cleaned_size += size
                            except Exception as e2:
                                logger.error(f"删除文件失败 {file_path}: {e1}, 再次尝试: {e2}")
                    else:
                        logger.warning(f"文件不存在或已被删除: {file_path}")
                except Exception as e:
                    logger.error(f"删除文件时出错 {file_path}: {e}")
            except Exception as e:
                logger.error(f"处理文件时出错 {file_path}: {e}")
        
        print(f"\n已清理 {cleaned_count} 个文件，总大小: {self._format_size(cleaned_size)}")
        
        return cleaned_count, cleaned_size
    
    def clean_image_files(self, force=False, test_mode=False):
        """清理输入目录中的图片文件"""
        print(f"\n===== 清理输入目录图片文件 ({self.input_dir}) =====")
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
        return self.clean_files(
            self.input_dir, 
            days=None,  # 不考虑天数，清理所有图片
            extensions=image_extensions,
            force=force,
            test_mode=test_mode
        )
    
    def clean_excel_files(self, force=False, test_mode=False):
        """清理输出目录中的Excel文件"""
        print(f"\n===== 清理输出目录Excel文件 ({self.output_dir}) =====")
        excel_extensions = ['.xlsx', '.xls']
        exclude_patterns = ['processed_files.json']  # 保留处理记录文件
        return self.clean_files(
            self.output_dir,
            days=None,  # 不考虑天数，清理所有Excel
            extensions=excel_extensions,
            exclude_patterns=exclude_patterns,
            force=force,
            test_mode=test_mode
        )

    def clean_log_files(self, days=None, force=False, test_mode=False):
        """清理日志目录中的旧日志文件
        
        参数:
            days (int): 保留的天数，超过这个天数的日志将被清理，None表示清理所有日志
            force (bool): 是否强制清理，不显示确认提示
            test_mode (bool): 测试模式，只显示要删除的文件而不实际删除
        """
        print(f"\n===== 清理日志文件 ({self.logs_dir}) =====")
        log_extensions = ['.log']
        # 排除当前正在使用的日志文件
        current_log = os.path.basename(log_file)
        logger.info(f"当前使用的日志文件: {current_log}")
        
        result = self.clean_files(
            self.logs_dir,
            days=days,  # 如果days=None，清理所有日志文件
            extensions=log_extensions,
            exclude_patterns=[current_log],  # 排除当前使用的日志文件
            force=force,
            test_mode=test_mode
        )
        
        return result

    def clean_logs(self, days=7, force=False, test=False):
        """清理日志目录中的日志文件"""
        try:
            logs_dir = self.logs_dir
            if not os.path.exists(logs_dir):
                logger.warning(f"日志目录不存在: {logs_dir}")
                return

            cutoff_date = datetime.now() - timedelta(days=days)
            files_to_delete = []
            
            # 检查是否有活跃标记文件
            active_files = set()
            for marker_file in glob.glob(os.path.join(logs_dir, '*.active')):
                active_log_name = os.path.basename(marker_file).replace('.active', '.log')
                active_files.add(active_log_name)
                logger.info(f"检测到活跃日志文件: {active_log_name}")
            
            for file_path in glob.glob(os.path.join(logs_dir, '*.log*')):
                file_name = os.path.basename(file_path)
                
                # 跳过活跃的日志文件
                if file_name in active_files:
                    logger.info(f"跳过活跃日志文件: {file_name}")
                    continue
                
                mtime = os.path.getmtime(file_path)
                if datetime.fromtimestamp(mtime) < cutoff_date:
                    files_to_delete.append(file_path)

            if not files_to_delete:
                logger.info("没有找到需要清理的日志文件")
                return

            logger.info(f"找到 {len(files_to_delete)} 个过期的日志文件")
            for file_path in files_to_delete:
                if test:
                    logger.info(f"测试模式 - 将删除: {os.path.basename(file_path)}")
                else:
                    if not force:
                        response = input(f"是否删除日志文件 {os.path.basename(file_path)}? (y/n): ")
                        if response.lower() != 'y':
                            logger.info(f"已跳过 {os.path.basename(file_path)}")
                            continue
                    
                    try:
                        os.remove(file_path)
                        logger.info(f"已删除日志文件: {os.path.basename(file_path)}")
                    except Exception as e:
                        logger.error(f"删除文件失败: {file_path}, 错误: {e}")

        except Exception as e:
            logger.error(f"清理日志文件时出错: {e}")

    def clean_all_logs(self, force=False, test=False, except_current=True):
        """清理所有日志文件"""
        try:
            logs_dir = self.logs_dir
            if not os.path.exists(logs_dir):
                logger.warning(f"日志目录不存在: {logs_dir}")
                return
            
            # 检查是否有活跃标记文件
            active_files = set()
            for marker_file in glob.glob(os.path.join(logs_dir, '*.active')):
                active_log_name = os.path.basename(marker_file).replace('.active', '.log')
                active_files.add(active_log_name)
                logger.info(f"检测到活跃日志文件: {active_log_name}")
            
            files_to_delete = []
            for file_path in glob.glob(os.path.join(logs_dir, '*.log*')):
                file_name = os.path.basename(file_path)
                
                # 跳过当前正在使用的日志文件
                if except_current and file_name in active_files:
                    logger.info(f"保留活跃日志文件: {file_name}")
                    continue
                
                files_to_delete.append(file_path)
            
            if not files_to_delete:
                logger.info("没有找到需要清理的日志文件")
                return
            
            logger.info(f"找到 {len(files_to_delete)} 个日志文件需要清理")
            for file_path in files_to_delete:
                if test:
                    logger.info(f"测试模式 - 将删除: {os.path.basename(file_path)}")
                else:
                    if not force:
                        response = input(f"是否删除日志文件 {os.path.basename(file_path)}? (y/n): ")
                        if response.lower() != 'y':
                            logger.info(f"已跳过 {os.path.basename(file_path)}")
                            continue
                        
                    try:
                        os.remove(file_path)
                        logger.info(f"已删除日志文件: {os.path.basename(file_path)}")
                    except Exception as e:
                        logger.error(f"删除文件失败: {file_path}, 错误: {e}")
        
        except Exception as e:
            logger.error(f"清理所有日志文件时出错: {e}")

def main():
    """主程序"""
    parser = argparse.ArgumentParser(description='文件清理工具')
    parser.add_argument('--stats', action='store_true', help='显示文件统计信息')
    parser.add_argument('--clean-input', action='store_true', help='清理输入目录中超过指定天数的文件')
    parser.add_argument('--clean-output', action='store_true', help='清理输出目录中超过指定天数的文件')
    parser.add_argument('--clean-images', action='store_true', help='清理输入目录中的所有图片文件')
    parser.add_argument('--clean-excel', action='store_true', help='清理输出目录中的所有Excel文件')
    parser.add_argument('--clean-logs', action='store_true', help='清理日志目录中超过指定天数的日志文件')
    parser.add_argument('--clean-all-logs', action='store_true', help='清理所有日志文件（除当前使用的）')
    parser.add_argument('--days', type=int, default=30, help='保留的天数，默认30天')
    parser.add_argument('--log-days', type=int, default=7, help='保留的日志天数，默认7天')
    parser.add_argument('--pattern', type=str, help='文件名匹配模式（正则表达式）')
    parser.add_argument('--force', action='store_true', help='强制清理，不显示确认提示')
    parser.add_argument('--test', action='store_true', help='测试模式，只显示要删除的文件而不实际删除')
    parser.add_argument('--input-dir', type=str, default='input', help='指定输入目录')
    parser.add_argument('--output-dir', type=str, default='output', help='指定输出目录')
    parser.add_argument('--help-only', action='store_true', help='只显示帮助信息，不执行任何操作')
    parser.add_argument('--all', action='store_true', help='清理所有类型的文件（输入、输出和日志）')
    
    args = parser.parse_args()
    
    cleaner = FileCleaner(args.input_dir, args.output_dir)
    
    # 显示统计信息
    if args.stats:
        cleaner.print_stats()
    
    # 如果指定了--help-only，只显示帮助信息
    if args.help_only:
        parser.print_help()
        return
    
    # 如果指定了--all，清理所有类型的文件
    if args.all:
        cleaner.clean_image_files(args.force, args.test)
        cleaner.clean_excel_files(args.force, args.test)
        cleaner.clean_log_files(args.log_days, args.force, args.test)
        cleaner.clean_all_logs(args.force, args.test)
        return
    
    # 清理输入目录中的图片文件
    if args.clean_images or not any([args.stats, args.clean_input, args.clean_output, 
                                     args.clean_excel, args.clean_logs, args.clean_all_logs, args.help_only]):
        cleaner.clean_image_files(args.force, args.test)
    
    # 清理输出目录中的Excel文件
    if args.clean_excel or not any([args.stats, args.clean_input, args.clean_output, 
                                   args.clean_images, args.clean_logs, args.clean_all_logs, args.help_only]):
        cleaner.clean_excel_files(args.force, args.test)
    
    # 清理日志文件（按天数）
    if args.clean_logs:
        cleaner.clean_log_files(args.log_days, args.force, args.test)
    
    # 清理所有日志文件
    if args.clean_all_logs:
        cleaner.clean_all_logs(args.force, args.test)
    
    # 清理输入目录（按天数）
    if args.clean_input:
        print(f"\n===== 清理输入目录 ({args.input_dir}) =====")
        cleaner.clean_files(
            args.input_dir, 
            days=args.days, 
            pattern=args.pattern, 
            force=args.force, 
            test_mode=args.test
        )
    
    # 清理输出目录（按天数）
    if args.clean_output:
        print(f"\n===== 清理输出目录 ({args.output_dir}) =====")
        cleaner.clean_files(
            args.output_dir, 
            days=args.days, 
            pattern=args.pattern, 
            force=args.force, 
            test_mode=args.test
        )

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n程序已被用户中断")
    except Exception as e:
        logger.error(f"程序运行出错: {e}", exc_info=True)
        print(f"程序运行出错: {e}")
        print("请查看日志文件了解详细信息")
    sys.exit(0) 