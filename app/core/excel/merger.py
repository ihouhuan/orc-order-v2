"""
订单合并模块
----------
提供采购单合并功能，将多个采购单合并为一个。
"""

import os
import re
import pandas as pd
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy as xlcopy
from typing import Dict, List, Optional, Tuple, Union, Any
from datetime import datetime

from ...config.settings import ConfigManager
from ..utils.log_utils import get_logger
from ..utils.file_utils import (
    ensure_dir,
    get_file_extension,
    get_files_by_extensions,
    load_json,
    save_json
)
from ..utils.string_utils import (
    clean_string,
    clean_barcode,
    format_barcode
)

logger = get_logger(__name__)

class PurchaseOrderMerger:
    """
    采购单合并器：将多个采购单Excel文件合并成一个文件
    """
    
    def __init__(self, config: Optional[ConfigManager] = None):
        """
        初始化采购单合并器
        
        Args:
            config: 配置管理器，如果为None则创建新的
        """
        logger.info("初始化PurchaseOrderMerger")
        self.config = config or ConfigManager()
        
        # 获取配置
        self.output_dir = self.config.get_path('Paths', 'output_folder', 'data/output', create=True)
        
        # 获取模板文件路径
        template_folder = self.config.get('Paths', 'template_folder', 'templates')
        template_name = self.config.get('Templates', 'purchase_order', '银豹-采购单模板.xls')
        
        self.template_path = os.path.join(template_folder, template_name)
        
        # 检查模板文件是否存在
        if not os.path.exists(self.template_path):
            logger.error(f"模板文件不存在: {self.template_path}")
            raise FileNotFoundError(f"模板文件不存在: {self.template_path}")
        
        # 用于记录已合并的文件
        self.cache_file = os.path.join(self.output_dir, "merged_files.json")
        self.merged_files = self._load_merged_files()
        
        logger.info(f"初始化完成，模板文件: {self.template_path}")
    
    def _load_merged_files(self) -> Dict[str, str]:
        """
        加载已合并文件的缓存
        
        Returns:
            合并记录字典
        """
        return load_json(self.cache_file, {})
        
    def _save_merged_files(self) -> None:
        """保存已合并文件的缓存"""
        save_json(self.merged_files, self.cache_file)
    
    def get_purchase_orders(self) -> List[str]:
        """
        获取output目录下的采购单Excel文件
        
        Returns:
            采购单文件路径列表
        """
        logger.info(f"搜索目录 {self.output_dir} 中的采购单Excel文件")
        
        # 获取所有Excel文件
        all_files = get_files_by_extensions(self.output_dir, ['.xls', '.xlsx'])
        
        # 筛选采购单文件
        purchase_orders = [
            file for file in all_files 
            if os.path.basename(file).startswith('采购单_')
        ]
        
        if not purchase_orders:
            logger.warning(f"未在 {self.output_dir} 目录下找到采购单Excel文件")
            return []
        
        # 按修改时间排序，最新的在前
        purchase_orders.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        
        logger.info(f"找到 {len(purchase_orders)} 个采购单Excel文件")
        return purchase_orders
    
    def read_purchase_order(self, file_path: str) -> Optional[pd.DataFrame]:
        """
        读取采购单Excel文件
        
        Args:
            file_path: 采购单文件路径
            
        Returns:
            数据帧，如果读取失败则返回None
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            logger.info(f"成功读取采购单文件: {file_path}")
            
            # 打印列名，用于调试
            logger.debug(f"Excel文件的列名: {df.columns.tolist()}")
            
            # 检查是否有特殊表头结构（如在第3行）
            special_header = False
            if len(df) > 3:  # 确保有足够的行
                row3 = df.iloc[3].astype(str)
                header_keywords = ['行号', '条形码', '条码', '商品名称', '规格', '单价', '数量', '金额', '单位']
                # 计算匹配的关键词数量
                matches = sum(1 for keyword in header_keywords if any(keyword in str(val) for val in row3.values))
                # 如果匹配了至少3个关键词，认为第3行是表头
                if matches >= 3:
                    logger.info(f"检测到特殊表头结构，使用第3行作为列名")
                    # 创建新的数据帧，使用第3行作为列名，数据从第4行开始
                    header_row = df.iloc[3]
                    data_rows = df.iloc[4:].reset_index(drop=True)
                    # 为每一列分配一个名称（避免重复的列名）
                    new_columns = []
                    for i, col in enumerate(header_row):
                        col_str = str(col)
                        if col_str == 'nan' or col_str == 'None' or pd.isna(col):
                            new_columns.append(f"Col_{i}")
                        else:
                            new_columns.append(col_str)
                    # 使用新列名创建新的DataFrame
                    data_rows.columns = new_columns
                    df = data_rows
                    special_header = True
                    logger.debug(f"重新构建的数据帧列名: {df.columns.tolist()}")
            
            # 定义可能的列名映射
            column_mapping = {
                '条码': ['条码', '条形码', '商品条码', 'barcode', '商品条形码', '条形码', '商品条码', '商品编码', '商品编号', '条形码', '条码（必填）'],
                '采购量': ['数量', '采购数量', '购买数量', '采购数量', '订单数量', '采购数量', '采购量（必填）'],
                '采购单价': ['单价', '价格', '采购单价', '销售价', '采购单价（必填）'],
                '赠送量': ['赠送量', '赠品数量', '赠送数量', '赠品']
            }
            
            # 映射实际的列名
            mapped_columns = {}
            for target_col, possible_names in column_mapping.items():
                for col in df.columns:
                    # 移除列名中的空白字符和括号内容以进行比较
                    clean_col = re.sub(r'\s+', '', str(col))
                    clean_col = re.sub(r'（.*?）', '', clean_col)  # 移除括号内容
                    for name in possible_names:
                        clean_name = re.sub(r'\s+', '', name)
                        clean_name = re.sub(r'（.*?）', '', clean_name)  # 移除括号内容
                        if clean_col == clean_name:
                            mapped_columns[target_col] = col
                            break
                    if target_col in mapped_columns:
                        break
            
            # 如果找到了必要的列，重命名列
            if mapped_columns:
                # 如果没有找到条码列，无法继续处理
                if '条码' not in mapped_columns:
                    logger.error(f"未找到条码列: {file_path}")
                    return None
                    
                df = df.rename(columns=mapped_columns)
                logger.info(f"列名映射结果: {mapped_columns}")
            
            return df
            
        except Exception as e:
            logger.error(f"读取采购单文件失败: {file_path}, 错误: {str(e)}")
            return None
    
    def merge_purchase_orders(self, file_paths: List[str]) -> Optional[pd.DataFrame]:
        """
        合并多个采购单文件
        
        Args:
            file_paths: 采购单文件路径列表
            
        Returns:
            合并后的数据帧，如果合并失败则返回None
        """
        if not file_paths:
            logger.warning("没有需要合并的采购单文件")
            return None
        
        # 读取所有采购单文件
        dfs = []
        for file_path in file_paths:
            df = self.read_purchase_order(file_path)
            if df is not None:
                dfs.append(df)
        
        if not dfs:
            logger.warning("没有成功读取的采购单文件")
            return None
        
        # 合并数据
        logger.info(f"开始合并 {len(dfs)} 个采购单文件")
        
        # 首先，整理每个数据帧以确保它们有相同的结构
        processed_dfs = []
        for i, df in enumerate(dfs):
            # 确保必要的列存在
            required_columns = ['条码', '采购量', '采购单价']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                logger.warning(f"数据帧 {i} 缺少必要的列: {missing_columns}")
                continue
            
            # 处理赠送量列不存在的情况
            if '赠送量' not in df.columns:
                df['赠送量'] = pd.NA
            
            # 选择需要的列
            selected_df = df[['条码', '采购量', '采购单价', '赠送量']].copy()
            
            # 清理和转换数据
            selected_df['条码'] = selected_df['条码'].apply(lambda x: format_barcode(x) if pd.notna(x) else x)
            selected_df['采购量'] = pd.to_numeric(selected_df['采购量'], errors='coerce')
            selected_df['采购单价'] = pd.to_numeric(selected_df['采购单价'], errors='coerce')
            selected_df['赠送量'] = pd.to_numeric(selected_df['赠送量'], errors='coerce')
            
            # 过滤无效行
            valid_df = selected_df.dropna(subset=['条码', '采购量'])
            
            processed_dfs.append(valid_df)
        
        if not processed_dfs:
            logger.warning("没有有效的数据帧用于合并")
            return None
        
        # 将所有数据帧合并
        merged_df = pd.concat(processed_dfs, ignore_index=True)
        
        # 按条码和单价分组，合并相同商品
        merged_df['采购单价'] = merged_df['采购单价'].round(4)  # 四舍五入到4位小数，避免浮点误差
        
        # 对于同一条码和单价的商品，合并数量和赠送量
        grouped = merged_df.groupby(['条码', '采购单价'], as_index=False).agg({
            '采购量': 'sum',
            '赠送量': lambda x: sum(x.dropna()) if len(x.dropna()) > 0 else pd.NA
        })
        
        # 计算其他信息
        grouped['采购金额'] = grouped['采购量'] * grouped['采购单价']
        
        # 排序，按条码升序
        result = grouped.sort_values('条码').reset_index(drop=True)
        
        logger.info(f"合并完成，共 {len(result)} 条商品记录")
        return result
    
    def create_merged_purchase_order(self, df: pd.DataFrame) -> Optional[str]:
        """
        创建合并的采购单文件
        
        Args:
            df: 合并后的数据帧
            
        Returns:
            输出文件路径，如果创建失败则返回None
        """
        try:
            # 打开模板文件
            template_workbook = xlrd.open_workbook(self.template_path, formatting_info=True)
            template_sheet = template_workbook.sheet_by_index(0)
            
            # 创建可写的副本
            output_workbook = xlcopy(template_workbook)
            output_sheet = output_workbook.get_sheet(0)
            
            # 填充商品信息
            start_row = 4  # 从第5行开始填充数据（索引从0开始）
            
            for i, (_, row) in enumerate(df.iterrows()):
                r = start_row + i
                
                # 序号
                output_sheet.write(r, 0, i + 1)
                # 商品编码（条码）
                output_sheet.write(r, 1, row['条码'])
                # 商品名称（合并单没有名称信息，留空）
                output_sheet.write(r, 2, "")
                # 规格（合并单没有规格信息，留空）
                output_sheet.write(r, 3, "")
                # 单位（合并单没有单位信息，留空）
                output_sheet.write(r, 4, "")
                # 单价
                output_sheet.write(r, 5, row['采购单价'])
                # 采购数量
                output_sheet.write(r, 6, row['采购量'])
                # 采购金额
                output_sheet.write(r, 7, row['采购金额'])
                # 税率
                output_sheet.write(r, 8, 0)
                # 赠送量
                if pd.notna(row['赠送量']):
                    output_sheet.write(r, 9, row['赠送量'])
                else:
                    output_sheet.write(r, 9, "")
            
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            output_file = os.path.join(self.output_dir, f"合并采购单_{timestamp}.xls")
            
            # 保存文件
            output_workbook.save(output_file)
            logger.info(f"合并采购单已保存到: {output_file}")
            return output_file
            
        except Exception as e:
            logger.error(f"创建合并采购单时出错: {e}")
            return None
    
    def process(self, file_paths: Optional[List[str]] = None) -> Optional[str]:
        """
        处理采购单合并
        
        Args:
            file_paths: 指定要合并的文件路径列表，如果为None则自动获取
            
        Returns:
            合并后的文件路径，如果合并失败则返回None
        """
        # 如果未指定文件路径，则获取所有采购单文件
        if file_paths is None:
            file_paths = self.get_purchase_orders()
        
        # 检查是否有文件需要合并
        if not file_paths:
            logger.warning("没有找到可合并的采购单文件")
            return None
        
        # 合并采购单
        merged_df = self.merge_purchase_orders(file_paths)
        if merged_df is None:
            logger.error("合并采购单失败")
            return None
        
        # 创建合并的采购单文件
        output_file = self.create_merged_purchase_order(merged_df)
        if output_file is None:
            logger.error("创建合并采购单文件失败")
            return None
        
        # 记录已合并文件
        for file_path in file_paths:
            self.merged_files[file_path] = output_file
        self._save_merged_files()
        
        return output_file 