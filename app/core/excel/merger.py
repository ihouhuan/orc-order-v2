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
            
            # 处理特殊情况：检查是否需要读取指定行作为标题行
            for header_row_idx in range(5):  # 检查前5行
                if len(df) <= header_row_idx:
                    continue
                
                potential_header = df.iloc[header_row_idx].astype(str)
                header_keywords = ['条码', '条形码', '商品条码', '商品名称', '规格', '单价', '数量', '金额', '单位', '必填']
                matches = sum(1 for keyword in header_keywords if any(keyword in str(val) for val in potential_header.values))
                
                if matches >= 3:  # 如果至少匹配3个关键词，认为是表头
                    logger.info(f"检测到表头在第 {header_row_idx+1} 行")
                    
                    # 使用此行作为列名，数据从下一行开始
                    header_row = potential_header
                    data_rows = df.iloc[header_row_idx+1:].reset_index(drop=True)
                    
                    # 为每一列分配名称（避免重复的列名）
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
                    logger.debug(f"重新构建的数据帧列名: {df.columns.tolist()}")
                    break
            
            # 定义可能的列名映射
            column_mapping = {
                '条码': ['条码', '条形码', '商品条码', 'barcode', '商品条形码', '条形码', '商品条码', '商品编码', '商品编号', '条形码', '条码（必填）'],
                '采购量': ['数量', '采购数量', '购买数量', '采购数量', '订单数量', '采购数量', '采购量（必填）', '采购量', '数量（必填）'],
                '采购单价': ['单价', '价格', '采购单价', '销售价', '采购单价（必填）', '单价（必填）', '价格（必填）'],
                '赠送量': ['赠送量', '赠品数量', '赠送数量', '赠品']
            }
            
            # 显示所有列名，用于调试
            all_columns = df.columns.tolist()
            logger.info(f"列名: {all_columns}")
            
            # 映射实际的列名
            mapped_columns = {}
            for target_col, possible_names in column_mapping.items():
                for col in all_columns:
                    # 清理列名以进行匹配
                    col_str = str(col).strip()
                    
                    # 直接匹配整个列名
                    if col_str in possible_names:
                        mapped_columns[target_col] = col
                        logger.info(f"直接匹配列名: {col_str} -> {target_col}")
                        break
                        
                    # 移除列名中的空白字符进行比较
                    clean_col = re.sub(r'\s+', '', col_str)
                    for name in possible_names:
                        clean_name = re.sub(r'\s+', '', name)
                        # 完全匹配
                        if clean_col == clean_name:
                            mapped_columns[target_col] = col
                            logger.info(f"清理后匹配列名: {col_str} -> {target_col}")
                            break
                        # 部分匹配（列名包含关键词）
                        elif clean_name in clean_col:
                            mapped_columns[target_col] = col
                            logger.info(f"部分匹配列名: {col_str} -> {target_col}")
                            break
                
                    if target_col in mapped_columns:
                        break
                        
                # 如果没有找到匹配，尝试模糊匹配
                if target_col not in mapped_columns:
                    for col in all_columns:
                        col_str = str(col).strip().lower()
                        for name in possible_names:
                            name_lower = name.lower()
                            if name_lower in col_str:
                                mapped_columns[target_col] = col
                                logger.info(f"模糊匹配列名: {col} -> {target_col}")
                                break
                        if target_col in mapped_columns:
                            break
            
            # 如果找到了必要的列，重命名列
            if mapped_columns:
                rename_dict = {mapped_columns[key]: key for key in mapped_columns}
                logger.info(f"列名重命名映射: {rename_dict}")
                df = df.rename(columns=rename_dict)
                logger.info(f"重命名后的列名: {df.columns.tolist()}")
            else:
                logger.warning(f"未找到可映射的列名: {file_path}")
            
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
                df['赠送量'] = 0
            
            # 选择并清理需要的列
            cleaned_df = pd.DataFrame()
            
            # 清理条码 - 确保是字符串且无小数点
            cleaned_df['条码'] = df['条码'].apply(lambda x: format_barcode(x) if pd.notna(x) else '')
            
            # 清理采购量 - 确保是数字
            cleaned_df['采购量'] = pd.to_numeric(df['采购量'], errors='coerce').fillna(0)
            
            # 清理单价 - 确保是数字并保留4位小数
            cleaned_df['采购单价'] = pd.to_numeric(df['采购单价'], errors='coerce').fillna(0).round(4)
            
            # 清理赠送量 - 确保是数字
            cleaned_df['赠送量'] = pd.to_numeric(df['赠送量'], errors='coerce').fillna(0)
            
            # 过滤无效行 - 条码为空或采购量为0的行跳过
            valid_df = cleaned_df[(cleaned_df['条码'] != '') & (cleaned_df['采购量'] > 0)]
            
            if len(valid_df) > 0:
                processed_dfs.append(valid_df)
                logger.info(f"处理文件 {i+1}: 有效记录 {len(valid_df)} 行")
            else:
                logger.warning(f"处理文件 {i+1}: 没有有效记录")
        
        if not processed_dfs:
            logger.warning("没有有效的数据帧用于合并")
            return None
        
        # 将所有数据帧合并
        merged_df = pd.concat(processed_dfs, ignore_index=True)
        
        # 按条码和单价分组，合并相同商品
        # 四舍五入到4位小数，避免浮点误差导致相同价格被当作不同价格
        merged_df['采购单价'] = merged_df['采购单价'].round(4)  
        
        # 对于同一条码和单价的商品，合并数量和赠送量
        result = merged_df.groupby(['条码', '采购单价'], as_index=False).agg({
            '采购量': 'sum',
            '赠送量': 'sum'
        })
        
        # 排序，按条码升序
        result = result.sort_values('条码').reset_index(drop=True)
        
        # 设置为0的赠送量设为空
        result.loc[result['赠送量'] == 0, '赠送量'] = pd.NA
        
        logger.info(f"合并完成，共 {len(result)} 条商品记录")
        return result
    
    def create_merged_purchase_order(self, df: pd.DataFrame) -> Optional[str]:
        """
        创建合并的采购单文件，完全按照银豹格式要求
        
        Args:
            df: 合并后的数据帧
            
        Returns:
            输出文件路径，如果创建失败则返回None
        """
        try:
            # 打开模板文件
            template_workbook = xlrd.open_workbook(self.template_path, formatting_info=True)
            template_sheet = template_workbook.sheet_by_index(0)
            
            # 首先分析模板结构，确定关键列的位置
            logger.info(f"分析模板结构")
            for i in range(min(5, template_sheet.nrows)):
                row_values = [str(cell.value).strip() for cell in template_sheet.row(i)]
                logger.debug(f"模板第{i+1}行: {row_values}")
            
            # 银豹模板的标准列位置：
            # 条码列(商品条码): B列(索引1)
            barcode_col = 1
            # 采购量列: C列(索引2)
            quantity_col = 2 
            # 赠送量列: D列(索引3)
            gift_col = 3
            # 采购单价列: E列(索引4)
            price_col = 4
            
            # 找到数据开始行 - 通常是第二行(索引1)
            data_start_row = 1
            
            # 创建可写的副本
            output_workbook = xlcopy(template_workbook)
            output_sheet = output_workbook.get_sheet(0)
            
            # 设置单价的格式样式（保留4位小数）
            price_style = xlwt.XFStyle()
            price_style.num_format_str = '0.0000'
            
            # 数量格式
            quantity_style = xlwt.XFStyle()
            quantity_style.num_format_str = '0'
            
            # 遍历数据并填充到Excel
            for i, (_, row) in enumerate(df.iterrows()):
                r = data_start_row + i
                
                # 只填充银豹采购单格式要求的4个列：条码、采购量、赠送量、采购单价
                
                # 条码（必填）- B列(1)
                output_sheet.write(r, barcode_col, row['条码'])
                
                # 采购量（必填）- C列(2)
                output_sheet.write(r, quantity_col, float(row['采购量']), quantity_style)
                
                # 赠送量 - D列(3)
                if pd.notna(row['赠送量']) and float(row['赠送量']) > 0:
                    output_sheet.write(r, gift_col, float(row['赠送量']), quantity_style)
                
                # 采购单价（必填）- E列(4)
                output_sheet.write(r, price_col, float(row['采购单价']), price_style)
            
            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            output_file = os.path.join(self.output_dir, f"合并采购单_{timestamp}.xls")
            
            # 保存文件
            output_workbook.save(output_file)
            logger.info(f"合并采购单已保存到: {output_file}，共{len(df)}条记录")
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