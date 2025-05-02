#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
合并采购单程序
-------------------
将多个采购单Excel文件合并成一个文件。
"""

import os
import sys
import logging
import pandas as pd
import xlrd
import xlwt
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Any
from datetime import datetime
import random
from xlutils.copy import copy as xlcopy
import time
import json
import re

# 配置日志
logger = logging.getLogger(__name__)
if not logger.handlers:
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs', 'merge_purchase_orders.log')
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
logger.info("初始化日志系统")

class PurchaseOrderMerger:
    """
    采购单合并器：将多个采购单Excel文件合并成一个文件
    """
    
    def __init__(self, output_dir="output"):
        """
        初始化采购单合并器，并设置输出目录
        """
        logger.info("初始化PurchaseOrderMerger")
        self.output_dir = output_dir
        
        # 确保输出目录存在
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logger.info(f"创建输出目录: {output_dir}")
    
        # 设置路径
        self.template_path = os.path.join("templets", "银豹-采购单模板.xls")
        
        # 检查模板文件是否存在
        if not os.path.exists(self.template_path):
            logger.error(f"模板文件不存在: {self.template_path}")
            raise FileNotFoundError(f"模板文件不存在: {self.template_path}")
        
        # 用于记录已处理的文件
        self.cache_file = os.path.join(output_dir, "merged_files.json")
        self.merged_files = self._load_merged_files()
        
        logger.info(f"初始化完成，模板文件: {self.template_path}")
    
    def _load_merged_files(self):
        """加载已合并文件的缓存"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                logger.info(f"加载已合并文件缓存，共{len(cache)} 条记录")
                return cache
            except Exception as e:
                logger.warning(f"读取缓存文件失败: {e}")
        return {}
        
    def _save_merged_files(self):
        """保存已合并文件的缓存"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.merged_files, f, ensure_ascii=False, indent=2)
            logger.info(f"已更新合并文件缓存，共{len(self.merged_files)} 条记录")
        except Exception as e:
            logger.warning(f"保存缓存文件失败: {e}")
    
    def get_latest_purchase_orders(self):
        """
        获取output目录下最新的采购单Excel文件
        """
        logger.info(f"搜索目录 {self.output_dir} 中的采购单Excel文件")
        excel_files = []
        
        for file in os.listdir(self.output_dir):
            # 只处理以"采购单_"开头的Excel文件
            if file.lower().endswith('.xls') and file.startswith('采购单_'):
                file_path = os.path.join(self.output_dir, file)
                excel_files.append((file_path, os.path.getmtime(file_path)))
        
        if not excel_files:
            logger.warning(f"未在 {self.output_dir} 目录下找到采购单Excel文件")
            return []
        
        # 按修改时间排序，获取最新的文件
        sorted_files = sorted(excel_files, key=lambda x: x[1], reverse=True)
        logger.info(f"找到{len(sorted_files)} 个采购单Excel文件")
        return [file[0] for file in sorted_files]
    
    def read_purchase_order(self, file_path):
        """
        读取采购单Excel文件
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            logger.info(f"成功读取采购单文件: {file_path}")
            
            # 打印列名，用于调试
            logger.info(f"Excel文件的列名: {df.columns.tolist()}")
            
            # 检查是否有特殊表头结构（如"武侯环球乐百惠便利店3333.xlsx"）
            # 判断依据：检查第3行是否包含常见的商品表头信息
            special_header = False
            if len(df) > 3:  # 确保有足够的行
                row3 = df.iloc[3].astype(str)
                header_keywords = ['行号', '条形码', '条码', '商品名称', '规格', '单价', '数量', '金额', '单位']
                # 计算匹配的关键词数量
                matches = sum(1 for keyword in header_keywords if any(keyword in str(val) for val in row3.values))
                # 如果匹配了至少3个关键词，认为第3行是表头
                if matches >= 3:
                    logger.info(f"检测到特殊表头结构，使用第3行作为列名: {row3.values.tolist()}")
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
                    logger.info(f"重新构建的数据帧列名: {df.columns.tolist()}")
            
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
                df = df.rename(columns=mapped_columns)
                logger.info(f"列名映射结果: {mapped_columns}")
            
            return df
        except Exception as e:
            logger.error(f"读取采购单文件失败: {file_path}, 错误: {str(e)}")
            return None
    
    def merge_purchase_orders(self, file_paths):
        """
        合并多个采购单文件
        """
        if not file_paths:
            logger.warning("没有需要合并的采购单文件")
            return None
        
        # 读取所有采购单文件
        dfs = []
        for file_path in file_paths:
            df = self.read_purchase_order(file_path)
            if df is not None:
                # 确保条码列是字符串类型
                df['条码（必填）'] = df['条码（必填）'].astype(str)
                # 去除可能的小数点和.0
                df['条码（必填）'] = df['条码（必填）'].apply(lambda x: x.split('.')[0] if '.' in x else x)
                
                # 处理NaN值，将其转换为空字符串
                for col in df.columns:
                    df[col] = df[col].apply(lambda x: '' if pd.isna(x) else x)
                
                dfs.append(df)
        
        if not dfs:
            logger.error("没有成功读取任何采购单文件")
            return None
        
        # 合并所有数据框
        merged_df = pd.concat(dfs, ignore_index=True)
        logger.info(f"合并了{len(dfs)} 个采购单文件，共{len(merged_df)} 条记录")
        
        # 检查并合并相同条码和单价的数据
        merged_data = {}
        for _, row in merged_df.iterrows():
            # 使用映射后的列名访问数据
            barcode = str(row['条码（必填）'])  # 保持字符串格式
            # 移除条码中可能的小数点
            barcode = barcode.split('.')[0] if '.' in barcode else barcode
            
            unit_price = float(row['采购单价（必填）'])
            quantity = float(row['采购量（必填）'])
            
            # 检查赠送量是否为空
            has_gift = '赠送量' in row and row['赠送量'] != '' and not pd.isna(row['赠送量'])
            gift_quantity = float(row['赠送量']) if has_gift else ''
            
            # 商品名称处理，确保不会出现"nan"
            product_name = row['商品名称']
            if pd.isna(product_name) or product_name == 'nan' or product_name == 'None':
                product_name = ''
            
            # 创建唯一键：条码+单价
            key = f"{barcode}_{unit_price}"
            
            if key in merged_data:
                # 如果已存在相同条码和单价的数据，累加数量
                merged_data[key]['采购量（必填）'] += quantity
                
                # 如果当前记录有赠送量且之前的记录也有赠送量，则累加赠送量
                if has_gift and merged_data[key]['赠送量'] != '':
                    merged_data[key]['赠送量'] += gift_quantity
                # 如果当前记录有赠送量但之前的记录没有，则设置赠送量
                elif has_gift:
                    merged_data[key]['赠送量'] = gift_quantity
                # 其他情况保持原样（为空）
                
                logger.info(f"合并相同条码和单价的数据: 条码={barcode}, 单价={unit_price}, 数量={quantity}, 赠送量={gift_quantity}")
                
                # 如果当前商品名称不为空，且原来的为空，则更新商品名称
                if product_name and not merged_data[key]['商品名称']:
                    merged_data[key]['商品名称'] = product_name
            else:
                # 如果是新数据，直接添加
                merged_data[key] = {
                    '商品名称': product_name,
                    '条码（必填）': barcode,  # 使用处理后的条码
                    '采购量（必填）': quantity,
                    '赠送量': gift_quantity,
                    '采购单价（必填）': unit_price
                }
        
        # 将合并后的数据转换回DataFrame
        final_df = pd.DataFrame(list(merged_data.values()))
        logger.info(f"合并后剩余{len(final_df)} 条唯一记录")
        
        return final_df
    
    def create_merged_purchase_order(self, df):
        """
        创建合并后的采购单Excel文件
        """
        try:
            # 获取当前时间戳
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            
            # 创建输出文件路径
            output_file = os.path.join(self.output_dir, f"合并采购单_{timestamp}.xls")
            
            # 打开模板文件
            workbook = xlrd.open_workbook(self.template_path)
            workbook = xlcopy(workbook)
            worksheet = workbook.get_sheet(0)
            
            # 从第2行开始填充数据
            row_index = 1
            
            # 按条码排序
            df = df.sort_values('条码（必填）')
            
            # 填充数据
            for _, row in df.iterrows():
                # 1. 列A(0): 商品名称
                product_name = str(row['商品名称'])
                # 检查并处理nan值
                if product_name == 'nan' or product_name == 'None':
                    product_name = ''
                worksheet.write(row_index, 0, product_name)
                
                # 2. 列B(1): 条码
                worksheet.write(row_index, 1, str(row['条码（必填）']))
                
                # 3. 列C(2): 采购量
                worksheet.write(row_index, 2, float(row['采购量（必填）']))
                
                # 4. 列D(3): 赠送量
                # 只有当赠送量不为空且不为0时才写入
                if '赠送量' in row and row['赠送量'] != '' and not pd.isna(row['赠送量']):
                    # 将赠送量转换为数字
                    try:
                        gift_quantity = float(row['赠送量'])
                        # 只有当赠送量大于0时才写入
                        if gift_quantity > 0:
                            worksheet.write(row_index, 3, gift_quantity)
                    except (ValueError, TypeError):
                        # 如果转换失败，忽略赠送量
                        pass
                
                # 5. 列E(4): 采购单价
                        style = xlwt.XFStyle()
                style.num_format_str = '0.0000'
                worksheet.write(row_index, 4, float(row['采购单价（必填）']), style)
                
                row_index += 1
            
            # 保存文件
            workbook.save(output_file)
            logger.info(f"合并采购单已保存: {output_file}")
            
            # 记录已合并文件
            for file_path in self.get_latest_purchase_orders():
                file_stat = os.stat(file_path)
                file_key = f"{os.path.basename(file_path)}_{file_stat.st_size}_{file_stat.st_mtime}"
                self.merged_files[file_key] = output_file
            
            self._save_merged_files()
            
            return output_file
            
        except Exception as e:
            logger.error(f"创建合并采购单失败: {str(e)}")
            return None
    
    def process(self):
        """
        处理最新的采购单文件
        """
        # 获取最新的采购单文件
        file_paths = self.get_latest_purchase_orders()
        if not file_paths:
            logger.error("未找到可处理的采购单文件")
            return False
        
        # 合并采购单
        merged_df = self.merge_purchase_orders(file_paths)
        if merged_df is None:
            logger.error("合并采购单失败")
                return False
            
        # 创建合并后的采购单
        output_file = self.create_merged_purchase_order(merged_df)
        if output_file is None:
            logger.error("创建合并采购单失败")
            return False
        
        logger.info(f"处理完成，合并采购单已保存至: {output_file}")
        return True

def main():
    """主程序"""
    import argparse
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='合并采购单程序')
    parser.add_argument('--input', type=str, help='指定输入采购单文件路径，多个文件用逗号分隔')
    args = parser.parse_args()
    
    merger = PurchaseOrderMerger()
    
    # 处理采购单文件
    try:
        if args.input:
            # 使用指定文件处理
            file_paths = [path.strip() for path in args.input.split(',')]
            merged_df = merger.merge_purchase_orders(file_paths)
            if merged_df is not None:
                output_file = merger.create_merged_purchase_order(merged_df)
                if output_file:
                    print(f"处理成功！合并采购单已保存至: {output_file}")
                else:
                    print("处理失败！请查看日志了解详细信息")
            else:
                print("处理失败！请查看日志了解详细信息")
        else:
            # 使用默认处理流程（查找最新文件）
            result = merger.process()
            if result:
                print("处理成功！已将数据合并并保存")
            else:
                print("处理失败！请查看日志了解详细信息")
    except Exception as e:
        logger.error(f"处理过程中发生错误: {e}", exc_info=True)
        print(f"处理过程中发生错误: {e}")
        print("请查看日志文件了解详细信息")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"程序执行过程中发生错误: {e}", exc_info=True)
        sys.exit(1) 