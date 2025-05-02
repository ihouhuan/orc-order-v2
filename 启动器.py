#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
益选-OCR订单处理系统启动器
-----------------
提供简单的图形界面，方便用户选择功能
"""

import os
import sys
import time
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext, ttk
from tkinter import font as tkfont
from threading import Thread
import datetime
import json
import re
from typing import Dict, List, Optional, Any

# 全局变量，用于跟踪任务状态
RUNNING_TASK = None
THEME_MODE = "light"  # 默认浅色主题

# 定义浅色和深色主题颜色
THEMES = {
    "light": {
        "bg": "#f0f0f0",
        "fg": "#000000",
        "button_bg": "#e0e0e0",
        "button_fg": "#000000",
        "log_bg": "#ffffff",
        "log_fg": "#000000",
        "highlight_bg": "#4a6984",
        "highlight_fg": "#ffffff",
        "border": "#cccccc",
        "success": "#28a745",
        "error": "#dc3545",
        "warning": "#ffc107",
        "info": "#17a2b8"
    },
    "dark": {
        "bg": "#2d2d2d",
        "fg": "#ffffff",
        "button_bg": "#444444",
        "button_fg": "#ffffff",
        "log_bg": "#1e1e1e",
        "log_fg": "#e0e0e0",
        "highlight_bg": "#4a6984",
        "highlight_fg": "#ffffff",
        "border": "#555555",
        "success": "#28a745",
        "error": "#dc3545",
        "warning": "#ffc107",
        "info": "#17a2b8"
    }
}

class StatusBar(tk.Frame):
    """状态栏，显示当前系统状态和进度"""
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(height=25, relief=tk.SUNKEN, borderwidth=1)
        
        # 状态标签
        self.status_label = tk.Label(self, text="就绪", anchor=tk.W, padx=5)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 进度条
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=5, pady=2)
        
        # 隐藏进度条（初始状态）
        self.progress.pack_forget()
        
    def set_status(self, text, progress=None):
        """设置状态栏文本和进度"""
        self.status_label.config(text=text)
        
        if progress is not None and 0 <= progress <= 100:
            self.progress.pack(side=tk.RIGHT, padx=5, pady=2)
            self.progress.config(value=progress)
        else:
            self.progress.pack_forget()
            
    def set_running(self, is_running=True):
        """设置运行状态"""
        if is_running:
            self.status_label.config(text="处理中...", foreground=THEMES[THEME_MODE]["info"])
            self.progress.pack(side=tk.RIGHT, padx=5, pady=2)
            self.progress.config(mode='indeterminate')
            self.progress.start()
        else:
            self.status_label.config(text="就绪", foreground=THEMES[THEME_MODE]["fg"])
            self.progress.stop()
            self.progress.pack_forget()
            
def run_command_with_logging(command, log_widget, status_bar=None, on_complete=None):
    """运行命令并将输出重定向到日志窗口"""
    global RUNNING_TASK
    
    # 如果已有任务在运行，提示用户
    if RUNNING_TASK is not None:
        messagebox.showinfo("任务进行中", "请等待当前任务完成后再执行新的操作。")
        return
        
    def run_in_thread():
        global RUNNING_TASK
        RUNNING_TASK = command
        
        # 更新状态栏
        if status_bar:
            status_bar.set_running(True)
        
        # 记录命令开始执行的时间
        start_time = datetime.datetime.now()
        log_widget.configure(state=tk.NORMAL)
        log_widget.delete(1.0, tk.END)  # 清空之前的日志
        log_widget.insert(tk.END, f"执行命令: {' '.join(command)}\n", "command")
        log_widget.insert(tk.END, f"开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n", "time")
        log_widget.insert(tk.END, "=" * 50 + "\n\n", "separator")
        log_widget.configure(state=tk.DISABLED)
        
        # 获取原始的stdout和stderr
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        
        # 创建日志重定向器
        log_redirector = LogRedirector(log_widget)
        
        # 设置环境变量，强制OCR模块输出到data目录
        env = os.environ.copy()
        env["OCR_OUTPUT_DIR"] = os.path.abspath("data/output")
        env["OCR_INPUT_DIR"] = os.path.abspath("data/input")
        env["OCR_LOG_LEVEL"] = "DEBUG"  # 设置更详细的日志级别
        
        try:
            # 重定向stdout和stderr到日志重定向器
            sys.stdout = log_redirector
            sys.stderr = log_redirector
            
            # 打印一条消息，确认重定向已生效
            print("日志重定向已启动，现在同时输出到终端和GUI")
            
            # 运行命令并捕获输出
            process = subprocess.Popen(
                command, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                env=env
            )
            
            output_data = []
            # 读取并显示输出
            for line in process.stdout:
                output_data.append(line)
                print(line.rstrip())  # 直接打印到已重定向的stdout
                
                # 尝试从输出中提取进度信息
                if status_bar:
                    progress = extract_progress_from_log(line)
                    if progress is not None:
                        log_widget.after(0, lambda p=progress: status_bar.set_status(f"处理中: {p}%完成", p))
                
            # 等待进程结束
            process.wait()
            
            # 记录命令结束时间
            end_time = datetime.datetime.now()
            duration = end_time - start_time
            
            print(f"\n{'=' * 50}")
            print(f"执行完毕！返回码: {process.returncode}")
            print(f"结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"耗时: {duration.total_seconds():.2f} 秒")
            
            # 获取输出内容
            output_text = ''.join(output_data)
            
            # 检查是否是完整流程命令且遇到了"未找到可合并的文件"的情况
            is_pipeline = "pipeline" in command
            no_merge_files = "未找到采购单文件" in output_text
            single_file = "只有1个采购单文件" in output_text
            
            # 如果是完整流程且只是没有找到可合并文件或只有一个文件，则视为成功
            if is_pipeline and (no_merge_files or single_file):
                print("完整流程中没有需要合并的文件，但其他步骤执行成功，视为成功完成")
                if status_bar:
                    log_widget.after(0, lambda: status_bar.set_status("处理完成", 100))
                log_widget.after(0, lambda: show_result_preview(command, output_text))
            else:
                # 执行完成后处理结果
                if on_complete:
                    log_widget.after(0, lambda: on_complete(process.returncode, output_text))
                
                # 如果处理成功，显示成功信息
                if process.returncode == 0:
                    if status_bar:
                        log_widget.after(0, lambda: status_bar.set_status("处理完成", 100))
                    log_widget.after(0, lambda: show_result_preview(command, output_text))
                else:
                    if status_bar:
                        log_widget.after(0, lambda: status_bar.set_status(f"处理失败 (返回码: {process.returncode})", 0))
                    log_widget.after(0, lambda: messagebox.showerror("操作失败", f"处理失败，返回码：{process.returncode}"))
                
        except Exception as e:
            print(f"\n执行出错: {str(e)}")
            if status_bar:
                log_widget.after(0, lambda: status_bar.set_status(f"执行出错: {str(e)}", 0))
            log_widget.after(0, lambda: messagebox.showerror("执行错误", f"执行命令时出错: {str(e)}"))
        finally:
            # 恢复原始stdout和stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            
            # 任务完成，重置状态
            RUNNING_TASK = None
            if status_bar:
                log_widget.after(0, lambda: status_bar.set_running(False))
    
    # 在新线程中运行，避免UI阻塞
    Thread(target=run_in_thread).start()

def extract_progress_from_log(log_line):
    """从日志行中提取进度信息"""
    # 尝试匹配"处理批次 x/y"格式的进度信息
    batch_match = re.search(r'处理批次 (\d+)/(\d+)', log_line)
    if batch_match:
        current = int(batch_match.group(1))
        total = int(batch_match.group(2))
        return int(current / total * 100)
    
    # 尝试匹配百分比格式
    percent_match = re.search(r'(\d+)%', log_line)
    if percent_match:
        return int(percent_match.group(1))
    
    return None

def show_result_preview(command, output):
    """显示处理结果预览"""
    # 根据命令类型提取不同的结果信息
    if "ocr" in command:
        show_ocr_result_preview(output)
    elif "excel" in command:
        show_excel_result_preview(output)
    elif "merge" in command:
        show_merge_result_preview(output)
    elif "pipeline" in command:
        show_pipeline_result_preview(output)
    else:
        messagebox.showinfo("处理完成", "操作已成功完成！\n请在data/output目录查看结果。")

def show_ocr_result_preview(output):
    """显示OCR处理结果预览"""
    # 提取处理的文件数量
    files_match = re.search(r'找到 (\d+) 个图片文件，其中 (\d+) 个未处理', output)
    processed_match = re.search(r'所有图片处理完成, 总计: (\d+), 成功: (\d+)', output)
    
    if processed_match:
        total = int(processed_match.group(1))
        success = int(processed_match.group(2))
        
        # 创建结果预览对话框
        preview = tk.Toplevel()
        preview.title("OCR处理结果")
        preview.geometry("400x300")
        preview.resizable(False, False)
        
        # 居中显示
        center_window(preview)
        
        # 添加内容
        tk.Label(preview, text="OCR处理完成", font=("Arial", 16, "bold")).pack(pady=10)
        
        result_frame = tk.Frame(preview)
        result_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(result_frame, text=f"总共处理: {total} 个文件", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        tk.Label(result_frame, text=f"成功处理: {success} 个文件", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        tk.Label(result_frame, text=f"失败数量: {total - success} 个文件", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        
        # 处理结果评估
        if success == total:
            result_text = "全部处理成功！"
            result_color = "#28a745"
        elif success > total * 0.8:
            result_text = "大部分处理成功。"
            result_color = "#ffc107"
        else:
            result_text = "处理失败较多，请检查日志。"
            result_color = "#dc3545"
            
        tk.Label(result_frame, text=result_text, font=("Arial", 12, "bold"), fg=result_color).pack(pady=10)
        
        # 添加按钮
        button_frame = tk.Frame(preview)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="查看输出文件", command=lambda: os.startfile(os.path.abspath("data/output"))).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="关闭", command=preview.destroy).pack(side=tk.LEFT, padx=10)
    else:
        messagebox.showinfo("OCR处理完成", "OCR处理已完成，请在data/output目录查看结果。")

def show_excel_result_preview(output):
    """显示Excel处理结果预览"""
    # 提取处理的Excel信息
    extract_match = re.search(r'提取到 (\d+) 个商品信息', output)
    file_match = re.search(r'采购单已保存到: (.+?)(?:\n|$)', output)
    
    if extract_match and file_match:
        products_count = int(extract_match.group(1))
        output_file = file_match.group(1)
        
        # 创建结果预览对话框
        preview = tk.Toplevel()
        preview.title("Excel处理结果")
        preview.geometry("450x320")
        preview.resizable(False, False)
        
        # 使弹窗居中显示
        center_window(preview)
        
        # 添加内容
        tk.Label(preview, text="Excel处理完成", font=("Arial", 16, "bold")).pack(pady=10)
        
        result_frame = tk.Frame(preview)
        result_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(result_frame, text=f"提取商品数量: {products_count} 个", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        tk.Label(result_frame, text=f"输出文件: {os.path.basename(output_file)}", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        
        # 处理成功提示
        tk.Label(result_frame, text="采购单已成功生成！", font=("Arial", 12, "bold"), fg="#28a745").pack(pady=10)
        
        # 文件信息框
        file_frame = tk.Frame(result_frame, relief=tk.GROOVE, borderwidth=1)
        file_frame.pack(fill=tk.X, padx=15, pady=5)
        
        tk.Label(file_frame, text="文件信息", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=5)
        
        # 获取文件大小和时间
        try:
            file_size = os.path.getsize(output_file)
            file_time = datetime.datetime.fromtimestamp(os.path.getmtime(output_file))
            
            size_text = f"{file_size / 1024:.1f} KB" if file_size < 1024*1024 else f"{file_size / (1024*1024):.1f} MB"
            
            tk.Label(file_frame, text=f"文件大小: {size_text}", font=("Arial", 10)).pack(anchor=tk.W, padx=10, pady=2)
            tk.Label(file_frame, text=f"创建时间: {file_time.strftime('%Y-%m-%d %H:%M:%S')}", font=("Arial", 10)).pack(anchor=tk.W, padx=10, pady=2)
        except:
            tk.Label(file_frame, text="无法获取文件信息", font=("Arial", 10)).pack(anchor=tk.W, padx=10, pady=2)
        
        # 添加按钮
        button_frame = tk.Frame(preview)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="打开文件", command=lambda: os.startfile(output_file)).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="打开所在文件夹", command=lambda: os.startfile(os.path.dirname(output_file))).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="关闭", command=preview.destroy).pack(side=tk.LEFT, padx=5)
    else:
        messagebox.showinfo("Excel处理完成", "Excel处理已完成，请在data/output目录查看结果。")

def show_merge_result_preview(output):
    """显示合并结果预览"""
    # 提取合并信息
    merged_match = re.search(r'合并了 (\d+) 个采购单', output)
    product_match = re.search(r'共处理 (\d+) 个商品', output)
    output_match = re.search(r'已保存到: (.+?)(?:\n|$)', output)
    
    if merged_match and output_match:
        merged_count = int(merged_match.group(1))
        product_count = int(product_match.group(1)) if product_match else 0
        output_file = output_match.group(1)
        
        # 创建结果预览对话框
        preview = tk.Toplevel()
        preview.title("采购单合并结果")
        preview.geometry("450x300")
        preview.resizable(False, False)
        
        # 设置主题
        apply_theme(preview)
        
        # 添加内容
        tk.Label(preview, text="采购单合并完成", font=("Arial", 16, "bold")).pack(pady=10)
        
        result_frame = tk.Frame(preview)
        result_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(result_frame, text=f"合并采购单数量: {merged_count} 个", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        tk.Label(result_frame, text=f"处理商品数量: {product_count} 个", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        tk.Label(result_frame, text=f"输出文件: {os.path.basename(output_file)}", font=("Arial", 12)).pack(anchor=tk.W, padx=20, pady=5)
        
        # 处理成功提示
        tk.Label(result_frame, text="采购单已成功合并！", font=("Arial", 12, "bold"), fg=THEMES[THEME_MODE]["success"]).pack(pady=10)
        
        # 添加按钮
        button_frame = tk.Frame(preview)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="打开文件", command=lambda: os.startfile(output_file)).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="打开所在文件夹", command=lambda: os.startfile(os.path.dirname(output_file))).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="关闭", command=preview.destroy).pack(side=tk.LEFT, padx=10)
    else:
        messagebox.showinfo("采购单合并完成", "采购单合并已完成，请在data/output目录查看结果。")

def show_pipeline_result_preview(output):
    """显示完整流程结果预览"""
    # 提取关键信息
    ocr_match = re.search(r'所有图片处理完成, 总计: (\d+), 成功: (\d+)', output)
    excel_match = re.search(r'提取到 (\d+) 个商品信息', output)
    output_file_match = re.search(r'采购单已保存到: (.+?)(?:\n|$)', output)
    
    # 创建结果预览对话框
    preview = tk.Toplevel()
    preview.title("完整流程处理结果")
    preview.geometry("500x400")
    preview.resizable(False, False)
    
    # 居中显示
    center_window(preview)
    
    # 添加内容
    tk.Label(preview, text="完整处理流程已完成", font=("Arial", 16, "bold")).pack(pady=10)
    
    # 添加处理结果提示（即使没有可合并文件也显示成功）
    no_files_match = re.search(r'未找到可合并的文件', output)
    if no_files_match:
        tk.Label(preview, text="未找到可合并文件，但其他步骤已成功执行", font=("Arial", 12)).pack(pady=0)
    
    result_frame = tk.Frame(preview)
    result_frame.pack(pady=10, fill=tk.BOTH, expand=True)
    
    # 创建多行结果区域
    result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, height=15, width=60)
    result_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
    result_text.configure(state=tk.NORMAL)
    
    # 填充结果文本
    result_text.insert(tk.END, "===== 流程执行结果 =====\n\n", "title")
    
    # OCR处理结果
    result_text.insert(tk.END, "步骤1: OCR识别\n", "step")
    if ocr_match:
        total = int(ocr_match.group(1))
        success = int(ocr_match.group(2))
        result_text.insert(tk.END, f"  处理图片: {total} 个\n", "info")
        result_text.insert(tk.END, f"  成功识别: {success} 个\n", "info")
        if success == total:
            result_text.insert(tk.END, "  结果: 全部识别成功\n", "success")
        else:
            result_text.insert(tk.END, f"  结果: 部分识别成功 ({success}/{total})\n", "warning")
    else:
        result_text.insert(tk.END, "  结果: 无OCR处理或处理信息不完整\n", "warning")
    
    # Excel处理结果
    result_text.insert(tk.END, "\n步骤2: Excel处理\n", "step")
    if excel_match:
        products = int(excel_match.group(1))
        result_text.insert(tk.END, f"  提取商品: {products} 个\n", "info")
        result_text.insert(tk.END, "  结果: 成功生成采购单\n", "success")
        if output_file_match:
            output_file = output_file_match.group(1)
            result_text.insert(tk.END, f"  输出文件: {os.path.basename(output_file)}\n", "info")
    else:
        result_text.insert(tk.END, "  结果: 无Excel处理或处理信息不完整\n", "warning")
    
    # 总体评估
    result_text.insert(tk.END, "\n===== 整体评估 =====\n", "title")
    
    has_errors = "错误" in output or "失败" in output
    
    no_files_match = re.search(r'未找到采购单文件', output)
    single_file_match = re.search(r'只有1个采购单文件', output)
    
    if no_files_match:
        result_text.insert(tk.END, "没有找到可合并的文件，但处理流程已成功完成。\n", "warning")
        result_text.insert(tk.END, "可以选择打开Excel文件或查看输出文件夹。\n", "info")
    elif single_file_match:
        result_text.insert(tk.END, "只有一个采购单文件，无需合并，处理流程已成功完成。\n", "warning")
        result_text.insert(tk.END, "可以选择打开生成的Excel文件。\n", "info")
    elif ocr_match and excel_match and not has_errors:
        result_text.insert(tk.END, "流程完整执行成功！\n", "success")
    elif ocr_match or excel_match:
        result_text.insert(tk.END, "流程部分执行成功，请检查日志获取详情。\n", "warning")
    else:
        result_text.insert(tk.END, "流程执行可能存在问题，请查看详细日志。\n", "error")
    
    # 设置标签样式
    result_text.tag_configure("title", font=("Arial", 12, "bold"))
    result_text.tag_configure("step", font=("Arial", 11, "bold"))
    result_text.tag_configure("info", font=("Arial", 10))
    result_text.tag_configure("success", font=("Arial", 10, "bold"), foreground="#28a745")
    result_text.tag_configure("warning", font=("Arial", 10, "bold"), foreground="#ffc107")
    result_text.tag_configure("error", font=("Arial", 10, "bold"), foreground="#dc3545")
    
    result_text.configure(state=tk.DISABLED)
    
    # 添加按钮
    button_frame = tk.Frame(preview)
    button_frame.pack(pady=10)
    
    if output_file_match:
        output_file = output_file_match.group(1)
        tk.Button(button_frame, text="打开Excel文件", command=lambda: os.startfile(output_file)).pack(side=tk.LEFT, padx=10)
    else:
        # 如果没有找到合并后的文件，但Excel处理成功，提供打开最新Excel文件的选项
        if excel_match or no_files_match or single_file_match:
            # 找到输出目录中最新的采购单Excel文件
            output_dir = os.path.abspath("data/output")
            excel_files = [f for f in os.listdir(output_dir) if f.startswith('采购单_') and (f.endswith('.xls') or f.endswith('.xlsx'))]
            if excel_files:
                # 按修改时间排序，获取最新的文件
                excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(output_dir, x)), reverse=True)
                latest_file = os.path.join(output_dir, excel_files[0])
                tk.Button(button_frame, text="打开最新Excel文件", 
                         command=lambda: os.startfile(latest_file)).pack(side=tk.LEFT, padx=10)
    
    tk.Button(button_frame, text="查看输出文件夹", command=lambda: os.startfile(os.path.abspath("data/output"))).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="关闭", command=preview.destroy).pack(side=tk.LEFT, padx=10)

def apply_theme(widget, theme_mode=None):
    """应用主题到小部件"""
    global THEME_MODE
    
    if theme_mode is None:
        theme_mode = THEME_MODE
    
    theme = THEMES[theme_mode]
    
    try:
        widget.configure(bg=theme["bg"], fg=theme["fg"])
    except:
        pass
        
    # 递归应用到所有子部件
    for child in widget.winfo_children():
        if isinstance(child, tk.Button) and not isinstance(child, ttk.Button):
            child.configure(bg=theme["button_bg"], fg=theme["button_fg"])
        elif isinstance(child, scrolledtext.ScrolledText):
            child.configure(bg=theme["log_bg"], fg=theme["log_fg"])
        else:
            try:
                child.configure(bg=theme["bg"], fg=theme["fg"])
            except:
                pass
                
        # 递归处理子部件的子部件
        apply_theme(child, theme_mode)

def toggle_theme(root, log_widget, status_bar=None):
    """切换主题模式"""
    global THEME_MODE
    
    # 切换主题模式
    THEME_MODE = "dark" if THEME_MODE == "light" else "light"
    
    # 应用主题到整个界面
    apply_theme(root)
    
    # 配置日志样式
    log_widget.configure(bg=THEMES[THEME_MODE]["log_bg"], fg=THEMES[THEME_MODE]["log_fg"])
    
    # 设置状态栏
    if status_bar:
        apply_theme(status_bar)
        
    # 保存主题设置
    try:
        with open("data/user_settings.json", "w") as f:
            json.dump({"theme": THEME_MODE}, f)
    except:
        pass
    
    return THEME_MODE

def ensure_directories():
    """确保必要的目录结构存在"""
    directories = ["data/input", "data/output", "data/temp", "logs"]
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
            print(f"创建目录: {directory}")

class LogRedirector:
    """日志重定向器，用于捕获命令输出并显示到界面"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""
        self.terminal = sys.__stdout__  # 保存原始的stdout引用
        
    def write(self, string):
        self.buffer += string
        # 同时输出到终端
        self.terminal.write(string)
        # 在UI线程中更新文本控件
        self.text_widget.after(0, self.update_text_widget)
        
    def update_text_widget(self):
        self.text_widget.configure(state=tk.NORMAL)
        
        # 根据内容使用不同的标签
        if self.buffer.strip():
            # 检测不同类型的消息并应用相应样式
            if any(marker in self.buffer.lower() for marker in ["错误", "error", "失败", "异常", "exception"]):
                self.text_widget.insert(tk.END, self.buffer, "error")
            elif any(marker in self.buffer.lower() for marker in ["警告", "warning"]):
                self.text_widget.insert(tk.END, self.buffer, "warning")
            elif any(marker in self.buffer.lower() for marker in ["成功", "success", "完成", "成功处理"]):
                self.text_widget.insert(tk.END, self.buffer, "success")
            elif any(marker in self.buffer.lower() for marker in ["info", "信息", "开始", "处理中"]):
                self.text_widget.insert(tk.END, self.buffer, "info")
            else:
                self.text_widget.insert(tk.END, self.buffer, "normal")
        else:
            self.text_widget.insert(tk.END, self.buffer)
            
        # 自动滚动到底部
        self.text_widget.see(tk.END)
        self.text_widget.configure(state=tk.DISABLED)
        self.buffer = ""
        
    def flush(self):
        self.terminal.flush()  # 确保终端也被刷新

def create_collapsible_frame(parent, title, initial_state=True):
    """创建可折叠的面板"""
    frame = tk.Frame(parent)
    frame.pack(fill=tk.X, pady=5)
    
    # 标题栏
    title_frame = tk.Frame(frame)
    title_frame.pack(fill=tk.X)
    
    # 折叠指示器
    state_var = tk.BooleanVar(value=initial_state)
    indicator = "▼" if initial_state else "►"
    state_label = tk.Label(title_frame, text=indicator, font=("Arial", 10, "bold"))
    state_label.pack(side=tk.LEFT, padx=5)
    
    # 标题
    title_label = tk.Label(title_frame, text=title, font=("Arial", 11, "bold"))
    title_label.pack(side=tk.LEFT, padx=5)
    
    # 内容区域
    content_frame = tk.Frame(frame)
    if initial_state:
        content_frame.pack(fill=tk.X, padx=20, pady=5)
    
    # 点击事件处理函数
    def toggle_collapse(event=None):
        current_state = state_var.get()
        new_state = not current_state
        state_var.set(new_state)
        
        # 更新指示器
        state_label.config(text="▼" if new_state else "►")
        
        # 显示或隐藏内容
        if new_state:
            content_frame.pack(fill=tk.X, padx=20, pady=5)
        else:
            content_frame.pack_forget()
    
    # 绑定点击事件
    title_frame.bind("<Button-1>", toggle_collapse)
    state_label.bind("<Button-1>", toggle_collapse)
    title_label.bind("<Button-1>", toggle_collapse)
    
    return content_frame, state_var

def main():
    """主函数"""
    # 确保必要的目录结构存在并转移旧目录内容
    ensure_directories()
    
    # 创建窗口
    root = tk.Tk()
    root.title("益选-OCR订单处理系统 v1.0")
    root.geometry("1200x650")  # 增加窗口高度以容纳更多元素
    
    # 创建主区域分割
    main_pane = tk.PanedWindow(root, orient=tk.HORIZONTAL)
    main_pane.pack(fill=tk.BOTH, expand=1, padx=5, pady=5)
    
    # 左侧操作区域
    left_frame = tk.Frame(main_pane, width=300)
    main_pane.add(left_frame)
    
    # 标题
    title_frame = tk.Frame(left_frame)
    title_frame.pack(fill=tk.X, pady=10)
    
    # 主标题
    tk.Label(title_frame, text="益选-OCR订单处理系统", font=("Arial", 16, "bold")).pack(side=tk.LEFT, padx=10)
    
    # 添加作者信息
    author_frame = tk.Frame(left_frame)
    author_frame.pack(fill=tk.X, pady=0)
    tk.Label(author_frame, text="作者：欢欢欢", font=("Arial", 10)).pack(side=tk.LEFT, padx=15)
    
    # 创建日志显示区域
    log_frame = tk.Frame(main_pane)
    main_pane.add(log_frame)
    
    # 日志标题
    tk.Label(log_frame, text="处理日志", font=("Arial", 12, "bold")).pack(pady=5)
    
    # 日志文本区域
    log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=30, width=60)
    log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    log_text.configure(state=tk.DISABLED)  # 设置为只读
    
    # 为日志文本添加标签样式
    log_text.tag_configure("normal", foreground="#000000")
    log_text.tag_configure("command", foreground="#17a2b8", font=("Arial", 10, "bold"))
    log_text.tag_configure("time", foreground="#17a2b8", font=("Arial", 9))
    log_text.tag_configure("separator", foreground="#cccccc")
    log_text.tag_configure("error", foreground="#dc3545")
    log_text.tag_configure("warning", foreground="#ffc107")
    log_text.tag_configure("success", foreground="#28a745")
    log_text.tag_configure("info", foreground="#17a2b8")
    
    # 创建状态栏
    status_bar = StatusBar(root)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # 日志初始内容
    add_to_log(log_text, "益选-OCR订单处理系统启动器 v1.0\n", "command")
    add_to_log(log_text, f"当前工作目录: {os.getcwd()}\n", "info")
    add_to_log(log_text, "系统已准备就绪，请选择要执行的操作。\n", "normal")
    
    # 创建按钮区域（使用两列布局）
    button_area = tk.Frame(left_frame)
    button_area.pack(fill=tk.BOTH, expand=True, pady=10)
    
    # 按钮尺寸和间距
    button_width = 15
    button_height = 2
    button_padx = 5
    button_pady = 5
    
    # 第一行
    row1 = tk.Frame(button_area)
    row1.pack(fill=tk.X, pady=button_pady)
    
    # 处理Excel文件
    tk.Button(
        row1, 
        text="处理Excel文件", 
        width=button_width,
        height=button_height,
        command=lambda: process_excel_file(log_text, status_bar)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # OCR批量识别
    tk.Button(
        row1, 
        text="OCR批量识别", 
        width=button_width,
        height=button_height,
        command=lambda: run_command_with_logging(["python", "run.py", "ocr", "--batch"], log_text, status_bar)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 第二行
    row2 = tk.Frame(button_area)
    row2.pack(fill=tk.X, pady=button_pady)
    
    # 完整处理流程
    tk.Button(
        row2, 
        text="完整处理流程", 
        width=button_width,
        height=button_height,
        command=lambda: run_command_with_logging(["python", "run.py", "pipeline"], log_text, status_bar)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 处理单个图片
    tk.Button(
        row2, 
        text="处理单个图片", 
        width=button_width,
        height=button_height,
        command=lambda: process_single_image(log_text, status_bar)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 第三行
    row3 = tk.Frame(button_area)
    row3.pack(fill=tk.X, pady=button_pady)
    
    # 合并采购单按钮
    tk.Button(
        row3, 
        text="合并采购单", 
        width=button_width,
        height=button_height,
        command=lambda: run_command_with_logging(["python", "run.py", "merge"], log_text, status_bar)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 整理项目文件
    tk.Button(
        row3, 
        text="整理项目文件", 
        width=button_width,
        height=button_height,
        command=lambda: organize_project_files(log_text)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 第四行
    row4 = tk.Frame(button_area)
    row4.pack(fill=tk.X, pady=button_pady)
    
    # 清除处理缓存按钮
    tk.Button(
        row4, 
        text="清除处理缓存", 
        width=button_width,
        height=button_height,
        command=lambda: clean_cache(log_text)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 清理文件按钮
    tk.Button(
        row4, 
        text="清理文件", 
        width=button_width,
        height=button_height,
        command=lambda: clean_data_files(log_text)
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 第五行
    row5 = tk.Frame(button_area)
    row5.pack(fill=tk.X, pady=button_pady)
    
    # 打开输入目录
    tk.Button(
        row5, 
        text="打开输入目录", 
        width=button_width,
        height=button_height,
        command=lambda: os.startfile(os.path.abspath("data/input"))
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 打开输出目录
    tk.Button(
        row5, 
        text="打开输出目录", 
        width=button_width,
        height=button_height,
        command=lambda: os.startfile(os.path.abspath("data/output"))
    ).pack(side=tk.LEFT, padx=button_padx)
    
    # 底部说明
    tk.Label(left_frame, text="© 2025 益选-OCR订单处理系统 v1.0 by 欢欢欢", font=("Arial", 9)).pack(side=tk.BOTTOM, pady=10)
    
    # 修改单个图片和Excel处理函数以使用状态栏
    def process_single_image_with_status(log_widget, status_bar):
        status_bar.set_status("选择图片中...")
        file_path = select_file(log_widget)
        if file_path:
            status_bar.set_status("开始处理图片...")
            run_command_with_logging(["python", "run.py", "ocr", "--input", file_path], log_widget, status_bar)
        else:
            status_bar.set_status("操作已取消")
            add_to_log(log_widget, "未选择文件，操作已取消\n", "warning")
    
    def process_excel_file_with_status(log_widget, status_bar):
        status_bar.set_status("选择Excel文件中...")
        file_path = select_excel_file(log_widget)
        if file_path:
            status_bar.set_status("开始处理Excel文件...")
            run_command_with_logging(["python", "run.py", "excel", "--input", file_path], log_widget, status_bar)
        else:
            status_bar.set_status("开始处理最新Excel文件...")
            add_to_log(log_widget, "未选择文件，尝试处理最新的Excel文件\n", "info")
            run_command_with_logging(["python", "run.py", "excel"], log_widget, status_bar)
    
    # 替换原始函数引用
    global process_single_image, process_excel_file
    process_single_image = process_single_image_with_status
    process_excel_file = process_excel_file_with_status
    
    # 启动主循环
    root.mainloop()

def add_to_log(log_widget, text, tag="normal"):
    """向日志窗口添加文本，支持样式标签"""
    log_widget.configure(state=tk.NORMAL)
    log_widget.insert(tk.END, text, tag)
    log_widget.see(tk.END)  # 自动滚动到底部
    log_widget.configure(state=tk.DISABLED)

def select_file(log_widget, file_types=[("所有文件", "*.*")], title="选择文件"):
    """通用文件选择对话框"""
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    if file_path:
        add_to_log(log_widget, f"已选择文件: {file_path}\n", "info")
    return file_path

def select_excel_file(log_widget):
    """选择Excel文件"""
    return select_file(
        log_widget, 
        [("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")], 
        "选择Excel文件"
    )

def clean_cache(log_widget):
    """清除处理缓存"""
    try:
        # 清除OCR缓存文件
        cache_files = [
            os.path.join("data", "processed_files.json"),
            os.path.join("data/output", "processed_files.json"),
            os.path.join("data/output", "merged_files.json")
        ]
        
        for cache_file in cache_files:
            if os.path.exists(cache_file):
                os.remove(cache_file)
                add_to_log(log_widget, f"已清除缓存文件: {cache_file}\n", "success")
        
        # 清除临时文件夹中所有文件
        temp_dir = os.path.join("data/temp")
        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, file)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        add_to_log(log_widget, f"已清除临时文件: {file_path}\n", "info")
                except Exception as e:
                    add_to_log(log_widget, f"清除文件时出错: {file_path}, 错误: {str(e)}\n", "error")
        
        # 清除日志文件中的active标记
        log_dir = "logs"
        if os.path.exists(log_dir):
            for file in os.listdir(log_dir):
                if file.endswith(".active"):
                    file_path = os.path.join(log_dir, file)
                    try:
                        os.remove(file_path)
                        add_to_log(log_widget, f"已清除活动日志标记: {file_path}\n", "info")
                    except Exception as e:
                        add_to_log(log_widget, f"清除文件时出错: {file_path}, 错误: {str(e)}\n", "error")
        
        # 重置全局状态
        global RUNNING_TASK
        RUNNING_TASK = None
        
        add_to_log(log_widget, "缓存清除完成，系统将重新处理所有文件\n", "success")
        messagebox.showinfo("缓存清除", "缓存已清除，系统将重新处理所有文件。")
    except Exception as e:
        add_to_log(log_widget, f"清除缓存时出错: {str(e)}\n", "error")
        messagebox.showerror("错误", f"清除缓存时出错: {str(e)}")

def organize_project_files(log_widget):
    """整理项目文件结构"""
    try:
        # 创建必要的目录
        directories = ["data/input", "data/output", "data/temp", "logs"]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory, exist_ok=True)
                add_to_log(log_widget, f"创建目录: {directory}\n", "info")
        
        # 移动日志文件到logs目录
        for file in os.listdir("."):
            if file.endswith(".log") and os.path.isfile(file):
                dest_path = os.path.join("logs", file)
                try:
                    shutil.move(file, dest_path)
                    add_to_log(log_widget, f"移动日志文件: {file} -> {dest_path}\n", "info")
                except Exception as e:
                    add_to_log(log_widget, f"移动文件时出错: {file}, 错误: {str(e)}\n", "error")
        
        # 移动配置文件到config目录
        if not os.path.exists("config"):
            os.makedirs("config", exist_ok=True)
        
        for file in os.listdir("."):
            if file.endswith(".ini") or file.endswith(".cfg") or file.endswith(".json"):
                if os.path.isfile(file) and file != "data/user_settings.json":
                    dest_path = os.path.join("config", file)
                    try:
                        shutil.move(file, dest_path)
                        add_to_log(log_widget, f"移动配置文件: {file} -> {dest_path}\n", "info")
                    except Exception as e:
                        add_to_log(log_widget, f"移动文件时出错: {file}, 错误: {str(e)}\n", "error")
        
        add_to_log(log_widget, "项目文件整理完成\n", "success")
    except Exception as e:
        add_to_log(log_widget, f"整理项目文件时出错: {str(e)}\n", "error")

def clean_data_files(log_widget):
    """清理数据文件"""
    try:
        # 确认清理
        if not messagebox.askyesno("确认清理", "确定要清理所有数据文件吗？这将删除所有输入和输出数据。"):
            add_to_log(log_widget, "操作已取消\n", "info")
            return
        
        # 清理输入目录
        input_dir = "data/input"
        files_cleaned = 0
        for file in os.listdir(input_dir):
            file_path = os.path.join(input_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                files_cleaned += 1
        
        # 清理输出目录
        output_dir = "data/output"
        for file in os.listdir(output_dir):
            file_path = os.path.join(output_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                files_cleaned += 1
                
        # 清理临时目录
        temp_dir = "data/temp"
        for file in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                files_cleaned += 1
        
        add_to_log(log_widget, f"已清理 {files_cleaned} 个数据文件\n", "success")
    except Exception as e:
        add_to_log(log_widget, f"清理数据文件时出错: {str(e)}\n", "error")

def center_window(window):
    """使窗口居中显示"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

if __name__ == "__main__":
    main() 