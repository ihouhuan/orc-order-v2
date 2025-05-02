#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
OCR订单处理系统启动器
-----------------
提供简单的图形界面，方便用户选择功能
"""

import os
import sys
import time
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from threading import Thread
import datetime

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
        
    def write(self, string):
        self.buffer += string
        # 在UI线程中更新文本控件
        self.text_widget.after(0, self.update_text_widget)
        
    def update_text_widget(self):
        self.text_widget.configure(state=tk.NORMAL)
        self.text_widget.insert(tk.END, self.buffer)
        # 自动滚动到底部
        self.text_widget.see(tk.END)
        self.text_widget.configure(state=tk.DISABLED)
        self.buffer = ""
        
    def flush(self):
        pass

def run_command_with_logging(command, log_widget):
    """运行命令并将输出重定向到日志窗口"""
    def run_in_thread():
        # 记录命令开始执行的时间
        start_time = datetime.datetime.now()
        log_widget.configure(state=tk.NORMAL)
        log_widget.delete(1.0, tk.END)  # 清空之前的日志
        log_widget.insert(tk.END, f"执行命令: {' '.join(command)}\n")
        log_widget.insert(tk.END, f"开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_widget.insert(tk.END, "=" * 50 + "\n\n")
        log_widget.configure(state=tk.DISABLED)
        
        # 设置环境变量，强制OCR模块输出到data目录
        env = os.environ.copy()
        env["OCR_OUTPUT_DIR"] = os.path.abspath("data/output")
        env["OCR_INPUT_DIR"] = os.path.abspath("data/input")
        env["OCR_LOG_LEVEL"] = "DEBUG"  # 设置更详细的日志级别
        
        try:
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
            
            # 读取并显示输出
            for line in process.stdout:
                log_widget.after(0, lambda l=line: add_to_log(log_widget, l))
                
            # 等待进程结束
            process.wait()
            
            # 记录命令结束时间
            end_time = datetime.datetime.now()
            duration = end_time - start_time
            
            log_widget.after(0, lambda: add_to_log(
                log_widget, 
                f"\n{'=' * 50}\n执行完毕！返回码: {process.returncode}\n"
                f"结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"耗时: {duration.total_seconds():.2f} 秒\n"
            ))
            
            # 如果处理成功，显示成功信息
            if process.returncode == 0:
                log_widget.after(0, lambda: messagebox.showinfo("操作成功", "处理完成！\n请在data/output目录查看结果。"))
            else:
                log_widget.after(0, lambda: messagebox.showerror("操作失败", f"处理失败，返回码：{process.returncode}"))
                
        except Exception as e:
            log_widget.after(0, lambda: add_to_log(log_widget, f"\n执行出错: {str(e)}\n"))
            log_widget.after(0, lambda: messagebox.showerror("执行错误", f"执行命令时出错: {str(e)}"))
    
    # 在新线程中运行，避免UI阻塞
    Thread(target=run_in_thread).start()

def add_to_log(log_widget, text):
    """向日志窗口添加文本"""
    log_widget.configure(state=tk.NORMAL)
    log_widget.insert(tk.END, text)
    log_widget.see(tk.END)  # 自动滚动到底部
    log_widget.configure(state=tk.DISABLED)

def select_file(log_widget):
    """选择图片文件并复制到data/input目录"""
    # 确保目录存在
    ensure_directories()
    
    # 获取输入目录的绝对路径
    input_dir = os.path.abspath("data/input")
    
    file_path = filedialog.askopenfilename(
        title="选择要处理的图片文件",
        initialdir=input_dir,  # 默认打开data/input目录
        filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
    )
    
    if not file_path:
        return None
        
    # 记录选择文件的信息
    add_to_log(log_widget, f"已选择文件: {file_path}\n")
    
    # 计算目标路径，始终放在data/input中
    output_path = os.path.join("data/input", os.path.basename(file_path))
    abs_output_path = os.path.abspath(output_path)
    
    # 检查是否是同一个文件
    if os.path.normpath(os.path.abspath(file_path)) != os.path.normpath(abs_output_path):
        # 如果是不同的文件，则复制
        try:
            shutil.copy2(file_path, output_path)
            add_to_log(log_widget, f"已复制文件到处理目录: {output_path}\n")
        except Exception as e:
            add_to_log(log_widget, f"复制文件失败: {e}\n")
            messagebox.showerror("错误", f"复制文件失败: {e}")
            return None
    
    # 返回绝对路径，确保命令行处理正确
    return abs_output_path

def select_excel_file(log_widget):
    """选择Excel文件并复制到data/output目录"""
    # 确保目录存在
    ensure_directories()
    
    # 获取输出目录的绝对路径
    output_dir = os.path.abspath("data/output")
    
    file_path = filedialog.askopenfilename(
        title="选择要处理的Excel文件",
        initialdir=output_dir,  # 默认打开data/output目录
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return None
        
    # 记录选择文件的信息
    add_to_log(log_widget, f"已选择文件: {file_path}\n")
    
    # 计算目标路径，始终放在data/output中
    output_path = os.path.join("data/output", os.path.basename(file_path))
    abs_output_path = os.path.abspath(output_path)
    
    # 检查是否是同一个文件
    if os.path.normpath(os.path.abspath(file_path)) != os.path.normpath(abs_output_path):
        # 如果是不同的文件，则复制
        try:
            shutil.copy2(file_path, output_path)
            add_to_log(log_widget, f"已复制文件到处理目录: {output_path}\n")
        except Exception as e:
            add_to_log(log_widget, f"复制文件失败: {e}\n")
            messagebox.showerror("错误", f"复制文件失败: {e}")
            return None
    
    # 返回绝对路径，确保命令行处理正确
    return abs_output_path

def process_single_image(log_widget):
    """处理单个图片"""
    file_path = select_file(log_widget)
    if file_path:
        # 确保文件存在
        if os.path.exists(file_path):
            add_to_log(log_widget, f"正在处理图片: {os.path.basename(file_path)}\n")
            # 使用绝对路径，并指定直接输出到data/output
            run_command_with_logging(["python", "run.py", "ocr", "--input", file_path], log_widget)
        else:
            add_to_log(log_widget, f"文件不存在: {file_path}\n")
            messagebox.showerror("错误", f"文件不存在: {file_path}")
    else:
        add_to_log(log_widget, "未选择文件，操作已取消\n")

def process_excel_file(log_widget):
    """处理Excel文件"""
    file_path = select_excel_file(log_widget)
    if file_path:
        # 确保文件存在
        if os.path.exists(file_path):
            add_to_log(log_widget, f"正在处理Excel文件: {os.path.basename(file_path)}\n")
            # 使用绝对路径
            run_command_with_logging(["python", "run.py", "excel", "--input", file_path], log_widget)
        else:
            add_to_log(log_widget, f"文件不存在: {file_path}\n")
            messagebox.showerror("错误", f"文件不存在: {file_path}")
    else:
        # 如果未选择文件，尝试处理最新的Excel
        add_to_log(log_widget, "未选择文件，尝试处理最新的Excel文件\n")
        run_command_with_logging(["python", "run.py", "excel"], log_widget)

def organize_project_files(log_widget):
    """整理项目中的文件到正确位置"""
    # 确保目录存在
    ensure_directories()
    
    add_to_log(log_widget, "开始整理项目文件...\n")
    
    # 转移根目录文件
    files_moved = 0
    
    # 处理日志文件
    log_files = [f for f in os.listdir('.') if f.endswith('.log')]
    for log_file in log_files:
        try:
            src_path = os.path.join('.', log_file)
            dst_path = os.path.join('logs', log_file)
            if not os.path.exists(dst_path) or os.path.getmtime(src_path) > os.path.getmtime(dst_path):
                shutil.copy2(src_path, dst_path)
                add_to_log(log_widget, f"已移动日志文件: {src_path} -> {dst_path}\n")
                files_moved += 1
        except Exception as e:
            add_to_log(log_widget, f"移动日志文件出错: {e}\n")
    
    # 处理JSON文件
    json_files = [f for f in os.listdir('.') if f.endswith('.json')]
    for json_file in json_files:
        try:
            src_path = os.path.join('.', json_file)
            dst_path = os.path.join('data', json_file)
            if not os.path.exists(dst_path) or os.path.getmtime(src_path) > os.path.getmtime(dst_path):
                shutil.copy2(src_path, dst_path)
                add_to_log(log_widget, f"已移动记录文件: {src_path} -> {dst_path}\n")
                files_moved += 1
        except Exception as e:
            add_to_log(log_widget, f"移动记录文件出错: {e}\n")
    
    # 处理input和output目录
    for old_dir, new_dir in {"input": "data/input", "output": "data/output"}.items():
        if os.path.exists(old_dir) and os.path.isdir(old_dir):
            for file in os.listdir(old_dir):
                src_path = os.path.join(old_dir, file)
                dst_path = os.path.join(new_dir, file)
                try:
                    if os.path.isfile(src_path):
                        if not os.path.exists(dst_path) or os.path.getmtime(src_path) > os.path.getmtime(dst_path):
                            shutil.copy2(src_path, dst_path)
                            add_to_log(log_widget, f"已转移文件: {src_path} -> {dst_path}\n")
                            files_moved += 1
                except Exception as e:
                    add_to_log(log_widget, f"移动文件出错: {e}\n")
    
    # 显示结果
    if files_moved > 0:
        add_to_log(log_widget, f"整理完成，共整理 {files_moved} 个文件\n")
        messagebox.showinfo("整理完成", f"已整理 {files_moved} 个文件到正确位置。\n"
                           "原始文件保留在原位置，以确保数据安全。")
    else:
        add_to_log(log_widget, "没有需要整理的文件\n")
        messagebox.showinfo("整理完成", "没有需要整理的文件。")

def main():
    """主函数"""
    # 确保必要的目录结构存在并转移旧目录内容
    ensure_directories()
    
    # 创建窗口
    root = tk.Tk()
    root.title("OCR订单处理系统 v2.0")
    root.geometry("800x600")  # 增加窗口宽度以容纳日志
    
    # 创建主区域分割
    main_pane = tk.PanedWindow(root, orient=tk.HORIZONTAL)
    main_pane.pack(fill=tk.BOTH, expand=1, padx=5, pady=5)
    
    # 左侧操作区域
    left_frame = tk.Frame(main_pane, width=300)
    main_pane.add(left_frame)
    
    # 标题
    tk.Label(left_frame, text="OCR订单处理系统", font=("Arial", 16)).pack(pady=10)
    
    # 功能按钮区域
    buttons_frame = tk.Frame(left_frame)
    buttons_frame.pack(pady=10, fill=tk.Y)
    
    # 创建日志显示区域
    log_frame = tk.Frame(main_pane)
    main_pane.add(log_frame)
    
    # 日志标题
    tk.Label(log_frame, text="处理日志", font=("Arial", 12)).pack(pady=5)
    
    # 日志文本区域
    log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=30, width=60)
    log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    log_text.configure(state=tk.DISABLED)  # 设置为只读
    
    # 日志初始内容
    add_to_log(log_text, "OCR订单处理系统启动器 v2.0\n")
    add_to_log(log_text, f"当前工作目录: {os.getcwd()}\n")
    add_to_log(log_text, "系统已准备就绪，请选择要执行的操作。\n")
    
    # OCR识别按钮
    tk.Button(
        buttons_frame, 
        text="OCR图像识别 (批量)", 
        width=20,
        height=2,
        command=lambda: run_command_with_logging(["python", "run.py", "ocr", "--batch"], log_text)
    ).pack(pady=5)
    
    # 单个图片处理
    tk.Button(
        buttons_frame, 
        text="处理单个图片", 
        width=20,
        height=2,
        command=lambda: process_single_image(log_text)
    ).pack(pady=5)
    
    # Excel处理按钮
    tk.Button(
        buttons_frame, 
        text="处理Excel文件", 
        width=20,
        height=2,
        command=lambda: process_excel_file(log_text)
    ).pack(pady=5)
    
    # 订单合并按钮
    tk.Button(
        buttons_frame, 
        text="合并采购单", 
        width=20,
        height=2,
        command=lambda: run_command_with_logging(["python", "run.py", "merge"], log_text)
    ).pack(pady=5)
    
    # 完整流程按钮
    tk.Button(
        buttons_frame, 
        text="完整处理流程", 
        width=20,
        height=2,
        command=lambda: run_command_with_logging(["python", "run.py", "pipeline"], log_text)
    ).pack(pady=5)
    
    # 整理文件按钮
    tk.Button(
        buttons_frame, 
        text="整理项目文件", 
        width=20,
        height=2,
        command=lambda: organize_project_files(log_text)
    ).pack(pady=5)
    
    # 打开输入目录
    tk.Button(
        buttons_frame, 
        text="打开输入目录", 
        width=20,
        command=lambda: os.startfile(os.path.abspath("data/input"))
    ).pack(pady=5)
    
    # 打开输出目录
    tk.Button(
        buttons_frame, 
        text="打开输出目录", 
        width=20,
        command=lambda: os.startfile(os.path.abspath("data/output"))
    ).pack(pady=5)
    
    # 清空日志按钮
    tk.Button(
        buttons_frame,
        text="清空日志",
        width=20,
        command=lambda: log_text.delete(1.0, tk.END)
    ).pack(pady=5)
    
    # 底部说明
    tk.Label(left_frame, text="© 2025 OCR订单处理系统", font=("Arial", 10)).pack(side=tk.BOTTOM, pady=10)
    
    # 启动主循环
    root.mainloop()

if __name__ == "__main__":
    main() 