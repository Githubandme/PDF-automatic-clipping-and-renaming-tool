
# enhanced_barcode_processor_ui.py
# 使用CustomTkinter构建PDF处理工具的图形界面

import os
import sys
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from enhanced_barcode_processor import EnhancedPDFProcessor

# 设置CustomTkinter的外观模式和颜色主题
ctk.set_appearance_mode("System")  # 系统模式，自动适应系统主题
ctk.set_default_color_theme("blue")  # 蓝色主题

class ScrollableTextFrame(ctk.CTkFrame):
    """可滚动的文本框框架"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # 创建文本框和滚动条
        self.text = ctk.CTkTextbox(self, wrap="word", height=200)
        self.text.pack(fill="both", expand=True, padx=10, pady=10)

    def insert_text(self, text):
        """插入文本并自动滚动到底部"""
        self.text.configure(state="normal")
        self.text.insert("end", text + "\n")
        self.text.configure(state="disabled")
        self.text.see("end")

    def clear(self):
        """清空文本框"""
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

class LogRedirector:
    """重定向日志输出到文本框"""

    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        if "\n" in self.buffer:
            lines = self.buffer.split("\n")
            for line in lines[:-1]:
                if line:  # 跳过空行
                    self.text_widget.insert_text(line)
            self.buffer = lines[-1]

    def flush(self):
        if self.buffer:
            self.text_widget.insert_text(self.buffer)
            self.buffer = ""

class EnhancedPDFProcessorUI(ctk.CTk):
    """增强版PDF处理工具的图形界面"""

    def __init__(self):
        super().__init__()

        # 配置窗口
        self.title("增强版PDF处理工具")
        self.geometry("800x700")
        self.minsize(800, 700)

        # 初始化处理器
        self.processor = EnhancedPDFProcessor()

        # 创建UI组件
        self.create_widgets()

        # 初始化处理状态
        self.processing = False
        self.current_log_file = None

    def create_widgets(self):
        """创建界面组件"""
        # 主框架 - 使用网格布局
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 设置主框架的网格
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=0)

        # ===== 输入设置部分 =====
        self.input_frame = ctk.CTkFrame(self.main_frame)
        self.input_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 输入文件夹
        self.input_label = ctk.CTkLabel(self.input_frame, text="输入文件夹:")
        self.input_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.input_entry = ctk.CTkEntry(self.input_frame, width=400)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        self.input_button = ctk.CTkButton(self.input_frame, text="浏览...", command=self.browse_input_folder)
        self.input_button.grid(row=0, column=2, padx=10, pady=10)

        # 输出文件夹
        self.output_label = ctk.CTkLabel(self.input_frame, text="输出文件夹:")
        self.output_label.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        self.output_entry = ctk.CTkEntry(self.input_frame, width=400)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        self.output_button = ctk.CTkButton(self.input_frame, text="浏览...", command=self.browse_output_folder)
        self.output_button.grid(row=1, column=2, padx=10, pady=10)

        # 配置输入框架的网格
        self.input_frame.grid_columnconfigure(1, weight=1)

        # ===== 参数设置部分 =====
        self.params_frame = ctk.CTkFrame(self.main_frame)
        self.params_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 边框宽度
        self.border_label = ctk.CTkLabel(self.params_frame, text="裁剪边框宽度:")
        self.border_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.border_slider = ctk.CTkSlider(self.params_frame, from_=0, to=20, number_of_steps=20)
        self.border_slider.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        self.border_slider.set(5)  # 默认值

        self.border_value = ctk.CTkLabel(self.params_frame, text="5")
        self.border_value.grid(row=0, column=2, padx=5, pady=10)

        # 更新边框值显示
        self.border_slider.configure(command=self.update_border_value)

        # DPI设置
        self.dpi_label = ctk.CTkLabel(self.params_frame, text="条码识别DPI:")
        self.dpi_label.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        self.dpi_slider = ctk.CTkSlider(self.params_frame, from_=150, to=600, number_of_steps=45)
        self.dpi_slider.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        self.dpi_slider.set(300)  # 默认值

        self.dpi_value = ctk.CTkLabel(self.params_frame, text="300")
        self.dpi_value.grid(row=1, column=2, padx=5, pady=10)

        # 更新DPI值显示
        self.dpi_slider.configure(command=self.update_dpi_value)

        # 处理步骤选择
        self.steps_label = ctk.CTkLabel(self.params_frame, text="处理步骤:")
        self.steps_label.grid(row=2, column=0, sticky="w", padx=10, pady=10)

        self.steps_frame = ctk.CTkFrame(self.params_frame, fg_color="transparent")
        self.steps_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=10, pady=10)

        self.step1_var = ctk.BooleanVar(value=True)
        self.step1_check = ctk.CTkCheckBox(self.steps_frame, text="分页", variable=self.step1_var)
        self.step1_check.grid(row=0, column=0, padx=5, pady=5)

        self.step2_var = ctk.BooleanVar(value=True)
        self.step2_check = ctk.CTkCheckBox(self.steps_frame, text="空白裁剪", variable=self.step2_var)
        self.step2_check.grid(row=0, column=1, padx=5, pady=5)

        self.step3_var = ctk.BooleanVar(value=True)
        self.step3_check = ctk.CTkCheckBox(self.steps_frame, text="条码识别重命名", variable=self.step3_var)
        self.step3_check.grid(row=0, column=2, padx=5, pady=5)

        # 配置参数框架的网格
        self.params_frame.grid_columnconfigure(1, weight=1)

        # ===== 进度部分 =====
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        self.progress_label = ctk.CTkLabel(self.progress_frame, text="进度:")
        self.progress_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        self.progress_bar.set(0)

        # 配置进度框架的网格
        self.progress_frame.grid_columnconfigure(1, weight=1)

        # ===== 日志部分 =====
        self.log_label = ctk.CTkLabel(self.main_frame, text="处理日志:")
        self.log_label.grid(row=3, column=0, sticky="w", padx=10, pady=(10, 0))

        self.log_frame = ScrollableTextFrame(self.main_frame)
        self.log_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        # 重定向标准输出到日志框
        self.log_redirector = LogRedirector(self.log_frame)

        # ===== 操作按钮 =====
        self.button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.button_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        self.start_button = ctk.CTkButton(
            self.button_frame, 
            text="开始处理", 
            command=self.start_processing,
            fg_color="#28a745",  # 绿色
            hover_color="#218838"
        )
        self.start_button.pack(side="left", padx=10)

        self.stop_button = ctk.CTkButton(
            self.button_frame, 
            text="停止", 
            command=self.stop_processing,
            fg_color="#dc3545",  # 红色
            hover_color="#c82333",
            state="disabled"
        )
        self.stop_button.pack(side="left", padx=10)

        self.open_output_button = ctk.CTkButton(
            self.button_frame, 
            text="打开输出文件夹", 
            command=self.open_output_folder
        )
        self.open_output_button.pack(side="left", padx=10)

        self.clear_log_button = ctk.CTkButton(
            self.button_frame, 
            text="清除日志", 
            command=self.clear_log
        )
        self.clear_log_button.pack(side="right", padx=10)

        # 配置主框架的网格
        self.main_frame.grid_rowconfigure(4, weight=1)  # 让日志框架占据多余空间

    def update_border_value(self, value):
        """更新边框宽度值显示"""
        int_value = int(value)
        self.border_value.configure(text=str(int_value))

    def update_dpi_value(self, value):
        """更新DPI值显示"""
        int_value = int(value)
        self.dpi_value.configure(text=str(int_value))

    def browse_input_folder(self):
        """浏览选择输入文件夹"""
        folder = filedialog.askdirectory(title="选择输入文件夹")
        if folder:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, folder)

    def browse_output_folder(self):
        """浏览选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)

    def open_output_folder(self):
        """打开输出文件夹"""
        output_folder = self.output_entry.get()
        if not output_folder or not os.path.exists(output_folder):
            messagebox.showerror("错误", "输出文件夹不存在！")
            return

        # 根据操作系统打开文件夹
        if sys.platform == 'win32':
            os.startfile(output_folder)
        elif sys.platform == 'darwin':  # macOS
            os.system(f'open "{output_folder}"')
        else:  # Linux
            os.system(f'xdg-open "{output_folder}"')

    def clear_log(self):
        """清除日志框内容"""
        self.log_frame.clear()

    def validate_inputs(self):
        """验证输入是否有效"""
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()

        if not input_folder:
            messagebox.showerror("错误", "请选择输入文件夹！")
            return False

        if not os.path.exists(input_folder):
            messagebox.showerror("错误", "输入文件夹不存在！")
            return False

        if not output_folder:
            messagebox.showerror("错误", "请选择输出文件夹！")
            return False

        # 检查输入文件夹中是否有PDF文件
        pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
        if not pdf_files:
            messagebox.showerror("错误", "输入文件夹中没有PDF文件！")
            return False

        return True

    def start_processing(self):
        """开始处理文件"""
        if self.processing:
            return

        if not self.validate_inputs():
            return

        # 设置处理状态
        self.processing = True
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")

        # 获取参数
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()
        border_width = int(self.border_slider.get())
        dpi = int(self.dpi_slider.get())

        # 创建输出文件夹
        os.makedirs(output_folder, exist_ok=True)

        # 确定要执行的步骤
        steps = []
        if self.step1_var.get():
            steps.append("split")
        if self.step2_var.get():
            steps.append("crop")
        if self.step3_var.get():
            steps.append("barcode")

        if not steps:
            messagebox.showerror("错误", "请至少选择一个处理步骤！")
            self.processing = False
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            return

        # 创建日志文件
        self.current_log_file = os.path.join(output_folder, f"处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

        # 清空日志
        self.clear_log()

        # 重定向标准输出
        old_stdout = sys.stdout
        sys.stdout = self.log_redirector

        # 在单独的线程中处理
        self.processing_thread = threading.Thread(
            target=self.process_files_thread,
            args=(input_folder, output_folder, border_width, dpi, steps)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()

        # 启动进度监控
        self.after(100, self.check_progress)

    def process_files_thread(self, input_folder, output_folder, border_width, dpi, steps):
        """在线程中处理文件"""
        try:
            # 设置处理器DPI
            self.processor = EnhancedPDFProcessor(dpi=dpi)

            # 创建必要的文件夹
            single_page_folder = os.path.join(output_folder, "单页PDF文件夹")
            cropped_folder = os.path.join(output_folder, "空白裁剪文件夹")
            renamed_folder = os.path.join(output_folder, "重命名文件夹")

            if "split" in steps:
                os.makedirs(single_page_folder, exist_ok=True)
            if "crop" in steps:
                os.makedirs(cropped_folder, exist_ok=True)
            if "barcode" in steps:
                os.makedirs(renamed_folder, exist_ok=True)

            # 处理步骤1: 分页
            if "split" in steps:
                print("===== 步骤1: PDF分割 =====")
                pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
                print(f"找到 {len(pdf_files)} 个PDF文件")

                for i, pdf_file in enumerate(pdf_files):
                    if not self.processing:
                        break

                    print(f"[{i+1}/{len(pdf_files)}] 正在分割: {pdf_file}")
                    input_path = os.path.join(input_folder, pdf_file)

                    try:
                        # 分割PDF为单页
                        split_pages = self.processor.split_pdf_to_single_pages(input_path, single_page_folder)
                        print(f"  成功分割为 {len(split_pages)} 页")
                    except Exception as e:
                        print(f"  分割失败: {str(e)}")

            # 处理步骤2: 空白裁剪
            if "crop" in steps:
                print("===== 步骤2: 空白裁剪 =====")

                # 确定要处理的文件夹
                source_folder = single_page_folder if "split" in steps else input_folder

                if not os.path.exists(source_folder):
                    print(f"错误: 源文件夹不存在 - {source_folder}")
                else:
                    single_page_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
                    print(f"找到 {len(single_page_files)} 个PDF文件")

                    for i, pdf_file in enumerate(single_page_files):
                        if not self.processing:
                            break

                        print(f"[{i+1}/{len(single_page_files)}] 正在裁剪: {pdf_file}")
                        input_path = os.path.join(source_folder, pdf_file)
                        output_path = os.path.join(cropped_folder, pdf_file)

                        try:
                            # 裁剪空白区域
                            self.processor.auto_crop_pdf(input_path, output_path, border_width)
                            print(f"  裁剪成功: {pdf_file}")
                        except Exception as e:
                            print(f"  裁剪失败: {str(e)}")

            # 处理步骤3: 条码识别重命名
            if "barcode" in steps:
                print("===== 步骤3: 条码识别重命名 =====")

                # 确定要处理的文件夹
                if "crop" in steps:
                    source_folder = cropped_folder
                elif "split" in steps:
                    source_folder = single_page_folder
                else:
                    source_folder = input_folder

                if not os.path.exists(source_folder):
                    print(f"错误: 源文件夹不存在 - {source_folder}")
                else:
                    source_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
                    print(f"找到 {len(source_files)} 个PDF文件")

                    for i, pdf_file in enumerate(source_files):
                        if not self.processing:
                            break

                        print(f"[{i+1}/{len(source_files)}] 正在识别条码: {pdf_file}")
                        input_path = os.path.join(source_folder, pdf_file)

                        try:
                            # 提取条码
                            barcode = self.processor.extract_barcode_from_pdf(input_path)

                            # 创建新文件名
                            if barcode and barcode != "未找到条码" and barcode != "条码提取错误":
                                print(f"  识别到条码: {barcode}")
                                new_filename = f"{barcode}.pdf"
                            else:
                                print(f"  未识别到条码, 使用原文件名")
                                new_filename = pdf_file

                            # 确保文件名不重复
                            output_path = os.path.join(renamed_folder, new_filename)
                            counter = 1
                            while os.path.exists(output_path):
                                name_parts = os.path.splitext(new_filename)
                                new_filename = f"{name_parts[0]}_{counter}{name_parts[1]}"
                                output_path = os.path.join(renamed_folder, new_filename)
                                counter += 1

                            # 复制文件到重命名文件夹
                            import shutil
                            shutil.copy2(input_path, output_path)
                            print(f"  已重命名为: {new_filename}")
                        except Exception as e:
                            print(f"  识别或重命名失败: {str(e)}")

            print("===== 处理完成 =====")

        except Exception as e:
            print(f"处理过程中发生错误: {str(e)}")

        finally:
            # 恢复标准输出
            sys.stdout = sys.__stdout__

            # 处理完成
            self.processing = False

            # 更新UI状态(在主线程中)
            self.after(0, self.update_ui_after_processing)

    def update_ui_after_processing(self):
        """处理完成后更新UI状态"""
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.progress_bar.set(1)  # 设置进度条为100%

        # 显示完成消息
        messagebox.showinfo("完成", "PDF处理已完成！")

    def stop_processing(self):
        """停止处理"""
        if not self.processing:
            return

        if messagebox.askyesno("确认", "确定要停止处理吗？"):
            self.processing = False
            print("用户请求停止处理...")

    def check_progress(self):
        """检查处理进度"""
        if self.processing:
            # 更新进度条(模拟进度)
            current = self.progress_bar.get()
            if current < 0.95:  # 保留最后5%，等待实际完成
                self.progress_bar.set(current + 0.01)

            # 继续检查
            self.after(100, self.check_progress)
        else:
            # 处理已停止或完成
            self.progress_bar.set(1)

def main():
    app = EnhancedPDFProcessorUI()
    app.mainloop()


if __name__ == "__main__":
    main()
