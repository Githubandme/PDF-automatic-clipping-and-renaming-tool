import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import os
import subprocess
import tempfile
import shutil
import pandas as pd
import cv2
from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
from datetime import datetime
import re
import sys
import logging
import threading

# 定义日志函数
def log_message(message, level="info"):
    """记录日志消息到日志框"""
    global log_text
    if log_text:
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_text.configure(state='normal')
        log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        log_text.see(tk.END)  # 自动滚动到底部
        log_text.configure(state='disabled')

# 创建主窗口 - 改为标准tkinter样式
window = tk.Tk()
window.title("PDF 自动裁剪与重命名工具")
window.resizable(False, False)  # 固定窗口大小
window_width = 1200
window_height = 1000
window.geometry(f"{window_width}x{window_height}")

# 检查并设置程序图标
icon_path = os.path.join(os.path.dirname(__file__), "PDF裁剪扫码.ico")
if os.path.exists(icon_path):
    try:
        window.iconbitmap(icon_path)
    except Exception as e:
        log_message(f"设置图标失败: {str(e)}", "warning")

# ==================== 打包环境支持 ====================
def resource_path(relative_path):
    """获取打包后资源的绝对路径"""
    try:
        # PyInstaller创建的临时文件夹
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    # 添加对 libiconv.dll 的特殊处理
    if "libiconv2.dll" in relative_path and is_frozen:
        return os.path.join(base_path, "_internal", "pyzbar", relative_path)
    
    return os.path.join(base_path, relative_path)

# 检查是否是打包环境
is_frozen = getattr(sys, 'frozen', False)

# 启用 DPI 感知
if os.name == 'nt':  # 仅在 Windows 上启用 DPI 感知
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print("Failed to set DPI awareness")

# ==================== 全局变量 ==================== 
enable_rename_var = tk.BooleanVar(value=True)
enable_logging_var = tk.BooleanVar(value=True)
report_path = tk.StringVar()  # 不再设置初始值，改为输出文件夹改变时动态更新
poppler_path = tk.StringVar(value="poppler/bin")  # 修改为默认相对路径
log_text = None  # 用于日志文本框的全局引用
is_processing = False  # 添加处理状态标志

# 添加路径设置函数
def select_poppler_path():
    path = filedialog.askdirectory(title="选择Poppler路径", initialdir=poppler_path.get())
    if path:
        # 转换为相对路径
        rel_path = os.path.relpath(path, os.path.dirname(__file__))
        poppler_path.set(rel_path)
        log_message(f"设置Poppler路径: {rel_path}")

def check_dll_files():
    """简化后的依赖检查，只检查poppler和libiconv2.dll"""
    poppler = resource_path(poppler_path.get()) if poppler_path.get() else resource_path("poppler/bin")
    libiconv = resource_path("libiconv2.dll")
    
    if not os.path.exists(poppler):
        log_message(f"警告: Poppler路径不存在: {poppler}", "warning")
        return False
    
    if not os.path.exists(libiconv):
        log_message(f"警告: libiconv2.dll不存在: {libiconv}", "warning")
        return False
    
    log_message(f"Poppler路径检查通过: {poppler}")
    log_message(f"libiconv2.dll路径检查通过: {libiconv}")
    return True

def check_dependencies():
    """简化后的依赖检查"""
    poppler = resource_path(poppler_path.get()) if poppler_path.get() else resource_path("poppler/bin")
    libiconv = resource_path("libiconv2.dll")
    
    log_message(f"实际使用的Poppler路径: {poppler}")
    log_message(f"实际使用的libiconv2.dll路径: {libiconv}")
    
    # 检查文件是否存在
    poppler_exists = os.path.exists(poppler)
    libiconv_exists = os.path.exists(libiconv)
    log_message(f"Poppler {'存在' if poppler_exists else '不存在'}")
    log_message(f"libiconv2.dll {'存在' if libiconv_exists else '不存在'}")
    
    # 功能测试
    try:
        # 测试条码识别功能
        test_image = np.zeros((100, 100), dtype=np.uint8)
        decode(test_image)  # 尝试解码空白图像
        log_message("条码识别功能测试通过")
    except Exception as e:
        log_message(f"条码识别功能测试失败: {str(e)}", "error")
    
    # 更新状态栏
    status_label.config(text=f"Poppler路径: {poppler}\nlibiconv2.dll路径: {libiconv}")

# ==================== 功能函数 ====================
def select_pdf_files():
    """打开文件对话框，选择多个PDF文件."""
    file_paths = filedialog.askopenfilenames(title="选择 PDF 文件", filetypes=[("PDF files", "*.pdf")])
    if file_paths:
        for file_path in file_paths:
            # 避免重复添加文件
            if file_path not in input_files_listbox.get(0, tk.END):
                input_files_listbox.insert(tk.END, file_path)
        status_label.config(text=f"已选择 {len(file_paths)} 个文件")
        log_message(f"已选择 {len(file_paths)} 个PDF文件")

def select_output_folder():
    """打开文件夹选择对话框，选择输出文件夹."""
    folder_path = filedialog.askdirectory(title="选择输出文件夹")
    if folder_path:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)
        # 自动更新报告文件路径
        report_path.set(os.path.join(folder_path, "重命名报告.xlsx"))
        status_label.config(text="输出文件夹已选择：" + folder_path)
        log_message(f"设置输出文件夹: {folder_path}")

def select_report_path():
    """删除此函数，不再需要手动选择报告路径"""
    pass

def split_pdf_to_single_pages(input_pdf_path, output_folder):
    """将PDF拆分为单页文件"""
    os.makedirs(output_folder, exist_ok=True)
    pdf_document = fitz.open(input_pdf_path)
    file_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
    page_files = []
    
    for page_number in range(pdf_document.page_count):
        # 创建单页PDF
        single_page_pdf = fitz.open()
        single_page_pdf.insert_pdf(pdf_document, from_page=page_number, to_page=page_number)
        
        # 保存单页文件
        output_path = os.path.join(output_folder, f"{file_name}_page{page_number+1}.pdf")
        single_page_pdf.save(output_path)
        single_page_pdf.close()
        page_files.append(output_path)
    
    pdf_document.close()
    return page_files

def auto_crop_pdf(input_pdf_path, output_pdf_path, border_width=5):
    """自动裁剪单页PDF文件中的内容区域"""
    pdf_document = fitz.open(input_pdf_path)
    output_pdf = fitz.open()

    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]
        pix = page.get_pixmap()
        
        # 将 PDF 页面转换为 PIL 图像
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # 转换为灰度图
        image = image.convert("L")
        image_array = np.array(image)

        # 获取图像尺寸
        height, width = image_array.shape

        # 初始化裁剪边界
        left, top, right, bottom = width, height, 0, 0

        # 判断像素是否位于边框
        def is_border_pixel(x, y):
            return x < border_width or y < border_width or x >= width - border_width or y >= height - border_width
        
        # 寻找所有非白色像素，并且忽略边框
        for y in range(height):
            for x in range(width):
                if not is_border_pixel(x, y) and image_array[y, x] < 255:
                    left = min(left, x)
                    right = max(right, x)
                    top = min(top, y)
                    bottom = max(bottom, y)
                    
        # 如果整个页面都是白色或只有边框，则不裁剪
        if left == width and top == height and right == 0 and bottom == 0:
            new_page = output_pdf.new_page(width=pix.width, height=pix.height)
            new_page.show_pdf_page(new_page.rect, pdf_document, page_number)
            continue
        
        # 创建裁剪区域的矩形
        crop_rect = fitz.Rect(left, top, right + 1, bottom + 1)  # 注意加一操作

        # 创建新的 PDF 页面
        new_page = output_pdf.new_page(width=crop_rect.width, height=crop_rect.height)

        # 从原始页面提取并显示内容
        new_page.show_pdf_page(new_page.rect, pdf_document, page_number, clip=crop_rect)

    # 保存输出 PDF
    output_pdf.save(output_pdf_path)
    pdf_document.close()
    output_pdf.close()

def resize_pdf_page(input_pdf_path, output_pdf_path, target_width_mm=100, target_height_mm=150):
    """
    调整PDF页面大小为指定的毫米尺寸
    
    Args:
        input_pdf_path: 输入PDF文件路径
        output_pdf_path: 输出PDF文件路径
        target_width_mm: 目标宽度(毫米)
        target_height_mm: 目标高度(毫米)
    """
    # 毫米转换为点 (1mm = 2.83465点)
    target_width_pt = target_width_mm * 2.83465
    target_height_pt = target_height_mm * 2.83465
    
    # 打开原始PDF
    doc = fitz.open(input_pdf_path)
    
    # 创建新PDF
    new_doc = fitz.open()
    
    for page in doc:
        # 获取原始页面的边界框
        original_bbox = page.rect
        
        # 创建新页面
        new_page = new_doc.new_page(width=target_width_pt, height=target_height_pt)
        
        # 计算缩放比例
        scale_x = target_width_pt / original_bbox.width
        scale_y = target_height_pt / original_bbox.height
        
        # 使用最小值确保内容完整显示（保持纵横比）
        scale = min(scale_x, scale_y)
        
        # 计算缩放后的宽高
        scaled_width = original_bbox.width * scale
        scaled_height = original_bbox.height * scale
        
        # 计算偏移量以居中内容
        offset_x = (target_width_pt - scaled_width) / 2
        offset_y = (target_height_pt - scaled_height) / 2
        
        # 在新页面上创建用于内容的矩形（居中）
        content_rect = fitz.Rect(offset_x, offset_y, 
                                offset_x + scaled_width, 
                                offset_y + scaled_height)
        
        # 复制原始内容到新页面
        new_page.show_pdf_page(
            content_rect,   # 目标矩形（居中）
            doc,            # 源文档
            page.number     # 源页面索引
        )
    
    # 保存调整后的PDF
    new_doc.save(output_pdf_path)
    doc.close()
    new_doc.close()

def detect_barcode_in_pdf(pdf_path):
    """检测PDF文件中的条码并返回条码内容"""
    try:
        # 尝试使用poppler（如果可用）
        poppler = poppler_path.get() if poppler_path.get() else None
        
        # 在打包环境中使用资源路径
        if is_frozen and not poppler:
            poppler = resource_path("poppler/bin")
        
        # 将PDF页面转换为图像
        images = convert_from_path(pdf_path, dpi=200, grayscale=True, poppler_path=poppler)
        
        for img in images:
            # 转换为OpenCV格式
            open_cv_image = np.array(img)
            height, width = open_cv_image.shape
            
            # 根据快递面单特点，条码通常在右上角
            start_x = int(width * 0.6)  # 右半部分
            start_y = int(height * 0.1)  # 上半部分
            end_x = int(width * 0.95)    # 保留边缘安全距离
            end_y = int(height * 0.4)    # 保证条码完整
            
            cropped_img = open_cv_image[start_y:end_y, start_x:end_x]
            
            # 增强对比度（黑白图像特别有效）
            enhanced_img = cv2.convertScaleAbs(cropped_img, alpha=1.8, beta=40)
            
            # 检测条码
            barcodes = decode(enhanced_img)
            
            # 如果未检测到，尝试整个页面
            if not barcodes:
                barcodes = decode(open_cv_image)
            
            for barcode in barcodes:
                try:
                    barcode_data = barcode.data.decode("utf-8")
                    if barcode_data:
                        return barcode_data
                except UnicodeDecodeError:
                    try:
                        barcode_data = barcode.data.decode("latin-1")
                        if barcode_data:
                            return barcode_data
                    except:
                        continue
        return None
    except Exception as e:
        log_message(f"条码检测失败: {os.path.basename(pdf_path)} - {str(e)}")
        return None

def generate_rename_report(report_data, report_file_path):
    """生成重命名报告Excel文件"""
    try:
        # 确保报告目录存在
        report_dir = os.path.dirname(report_file_path)
        if report_dir and not os.path.exists(report_dir):
            os.makedirs(report_dir, exist_ok=True)
        
        # 如果文件已存在，先删除
        if os.path.exists(report_file_path):
            os.remove(report_file_path)
        
        # 创建DataFrame
        df = pd.DataFrame(report_data)
        
        # 确保列顺序一致
        columns = ["原始文件名", "页码", "新文件名", "条码内容"]
        df = df[columns]
        
        # 保存Excel文件
        df.to_excel(report_file_path, index=False, engine='openpyxl')
        return True
    except Exception as e:
        log_message(f"生成Excel报告失败: {str(e)}")
        return False

def check_poppler_installed():
    """检查poppler是否安装"""
    try:
        # 在打包环境中直接返回True
        if is_frozen:
            return True
            
        # 尝试导入pdf2image
        from pdf2image import pdfinfo_from_path
        # 尝试获取PDF信息
        with tempfile.NamedTemporaryFile(suffix='.pdf') as temp_pdf:
            # 创建一个空白PDF
            doc = fitz.open()
            doc.new_page(width=100, height=100)
            doc.save(temp_pdf.name)
            doc.close()
            
            # 尝试获取信息
            pdfinfo_from_path(temp_pdf.name)
        return True
    except Exception as e:
        return False

def search_log():
    """搜索日志内容"""
    global log_text
    search_term = search_entry.get().strip()
    if not search_term or not log_text:
        return
        
    # 清除之前的标记
    log_text.tag_remove("found", "1.0", tk.END)
    
    # 获取日志内容
    log_content = log_text.get("1.0", tk.END)
    
    # 使用正则表达式查找所有匹配项
    pattern = re.compile(re.escape(search_term), re.IGNORECASE)
    matches = list(pattern.finditer(log_content))
    
    if not matches:
        messagebox.showinfo("搜索", f"未找到匹配项: {search_term}")
        return
        
    # 标记所有匹配项
    for match in matches:
        start_index = f"1.0+{match.start()}c"
        end_index = f"1.0+{match.end()}c"
        log_text.tag_add("found", start_index, end_index)
    
    # 配置标记样式
    log_text.tag_config("found", background="yellow", foreground="black")
    
    # 滚动到第一个匹配项
    first_match = matches[0]
    start_index = f"1.0+{first_match.start()}c"
    log_text.see(start_index)

def clear_log():
    """清空日志内容"""
    global log_text
    if log_text:
        log_text.configure(state='normal')
        log_text.delete("1.0", tk.END)
        log_text.configure(state='disabled')

def process_pdf_files():
    """处理选择的多个PDF文件（使用多线程）"""
    global is_processing
    
    # 检查是否已有处理线程在运行
    if is_processing:
        log_message("警告: 已有处理任务正在运行", "warning")
        return
    
    file_paths = input_files_listbox.get(0, tk.END)
    border_width_str = border_width_entry.get()
    output_folder = output_folder_entry.get()
    enable_rename = enable_rename_var.get()
    enable_logging = enable_logging_var.get()
    report_file_path = report_path.get()

    if not file_paths:
        status_label.config(text="错误: 请选择 PDF 文件")
        log_message("错误: 请选择 PDF 文件", "error")
        return
    
    try:
        border_width = int(border_width_str)
    except ValueError:
        status_label.config(text="错误: 边框宽度必须是整数")
        log_message("错误: 边框宽度必须是整数", "error")
        return
    
    if not output_folder:
        output_folder = "output"  # 设置默认输出文件夹为 output
        os.makedirs(output_folder, exist_ok=True)  # 如果文件夹不存在，创建它
    
    # 检查输出文件夹是否有效
    if not os.path.isdir(output_folder):
        try:
            os.makedirs(output_folder, exist_ok=True)
        except Exception as e:
            message = f"无法创建输出文件夹: {str(e)}"
            messagebox.showerror("错误", message)
            status_label.config(text=message)
            log_message(message, "error")
            return

    # 禁用处理按钮
    process_button.config(state=tk.DISABLED)
    is_processing = True
    status_label.config(text="正在处理，请稍候...")
    log_message("开始处理PDF文件...")
    
    # 创建后台处理线程
    processing_thread = threading.Thread(
        target=process_pdf_files_thread,
        args=(file_paths, border_width, output_folder, enable_rename, enable_logging, report_file_path),
        daemon=True
    )
    processing_thread.start()
    
    # 启动线程状态检查
    window.after(100, check_thread_status, processing_thread)

def process_pdf_files_thread(file_paths, border_width, output_folder, enable_rename, enable_logging, report_file_path):
    """PDF文件处理线程"""
    # 创建临时文件夹用于处理单页
    temp_folder = tempfile.mkdtemp()
    processed_files = []  # 保存处理后的文件路径
    report_data = []  # 保存重命名报告数据
    
    # 创建日志记录器（如果需要）
    logger = None
    if enable_logging:
        log_dir = os.path.join(output_folder, "日志")
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"处理日志_{timestamp}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),  # 记录到文件
                logging.StreamHandler()  # 同时在控制台输出
            ]
        )
        logger = logging.getLogger('PDF处理器')
        logger.info(f"===== 开始处理 PDF 文件 =====")
        logger.info(f"输出目录: {output_folder}")
        logger.info(f"边框宽度: {border_width} 像素")
        logger.info(f"启用重命名: {'是' if enable_rename else '否'}")
        logger.info(f"启用日志记录: {'是' if enable_logging else '否'}")
        logger.info(f"报告路径: {report_file_path}")
    
    try:
        for input_pdf_path in file_paths:
            # 检查输入文件是否存在
            if not os.path.isfile(input_pdf_path):
                msg = f"文件不存在: {input_pdf_path}"
                messagebox.showwarning("警告", msg)
                status_label.config(text=f"跳过不存在的文件: {os.path.basename(input_pdf_path)}")
                log_message(msg, "warning")
                if logger:
                    logger.warning(msg)
                window.update_idletasks()
                continue
                
            file_name = os.path.basename(input_pdf_path)
            base_name = os.path.splitext(file_name)[0]
            
            # 步骤1: 将PDF分割为单页
            window.after(0, lambda msg=f"分割 {file_name} 为单页...": status_label.config(text=msg))
            window.after(0, lambda: log_message(f"分割文件: {file_name}"))
            window.update_idletasks()
            if logger:
                logger.info(f"开始分割文件: {file_name}")
                
            try:
                page_files = split_pdf_to_single_pages(input_pdf_path, temp_folder)
                if logger:
                    logger.info(f"成功分割 {file_name} 为 {len(page_files)} 页")
                log_message(f"成功分割 {file_name} 为 {len(page_files)} 页")
            except Exception as e:
                msg = f"分割 {file_name} 时发生错误: {str(e)}"
                messagebox.showerror("错误", msg)
                status_label.config(text=f"分割 {file_name} 时发生错误")
                log_message(msg, "error")
                if logger:
                    logger.error(msg)
                window.update_idletasks()
                continue
            
            # 步骤2: 对每个单页进行裁剪和尺寸调整
            for i, page_file in enumerate(page_files):
                window.after(0, lambda msg=f"处理 {file_name} 第 {i+1} 页...": status_label.config(text=msg))
                window.after(0, lambda: log_message(f"处理第 {i+1} 页"))
                window.update_idletasks()
                if logger:
                    logger.info(f"开始处理第 {i+1} 页")
                
                # 裁剪后的临时文件名
                cropped_temp_name = f"{base_name}_page{i+1}_cropped_temp.pdf"
                cropped_temp_path = os.path.join(temp_folder, cropped_temp_name)
                
                try:
                    # 裁剪单页PDF
                    auto_crop_pdf(page_file, cropped_temp_path, border_width)
                    if logger:
                        logger.info(f"裁剪第 {i+1} 页完成")
                    log_message(f"裁剪第 {i+1} 页完成")
                    
                    # 最终输出的文件名
                    final_page_name = f"{base_name}_page{i+1}_final.pdf"
                    final_page_path = os.path.join(output_folder, final_page_name)
                    
                    # 调整页面大小到100x150mm
                    resize_pdf_page(cropped_temp_path, final_page_path, 100, 150)
                    
                    processed_files.append(final_page_path)
                    
                    # 更新状态
                    window.after(0, lambda msg=f"已完成 {file_name} 第 {i+1} 页的处理": status_label.config(text=msg))
                    if logger:
                        logger.info(f"调整大小完成: {final_page_name}")
                    log_message(f"调整大小完成: {final_page_name}")
                    
                    # 步骤3: 重命名文件（如果启用）
                    if enable_rename:
                        window.after(0, lambda msg=f"检测条码并重命名 {final_page_name}...": status_label.config(text=msg))
                        window.after(0, lambda: log_message(f"检测条码: {final_page_name}"))
                        window.update_idletasks()
                        if logger:
                            logger.info(f"开始检测条码: {final_page_name}")
                        
                        # 检测条码
                        barcode = detect_barcode_in_pdf(final_page_path)
                        
                        if barcode:
                            # 对条码内容进行特殊处理
                            original_barcode = barcode  # 保存原始条码内容
                            
                            # 条码处理规则
                            if barcode.startswith('4') and len(barcode) > 22:
                                barcode = barcode[-22:]  # 截取后22位
                            elif barcode.startswith('9') and len(barcode) > 22:
                                barcode = barcode[-12:]  # 截取后12位
                            
                            # 创建安全的新文件名
                            safe_barcode = ''.join(filter(str.isalnum, barcode))
                            
                            # 确保文件名长度合理
                            safe_barcode = safe_barcode[:50] if len(safe_barcode) > 50 else safe_barcode
                            
                            if safe_barcode:
                                # 创建唯一文件名
                                new_filename = f"{safe_barcode}.pdf"
                                new_file_path = os.path.join(output_folder, new_filename)
                                
                                # 重命名文件
                                os.rename(final_page_path, new_file_path)
                                
                                # 添加到报告
                                report_data.append({
                                    "原始文件名": file_name,
                                    "页码": i+1,
                                    "新文件名": new_filename,
                                    "条码内容": barcode
                                })
                                
                                window.after(0, lambda msg=f"已重命名为: {new_filename}": status_label.config(text=msg))
                                log_message(f"重命名成功: {final_page_name} -> {new_filename}")
                                if logger:
                                    logger.info(f"重命名成功: {final_page_name} -> {new_filename} (条码: {barcode})")
                            else:
                                msg = f"条码内容无效: {barcode}"
                                status_label.config(text=msg)
                                log_message(msg, "warning")
                                if logger:
                                    logger.warning(msg)
                        else:
                            msg = f"未检测到条码: {final_page_name}"
                            status_label.config(text=msg)
                            log_message(msg, "warning")
                            if logger:
                                logger.warning(msg)
                    
                    window.update_idletasks()
                except Exception as e:
                    msg = f"处理 {file_name} 第 {i+1} 页时发生错误: {str(e)}"
                    messagebox.showwarning("警告", msg)
                    status_label.config(text=f"处理 {file_name} 第 {i+1} 页时出错")
                    log_message(msg, "error")
                    if logger:
                        logger.error(msg)
                    window.update_idletasks()
    
    except Exception as e:
        msg = "处理 PDF 文件时发生错误: " + str(e)
        window.after(0, lambda: messagebox.showerror("错误", msg))
        window.after(0, lambda: status_label.config(text="发生错误，停止处理：" + str(e)))
        window.after(0, lambda: log_message(msg, "error"))
        if logger:
            logger.error(msg)
    finally:
        # 清理临时文件
        if os.path.exists(temp_folder):
            try:
                shutil.rmtree(temp_folder, ignore_errors=True)
                log_message("已清理临时文件夹")
                if logger:
                    logger.info("已清理临时文件夹")
            except Exception as e:
                log_message(f"清理临时文件夹时出错: {str(e)}", "error")
                if logger:
                    logger.error(f"清理临时文件夹时出错: {str(e)}")
    
    # 生成重命名报告（如果有数据）
    if enable_rename and report_data:
        report_generated = generate_rename_report(report_data, report_file_path)
        if report_generated:
            report_msg = f"重命名报告已生成: {report_file_path}"
            status_label.config(text=report_msg)
            log_message(report_msg)
            if logger:
                logger.info(report_msg)
        else:
            report_msg = "重命名报告生成失败"
            status_label.config(text=report_msg)
            log_message(report_msg, "error")
            if logger:
                logger.error(report_msg)
    
    if processed_files:
        msg = f"处理完成，共处理 {len(processed_files)} 页"
        status_label.config(text=msg)
        log_message(msg)
        messagebox.showinfo("完成", f"PDF 文件处理成功\n共处理了 {len(processed_files)} 页\n所有页面调整为100x150mm")
        if logger:
            logger.info(msg)
    else:
        msg = "处理完成，但未生成任何文件"
        status_label.config(text=msg)
        log_message(msg, "warning")
        messagebox.showwarning("警告", msg)
        if logger:
            logger.warning(msg)
    
    # 关闭日志记录器
    if logger:
        handlers = logger.handlers[:]
        for handler in handlers:
            handler.close()
            logger.removeHandler(handler)
    
    # 打开输出文件夹
    if processed_files:
        try:
            if os.name == 'nt':  # windows
                subprocess.Popen(['explorer', os.path.abspath(output_folder)])
            else:  # 其他系统
                subprocess.Popen(['open', os.path.abspath(output_folder)])
        except Exception as e:
            messagebox.showerror("错误", "无法打开输出文件夹：" + str(e))
            status_label.config(text="无法打开输出文件夹")
            log_message(f"无法打开输出文件夹: {str(e)}", "error")
        
        # 处理完成后启用按钮
        window.after(0, lambda: process_button.config(state=tk.NORMAL))
        window.after(0, lambda: status_label.config(text="处理完成"))
        global is_processing
        is_processing = False

def check_thread_status(thread):
    """检查线程状态并更新UI"""
    if thread.is_alive():
        window.after(100, check_thread_status, thread)
    else:
        # 线程完成后启用按钮
        process_button.config(state=tk.NORMAL)
        global is_processing
        is_processing = False

# ==================== 界面布局重构 ====================
# 创建主框架
main_frame = tk.Frame(window)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# 创建标签页容器
notebook = ttk.Notebook(main_frame)
notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# 创建文件处理标签页
file_tab = ttk.Frame(notebook)
notebook.add(file_tab, text="文件处理")

# 创建日志标签页
log_tab = ttk.Frame(notebook)
notebook.add(log_tab, text="处理日志")

# ==================== 文件处理标签页内容 ====================
# 左侧面板（文件选择和设置）
left_frame = ttk.Frame(file_tab)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

# PDF 文件选择框架
input_frame = ttk.Labelframe(left_frame, text="选择 PDF 文件")
input_frame.pack(fill=tk.X, padx=5, pady=5, ipadx=5, ipady=5)

# PDF 文件列表框
input_files_listbox = tk.Listbox(input_frame, height=6)
input_files_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.BOTH, expand=True)

# 文件选择按钮框架
button_frame = ttk.Frame(input_frame)
button_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)

# PDF 文件选择按钮
select_file_button = ttk.Button(button_frame, text="选择文件", command=select_pdf_files)
select_file_button.pack(padx=5, pady=5, fill=tk.X)

# 清除文件按钮
clear_files_button = ttk.Button(button_frame, text="清除列表", command=lambda: input_files_listbox.delete(0, tk.END))
clear_files_button.pack(padx=5, pady=5, fill=tk.X)

# 输出设置框架
output_frame = ttk.Labelframe(left_frame, text="输出设置")
output_frame.pack(fill=tk.X, padx=5, pady=5, ipadx=5, ipady=5)

# 输出文件夹
output_folder_frame = ttk.Frame(output_frame)
output_folder_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(output_folder_frame, text="输出文件夹:").pack(side=tk.LEFT)
output_folder_entry = ttk.Entry(output_folder_frame)
output_folder_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
output_folder_entry.insert(0, "output")  # 默认输出文件夹
# 初始化报告路径
report_path.set(os.path.join(output_folder_entry.get(), "重命名报告.xlsx"))
select_output_button = ttk.Button(output_folder_frame, text="浏览", command=select_output_folder)
select_output_button.pack(side=tk.LEFT, padx=5, pady=5)

# 边框宽度设置
border_frame = ttk.Frame(output_frame)
border_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(border_frame, text="边框宽度(像素):").pack(side=tk.LEFT)
border_width_entry = ttk.Entry(border_frame, width=5)
border_width_entry.pack(side=tk.LEFT, padx=5, pady=5)
border_width_entry.insert(0, "-400")  # 默认值

# 重命名设置框架
rename_frame = ttk.Labelframe(left_frame, text="文件重命名设置")
rename_frame.pack(fill=tk.X, padx=5, pady=5, ipadx=5, ipady=5)

# 启用重命名选项
enable_rename_check = ttk.Checkbutton(rename_frame, text="启用文件重命名", variable=enable_rename_var)
enable_rename_check.pack(anchor=tk.W, padx=5, pady=2)

# 报告文件路径
report_frame = ttk.Frame(rename_frame)
report_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(report_frame, text="报告文件:").pack(side=tk.LEFT)
report_entry = ttk.Entry(report_frame, textvariable=report_path, state='readonly')  # 改为只读
report_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

# 日志设置
logging_frame = ttk.Frame(rename_frame)
logging_frame.pack(fill=tk.X, padx=5, pady=5)
enable_logging_check = ttk.Checkbutton(logging_frame, text="启用日志记录", variable=enable_logging_var)
enable_logging_check.pack(side=tk.LEFT, padx=5, pady=5)

# 处理按钮
process_frame = ttk.Frame(left_frame)
process_frame.pack(fill=tk.X, padx=5, pady=10)
process_button = ttk.Button(process_frame, text="开始处理", command=process_pdf_files)
process_button.pack(pady=5, ipadx=10, ipady=5)

# 状态标签
status_label = ttk.Label(left_frame, text="等待操作...", relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(fill=tk.X, padx=20, pady=5)

# 尺寸信息标签
size_info = ttk.Label(left_frame, text="所有页面将被调整为100mm x 150mm大小", relief=tk.FLAT, anchor=tk.CENTER)
size_info.pack(fill=tk.X, padx=20, pady=5)

# 尺寸信息标签下方添加依赖库路径标签
dep_info = ttk.Label(left_frame, text="", relief=tk.FLAT, anchor=tk.CENTER)
dep_info.pack(fill=tk.X, padx=20, pady=5)

# 在文件处理标签页中添加路径设置框架
path_frame = ttk.Labelframe(left_frame, text="依赖库路径设置")
path_frame.pack(fill=tk.X, padx=5, pady=5, ipadx=5, ipady=5)

# Poppler路径设置
poppler_frame = ttk.Frame(path_frame)
poppler_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(poppler_frame, text="Poppler路径:").pack(side=tk.LEFT)
poppler_entry = ttk.Entry(poppler_frame, textvariable=poppler_path)
poppler_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
poppler_button = ttk.Button(poppler_frame, text="浏览", command=select_poppler_path)
poppler_button.pack(side=tk.LEFT, padx=5, pady=5)

# libiconv2.dll路径显示
libiconv_frame = ttk.Frame(path_frame)
libiconv_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(libiconv_frame, text="libiconv2.dll路径:").pack(side=tk.LEFT)
libiconv_label = ttk.Label(libiconv_frame, text=resource_path("libiconv2.dll"))
libiconv_label.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

# 检查路径按钮
check_button = ttk.Button(path_frame, text="检查依赖库", command=check_dependencies)
check_button.pack(pady=5)

# ==================== 日志标签页内容 ====================
# 日志框架
log_frame = ttk.Frame(log_tab)
log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# 日志搜索框
search_frame = ttk.Frame(log_frame)
search_frame.pack(fill=tk.X, padx=5, pady=5)
ttk.Label(search_frame, text="搜索日志:").pack(side=tk.LEFT)
search_entry = ttk.Entry(search_frame)
search_entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
search_button = ttk.Button(search_frame, text="搜索", command=search_log)
search_button.pack(side=tk.LEFT, padx=5, pady=5)
clear_button = ttk.Button(search_frame, text="清空日志", command=clear_log)
clear_button.pack(side=tk.LEFT, padx=5, pady=5)

# 日志文本框（带滚动条）
log_scroll = ttk.Scrollbar(log_frame)
log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

log_text = tk.Text(
    log_frame, 
    wrap=tk.WORD, 
    yscrollcommand=log_scroll.set,
    state='disabled',
    width=100,
    background='white',  # 添加标准样式
    foreground='black'
)
log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
log_scroll.config(command=log_text.yview)

# 配置搜索结果的标记样式
log_text.tag_config("found", background="yellow", foreground="black")

# 检查poppler是否安装
if not check_poppler_installed() and not is_frozen:
    log_message("警告: Poppler未安装，条码检测功能可能无法正常工作！", "warning")
    log_message("请安装Poppler: https://github.com/oschwartz10612/poppler-windows/releases/", "warning")

# 程序启动时检查依赖库路径
check_dependencies()

# 在程序启动时检查DLL
if not check_dll_files():
    log_message("警告: 部分依赖库缺失，功能可能受限", "warning")

# 运行 GUI 窗口
window.mainloop()
