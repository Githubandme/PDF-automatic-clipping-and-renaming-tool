![alt text](image.png)

# PDF 自动裁剪与重命名工具

一个基于Python的GUI工具，用于自动裁剪PDF文件、检测条码并根据条码内容重命名文件。特别适用于快递面单等包含条码的PDF文档处理。

## 功能特性

- **PDF分页处理**: 将多页PDF自动分割为单页文件
- **智能裁剪**: 自动检测并裁剪PDF中的有效内容区域
- **尺寸标准化**: 将所有页面调整为标准100mm x 150mm尺寸
- **条码识别**: 自动检测PDF中的一维条码/二维码
- **智能重命名**: 根据条码内容自动重命名文件
- **批量处理**: 支持同时处理多个PDF文件
- **处理报告**: 生成详细的重命名报告Excel文件
- **日志记录**: 完整的处理过程日志记录
- **用户友好界面**: 基于Tkinter的图形用户界面

## 系统要求

- Python 3.7+
- Windows/macOS/Linux

## 安装依赖

```bash
pip install -r requirements.txt
依赖包列表
 复制
 插入
 新文件

PyMuPDF>=1.23.0
Pillow>=9.0.0
numpy>=1.21.0
pandas>=1.3.0
opencv-python>=4.5.0
pyzbar>=0.1.8
pdf2image>=3.1.0
openpyxl>=3.0.9
外部依赖
Poppler
用于PDF到图像的转换，需要单独安装：

Windows: 下载 poppler-windows
macOS: brew install poppler
Ubuntu/Debian: sudo apt-get install poppler-utils
使用方法
运行程序：

bash
 复制
 插入
 运行

python PDF裁剪扫码.py
在"文件处理"标签页中：

点击"选择文件"添加PDF文件
设置输出文件夹
调整边框宽度参数（默认-400像素）
选择是否启用文件重命名功能
点击"开始处理"
在"处理日志"标签页中查看处理进度和结果

开源协议
本项目采用 GNU General Public License v3.0 开源协议。

使用的开源项目
本项目基于以下开源项目构建：

核心依赖
PyMuPDF - Apache License 2.0
PDF文档处理和操作
Pillow - PIL License
图像处理和格式转换
NumPy - BSD License
数值计算和数组操作
OpenCV - Apache License 2.0
计算机视觉和图像处理
pyzbar - MIT License
条码和二维码识别
pdf2image - MIT License
PDF转图像功能
数据处理
pandas - BSD License
数据分析和Excel文件操作
openpyxl - MIT License
Excel文件读写
外部工具
Poppler - GPL v2/v3
PDF渲染引擎
zbar - LGPL v2.1
条码识别库（pyzbar的底层依赖）
贡献指南
欢迎提交Issue和Pull Request！

Fork本项目
创建功能分支 (git checkout -b feature/AmazingFeature)
提交更改 (git commit -m 'Add some AmazingFeature')
推送到分支 (git push origin feature/AmazingFeature)
开启Pull Request
许可证
 复制
 插入
 新文件

PDF 自动裁剪与重命名工具
Copyright (C) 2024

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
联系方式
如有问题或建议，请通过GitHub Issues联系。

 复制
 插入
 新文件


## 创建requirements.txt文件

同时建议创建requirements.txt文件：

PyMuPDF>=1.23.0
Pillow>=9.0.0
numpy>=1.21.0
pandas>=1.3.0
opencv-python>=4.5.0
pyzbar>=0.1.8
pdf2image>=3.1.0
openpyxl>=3.0.9

 复制
 插入
 新文件


## 创建LICENSE文件

还需要创建GPL-3.0许可证文件，内容为标准的GNU General Public License v3.0文本。

这个README.md文件全面介绍了项目的功能、使用方法、依赖关系和开源协议信息，符合开源项目的标准文档要求。文件中详细列出了所有使用的开源项目及其对应的开源协议，确保了开源项目的合规性。
 
重新生成
AI生成


