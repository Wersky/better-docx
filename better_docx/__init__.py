# -*- coding: utf-8 -*-
"""
better_docx —— 基于 python-docx 的 Word 文档高级处理库。

将原生 OOXML 能力封装成易用函数，分三个模块：
- helpers：表格 / 图片 / 字体 / 底纹等基础元素，以及分页异常诊断与修复
- layout：分节、页眉页脚、页码域、段落缩进、列表、全局样式、目录
- navigator：文档解析（JSON/Markdown）、锚点定位、定点注入、书签与超链接

用法（扁平导入）：`from better_docx import set_font, add_image, docx_to_json`
或子模块导入：`from better_docx.helpers import shade_cell` 等。
"""
from .helpers import *
from .layout import *
from .navigator import *

__version__ = "1.0.0"