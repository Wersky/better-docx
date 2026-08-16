# -*- coding: utf-8 -*-
"""pytest 配置：把项目根目录加入 sys.path，使 tests 无需安装即可 import better_docx。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))