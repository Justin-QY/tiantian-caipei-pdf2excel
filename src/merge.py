# -*- coding: utf-8 -*-
"""把 converter_core.py + app.py + ui/index.html 合并成单个 .py"""
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "天天采配清单-PDF转Excel_源码.py")


def read(p):
    with open(os.path.join(BASE, p), encoding="utf-8") as f:
        return f.read()


def strip_imports(src):
    out = []
    for line in src.splitlines():
        st = line.strip()
        if re.match(r"^(import\s+\w|from\s+\w[\w.]*\s+import\b)", st):
            continue
        out.append(line)
    return "\n".join(out)


def strip_docstring(src):
    m = re.match(r'\s*(?:#.*\n)*\s*(r?)("""|\'\'\')', src)
    if not m:
        return src
    q = m.group(2)
    end = src.find(q, m.end())
    if end == -1:
        return src
    return src[end + 3:]


core = strip_imports(strip_docstring(read("converter_core.py"))).strip()
app = strip_imports(strip_docstring(read("app.py"))).strip()
html = read(os.path.join("ui", "index.html"))

assert '"""' not in html, "HTML 里不能出现三引号"
assert "\\" not in html, "HTML 里不能出现反斜杠"

# app 里对核心的引用改成本模块内直接调用
app = app.replace("cc.convert(", "convert(")
app = app.replace("cc.ConvertError", "ConvertError")
app = app.replace("cc.", "")

# 去掉打包路径处理和多文件资源读取
app = re.sub(r"def res_path\(rel\):.*?return os\.path\.join\(base, rel\)\n", "", app, flags=re.S)
app = re.sub(r"if getattr\(sys, \"frozen\", False\):.*?sys\.path\.insert\(0, os\.path\.dirname\(os\.path\.abspath\(__file__\)\)\)\n",
             "", app, flags=re.S)

# 界面资源改为运行时释放到临时目录
app = app.replace(
    'url=res_path(os.path.join("ui", "index.html")),',
    "url=_ui_file(),")

helper = '''
def _ui_file():
    """把内嵌的界面 HTML 释放到临时目录，返回文件路径。"""
    d = os.path.join(tempfile.gettempdir(), "pdf2xlsx_ui")
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, "index.html")
    with open(p, "w", encoding="utf-8") as f:
        f.write(UI_HTML)
    return p


'''

head = '''# -*- coding: utf-8 -*-
r"""
天天采配清单 PDF → Excel 转换器  —— 合并版源码（单文件）

把「转换核心 + 图形界面 + 界面 HTML」全部合进这一个 .py，方便保存和修改。

运行（需先装依赖）：
    pip install pdfplumber openpyxl pywebview
    python 天天采配清单-PDF转Excel_源码.py

命令行模式：
    python 天天采配清单-PDF转Excel_源码.py --convert a.pdf --out D:\\导出目录 --layout single

重新打包 exe：
    pyinstaller --onefile --windowed --name "天天采配清单-PDF转Excel" ^
        --hidden-import clr --collect-all webview PDF清单转换器_合并版.py

结构：
    第一部分  转换核心（表结构识别 / 跨页断行复原 / Excel 排版）
    第二部分  界面 HTML（磨玻璃风格，含全部前端逻辑）
    第三部分  程序主体（pywebview 窗口 + 线程 + 与 JS 的接口）
"""

# ============================ 依赖 ============================
import os
import re
import sys
import queue
import shutil
import tempfile
import threading
import traceback
import datetime
from collections import Counter

import pdfplumber
import webview
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

'''

parts = [
    head,
    "\n# ============================ 第一部分：转换核心 ============================\n",
    core,
    "\n\n# ============================ 第二部分：界面 HTML ============================\n"
    "UI_HTML = r'''\n" + html + "\n'''\n",
    helper,
    "\n# ============================ 第三部分：程序主体 ============================\n",
    app,
    "\n",
]

with open(OUT, "w", encoding="utf-8") as f:
    f.write("".join(parts))

print("merged ->", OUT, os.path.getsize(OUT), "bytes")
