# 天天采配清单 PDF 转 Excel

把「天天采配」导出的商品清单 PDF，按原始分表结构转换成排好版的 Excel。

Windows 桌面小工具，双击即用，不需要装 Python。

> 本项目是**桌面版**。浏览器访问请用另一个独立项目：[tiantian-caipei-pdf2excel-web](https://github.com/Justin-QY/tiantian-caipei-pdf2excel-web)

![界面](https://img.shields.io/badge/platform-Windows%2010%2F11%20x64-blue)
![界面](https://img.shields.io/badge/python-3.12-3776AB)

---

## 功能

- **批量导入**：单个/多个 PDF，或整个文件夹
- **勾选转换**：列表里勾选要转的文件，没勾选不给开始
- **表结构自动识别**：中农/含费价（11 列）和品牌价（9 列）按真实表头分别归位，不串列
- **跨页断行复原**：PDF 翻页时页尾行边框常被裁掉，普通解析器会整行丢弃；本工具用矩形网格 + 文字坐标重建行，品名被截断的部分会与下页顶部拼接
- **序号校验**：分段检查序号连续性，缺号写进日志
- **两种输出**：单工作表（多段连续排版）/ 多工作表（每张表独立）
- **自定义文件名**：留空用原文件名，也可指定前缀；重名自动加 `(1)(2)`，不会互相覆盖
- **实时日志**：每页解析、表段识别结果、导出文件名与行数、警告、失败堆栈全都有
- **转换前确认**：点「开始转换」先弹确认框，列出勾选数量、输出方式、文件名与导出目录；绿色「开始」/ 灰色「取消」，取消不会启动任何任务
- **服务三态指示**：右上角固定显示 就绪（绿，正常）/ 进行（黄，处理中）/ 异常（红，无法处理），异常时按钮同步禁用

## v1.6 更新

- 新增「是否开始执行转换文件？」确认弹框，支持 Esc 取消、Enter 开始、点遮罩取消
- 右上角状态由原来六种文案统一为三态，并用绿 / 黄 / 红区分；引擎未就绪时不再误显绿色
- 与网页版保持一致的交互和配色

## 界面

磨玻璃风格，单列布局，从上到下只有四块：文件列表 → 设置 → 进度 → 日志。

## 下载使用

到 [Releases](../../releases) 下载最新版 `TiantianCaipei-PDF2Excel-v1.6.exe`（用纯 ASCII 名避免下载时中文名损坏，下载后可随意改名），双击运行。

- 系统要求：Windows 10/11 64 位
- 依赖 WebView2 运行时（Win11 自带；Win10 若缺失，装一次
  [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/) 即可）
- exe 未做数字签名，部分杀软会提示，选「仍要运行」

## 从源码运行

```bash
pip install pdfplumber openpyxl pywebview
python "src/天天采配清单-PDF转Excel_源码.py"
```

`src/` 里有两个版本：

| 文件 | 说明 |
|---|---|
| `天天采配清单-PDF转Excel_源码.py` | **单文件版**，转换核心 + 界面 HTML + 程序主体全在一个 .py 里 |
| `app.py` + `converter_core.py` + `ui/index.html` | 模块化版，改界面改 `ui/index.html`，改解析改 `converter_core.py` |

单文件版和模块化版功能完全一致，单文件版由 `merge.py` 从模块化版生成。

## 命令行模式

```bash
天天采配清单-PDF转Excel.exe --convert "a.pdf" "b.pdf" --out "D:\导出目录" --layout single
```

`--layout` 可选 `single`（默认）/ `separate`。转换完成后在导出目录生成 `转换日志.txt`。

## 重新打包

```bash
pip install pyinstaller pdfplumber openpyxl pywebview
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "天天采配清单-PDF转Excel" ^
  --icon icon.ico --add-data "ui;ui" ^
  --hidden-import clr --collect-all webview app.py
```

产物在 `dist/`，约 34 MB。

## 输出格式

默认把所有源表连续排在一个工作表里。每段沿用自己的表头和列顺序，新表段之间重复价格说明和表头，不插空行。

```
中农/含费价表：11 列
  序号 | 编码 | 品名 | 规格 | 产地 | 单位 | 采价地点 |
  中农参考价 | 中农浮动参考价 | 含费参考价 | 含费浮动参考价

品牌价表：9 列
  序号 | 编码 | 品名 | 规格 | 产地 | 单位 | 采价地点 |
  品牌参考价 | 品牌浮动参考价
```

- 编码按文本存，前导零不丢（如 `0310185`）
- 价格是数值两位小数，可直接求和、透视
- 采价地点的 `？` 表示原文档无报价，按原样保留
- 表头冻结：单工作表模式不冻结（多段共用一张表，冻结只对第一段有效）；多工作表模式每表各自冻结

## 常见问题

**能处理扫描件吗？**
不能。只读 PDF 文字层，仅适用于天天采配导出的文字版 PDF。

**日志报「缺号」怎么办？**
说明该 PDF 有行没被识别。把缺号序号和对应 PDF 反馈过来。

**两个不同文件夹有同名 PDF 会覆盖吗？**
不会。导出时重名自动加 `(1)(2)`。

**双击没反应 / 窗口空白？**
检查 WebView2 运行时是否已安装。

## 技术栈

- 界面：HTML/CSS + [pywebview](https://pywebview.flowrl.com/) 6.x（WebView2 内核），磨玻璃用 `backdrop-filter`
- 解析：[pdfplumber](https://github.com/jsvine/pdfplumber)
- 输出：[openpyxl](https://openpyxl.readthedocs.io/)
- 打包：PyInstaller（onefile）
- Python 3.12
