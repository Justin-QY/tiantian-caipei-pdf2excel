# -*- coding: utf-8 -*-
r"""
天天采配清单 PDF → Excel 转换器  —— 合并版源码（单文件）

把「转换核心 + 图形界面 + 界面 HTML」全部合进这一个 .py，方便保存和修改。

运行（需先装依赖）：
    pip install pdfplumber openpyxl pywebview
    python 天天采配清单-PDF转Excel_源码.py

命令行模式：
    python 天天采配清单-PDF转Excel_源码.py --convert a.pdf --out D:\导出目录 --layout single

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


# ============================ 第一部分：转换核心 ============================
# ============================ 模板样式（克隆参考文件） ============================
HEADERS = ["序号", "编码", "品名", "规格", "产地", "单位", "采价地点",
           "中农参考价", "中农浮动参考价", "含费参考价", "含费浮动参考价",
           "品牌参考价", "品牌浮动参考价"]
COL_WIDTHS = {"A": 8, "B": 10, "C": 40, "D": 30, "E": 23, "F": 8, "G": 40,
              "H": 13, "I": 17, "J": 13, "K": 17, "L": 13, "M": 17}

F_TITLE = Font(name="微软雅黑", size=16, bold=True, color="FFFFFFFF")
F_META = Font(name="微软雅黑", size=10, color="FF000000")
F_HEAD = Font(name="微软雅黑", size=11, bold=True, color="FFFFFFFF")
F_DATA = Font(name="Calibri", size=11, color="FF000000")
FILL_TITLE = PatternFill(fill_type="solid", fgColor="FF17365D")
FILL_HEAD = PatternFill(fill_type="solid", fgColor="FF5B9BD5")
BORDER_HEAD = Border(bottom=Side(style="thin", color="FFBFBFBF"))
A_TITLE = Alignment(horizontal="center", vertical="center")
A_META_L = Alignment(horizontal="left", vertical="center")
A_META_R = Alignment(horizontal="right", vertical="center")
A_META_W = Alignment(horizontal="left", vertical="center", wrap_text=True)
A_HEAD = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_CENTER = Alignment(horizontal="center", vertical="center")
A_LEFT = Alignment(horizontal="left", vertical="center")
A_RIGHT = Alignment(horizontal="right", vertical="center")

PUBLIC_COLS = 7                     # 前 7 列（序号~采价地点）为各区块公共列
SUPPORTED_TABLE_COLS = {11, 9}
OUTPUT_MODE = "single"              # "single" or "separate"
PLACEHOLDER = ("", "？", "?", "-", "—", "/", "\\")


class ConvertError(Exception):
    pass


# ============================ 网格提取 ============================
def _cluster(items, tol=2.0):
    """把 [(值, 权重)] 按容差聚成 [(中心, 总权重)]"""
    if not items:
        return []
    items = sorted(items)
    groups, cur = [], [items[0]]
    for it in items[1:]:
        if it[0] - cur[-1][0] <= tol:
            cur.append(it)
        else:
            groups.append(cur)
            cur = [it]
    groups.append(cur)
    out = []
    for g in groups:
        w = sum(x[1] for x in g)
        out.append((round(sum(x[0] * x[1] for x in g) / w, 1), w))
    return out


def _grid(pg):
    cx, cy = Counter(), Counter()
    for r in pg.rects:
        cx[round(r["x0"], 1)] += 1
        cx[round(r["x1"], 1)] += 1
        cy[round(r["top"], 1)] += 1
        cy[round(r["bottom"], 1)] += 1
    xs = _cluster([(float(k), v) for k, v in cx.items() if v >= 2])
    ys = _cluster([(float(k), v) for k, v in cy.items() if v >= 2])
    return xs, ys


def _pick(lines, ratio=0.45, minw=2):
    """保留贯穿足够多行的边界线"""
    if not lines:
        return []
    mx = max(w for _, w in lines)
    th = max(minw, mx * ratio)
    return [c for c, w in lines if w >= th]


def extract_records(pdf_path, progress=None, log=None):
    """
    返回 (recs, blocks, header)
      recs   : [{'seq':int, 'page':int, 'cells':{列下标:文本}}]
      blocks : [{'ncol':int,'rows':int,'lo':int,'hi':int,'missing':[...],'p0':int,'p1':int}]
      header : {'title','print_date','table_no','note','ref_month'}
    """
    log = log or (lambda *_: None)
    recs = []
    table_signature = None
    table_id = 0
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception as e:
        raise ConvertError(f"无法打开 PDF：{e}")

    with pdf:
        total = len(pdf.pages)
        header = {"title": "", "print_date": "", "table_no": "", "note": "", "ref_month": ""}
        for pi in range(total):
            if progress:
                progress(pi, total, f"解析第 {pi+1}/{total} 页")
            pg = pdf.pages[pi]

            if pi == 0:
                header = _parse_header(pg.extract_text() or "")

            xs_c, ys_c = _grid(pg)
            xs = sorted(_pick(xs_c, 0.45))
            ys = sorted(_pick(ys_c, 0.45))
            # 列边界不足 6 条时，沿用上一页的列定义
            if len(xs) < 6 or len(ys) < 2:
                continue

            cols = list(zip(xs[:-1], xs[1:]))
            rows = list(zip(ys[:-1], ys[1:])) + [(ys[-1], 1e9)]

            cells = {}
            for w in pg.extract_words():
                xm = (w["x0"] + w["x1"]) / 2
                ym = (w["top"] + w["bottom"]) / 2
                ci = next((i for i, (a, b) in enumerate(cols) if a <= xm < b), None)
                if ci is None:
                    if xm < xs[0]:
                        ci = 0
                    else:
                        continue
                ri = next((i for i, (a, b) in enumerate(rows) if a <= ym < b), None)
                if ri is None:
                    continue
                cells.setdefault((ri, ci), []).append(w)

            for ri in sorted({r for r, _ in cells}):
                rowtxt = {}
                for (r, c), wl in cells.items():
                    if r != ri:
                        continue
                    wl.sort(key=lambda w: (round(w["top"]), w["x0"]))
                    rowtxt[c] = "".join(w["text"] for w in wl)
                full = "".join(rowtxt.values())
                if "价格说明" in full or "序号" in full:
                    if "序号" in full:
                        signature = tuple((c, re.sub(r"\s+", "", t))
                                          for c, t in sorted(rowtxt.items()) if t.strip())
                        if table_signature is not None and signature != table_signature:
                            table_id += 1
                        table_signature = signature
                    continue                                 # 页眉/表头区
                seq = rowtxt.get(0, "").strip()
                if re.fullmatch(r"\d+", seq):
                    recs.append({"seq": int(seq), "page": pi + 1,
                                 "table_id": table_id, "cells": dict(rowtxt)})
                elif recs:
                    for c, t in rowtxt.items():              # 续行 → 追加到上一条
                        if c == 0 and re.fullmatch(r"\d+", t.strip()):
                            continue
                        recs[-1]["cells"][c] = recs[-1]["cells"].get(c, "") + t

    if not recs:
        raise ConvertError("未能从该 PDF 中提取到清单数据，请确认文件格式。")

    blocks = _split_blocks(recs)
    return recs, blocks, header


def _split_blocks(recs):
    """按表头变化或序号回退切分区块，并统计完整性问题"""
    raw, cur = [], []
    for r in recs:
        if cur and (r.get("table_id", 0) != cur[-1].get("table_id", 0)
                    or r["seq"] <= cur[-1]["seq"]):
            raw.append(cur)
            cur = []
        cur.append(r)
    if cur:
        raw.append(cur)

    blocks = []
    for b in raw:
        seqs = [r["seq"] for r in b]
        ncol = max(max(int(k) for k in r["cells"]) for r in b) + 1
        blocks.append({
            "ncol": ncol,
            "rows": len(b),
            "lo": min(seqs),
            "hi": max(seqs),
            "missing": sorted(set(range(min(seqs), max(seqs) + 1)) - set(seqs)),
            "p0": b[0]["page"],
            "p1": b[-1]["page"],
        })
    return blocks


def _parse_header(text):
    """从首页文本解析标题 / 打印日期 / 表号 / 价格说明 / 参考月"""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    h = {"title": "", "print_date": "", "table_no": "", "note": "", "ref_month": ""}
    if lines:
        h["title"] = re.sub(r"-{2,}", "—", lines[0]).strip()

    for i, l in enumerate(lines):
        m = re.search(r"打印日期[：:]\s*([0-9\-/.年月日]+)", l)
        if m and not h["print_date"]:
            h["print_date"] = m.group(1).strip()
        m = re.search(r"表号[：:]\s*(\S+)", l)
        if m and not h["table_no"]:
            h["table_no"] = m.group(1).strip()
        if "价格说明" in l and not h["note"]:
            note = l
            j = i
            while "。" not in note and j + 1 < len(lines):
                j += 1
                note += lines[j]
            h["note"] = re.sub(r"\s+", "", note)
    m = re.search(r"(\d{4})年(\d{1,2})月", h["note"])
    if m:
        h["ref_month"] = f"{m.group(1)}年{int(m.group(2)):02d}月"
    return h


# ============================ 数据规整 ============================
def _clean_text(v):
    """文本列：仅去首尾空白；「？」表示原文档占位，按模板保留"""
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _clean_price(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in PLACEHOLDER:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def compose_rows(recs, blocks):
    """按每个源表的原始列布局生成行数据，并保留源序号"""
    view = []
    pos = 0
    for bi, meta in enumerate(blocks):
        nc = meta["ncol"]
        if nc not in SUPPORTED_TABLE_COLS:
            raise ConvertError(f"发现未知的表结构（{nc} 列），无法确定价格列位置。")
        for r in recs[pos: pos + meta["rows"]]:
            c = r["cells"]
            out = [None] * nc
            out[0] = r["seq"]
            for i in range(1, PUBLIC_COLS):
                out[i] = _clean_text(c.get(i))
            for i in range(nc - PUBLIC_COLS):
                out[PUBLIC_COLS + i] = _clean_price(c.get(PUBLIC_COLS + i))
            view.append((r["seq"], r["page"], out, bi, nc))
        pos += meta["rows"]

    return view


# ============================ 生成 Excel ============================
def _table_headers(nc):
    if nc == 11:
        return HEADERS[:11]
    if nc == 9:
        return HEADERS[:7] + HEADERS[11:13]
    raise ConvertError(f"发现未知的表结构（{nc} 列），无法生成表头。")


def _build_workbook_separate_sheets(view, header, out_path, sheet_name="商品清单"):
    wb = Workbook()
    title = header.get("title") or "商品清单采购清单"
    blocks = {}
    for item in view:
        blocks.setdefault(item[3], []).append(item)

    sheet_counts = Counter()
    max_row = 0
    for bi, items in blocks.items():
        nc = items[0][4]
        if nc == 11:
            base_name = "中农含费价"
        elif nc == 9:
            base_name = "品牌价"
        else:
            raise ConvertError(f"发现未知的表结构（{nc} 列），无法生成工作表。")
        headers = _table_headers(nc)

        sheet_counts[base_name] += 1
        suffix = sheet_counts[base_name]
        name = base_name if suffix == 1 else f"{base_name}{suffix}"
        ws = wb.active if bi == next(iter(blocks)) else wb.create_sheet()
        ws.title = name if sheet_name == "商品清单" else f"{sheet_name}{suffix}"

        ws.cell(1, 1, title)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
        c = ws.cell(1, 1)
        c.font, c.fill, c.alignment = F_TITLE, FILL_TITLE, A_TITLE
        ws.row_dimensions[1].height = 30

        ws.cell(2, 1, f"打印日期：{header.get('print_date','')}".rstrip("："))
        ws.cell(2, 1).font = F_META
        ws.cell(2, 1).alignment = A_META_L
        ws.cell(2, nc, f"表号：{header.get('table_no','')}".rstrip("："))
        ws.cell(2, nc).font = F_META
        ws.cell(2, nc).alignment = A_META_R
        ws.row_dimensions[2].height = 20

        ws.cell(3, 1, header.get("note") or "")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=nc)
        c = ws.cell(3, 1)
        c.font, c.alignment = F_META, A_META_W
        ws.row_dimensions[3].height = 30

        for i, h in enumerate(headers, start=1):
            cell = ws.cell(4, i, h)
            cell.font, cell.fill = F_HEAD, FILL_HEAD
            cell.alignment, cell.border = A_HEAD, BORDER_HEAD
        ws.row_dimensions[4].height = 28

        for row_idx, (_seq, _page, row, _block_no, _nc) in enumerate(items, start=5):
            cell = ws.cell(row_idx, 1, row[0])
            cell.font, cell.alignment = F_DATA, A_CENTER
            cell = ws.cell(row_idx, 2, row[1])
            cell.font, cell.alignment, cell.number_format = F_DATA, A_LEFT, "@"
            for col in range(3, 8):
                cell = ws.cell(row_idx, col, row[col - 1])
                cell.font, cell.alignment = F_DATA, A_LEFT
            for col in range(8, nc + 1):
                cell = ws.cell(row_idx, col, row[col - 1])
                cell.font, cell.alignment, cell.number_format = F_DATA, A_RIGHT, "0.00"

        for col in range(1, nc + 1):
            key = chr(64 + col)
            ws.column_dimensions[key].width = COL_WIDTHS[key]
        ws.freeze_panes = "A5"
        ws.sheet_view.showGridLines = False
        max_row = max(max_row, 4 + len(items))

    if not blocks:
        ws = wb.active
        ws.title = sheet_name
    wb.properties.title = sheet_name
    wb.save(out_path)
    return max_row


def _build_workbook_single_sheet(view, header, out_path, sheet_name="商品清单"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    max_cols = max((item[4] for item in view), default=11)

    title = header.get("title") or "商品清单采购清单"
    ws.cell(1, 1, title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    c = ws.cell(1, 1)
    c.font, c.fill, c.alignment = F_TITLE, FILL_TITLE, A_TITLE
    ws.row_dimensions[1].height = 30

    ws.cell(2, 1, f"打印日期：{header.get('print_date','')}".rstrip("："))
    ws.cell(2, 1).font, ws.cell(2, 1).alignment = F_META, A_META_L
    ws.cell(2, max_cols, f"表号：{header.get('table_no','')}".rstrip("："))
    ws.cell(2, max_cols).font, ws.cell(2, max_cols).alignment = F_META, A_META_R
    ws.row_dimensions[2].height = 20

    note = header.get("note") or ""
    ws.cell(3, 1, note)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_cols)
    c = ws.cell(3, 1)
    c.font, c.alignment = F_META, A_META_W
    ws.row_dimensions[3].height = 30

    grouped = {}
    for item in view:
        grouped.setdefault(item[3], []).append(item)

    row_cursor = 4
    for block_index, items in grouped.items():
        nc = items[0][4]
        if block_index:
            ws.cell(row_cursor, 1, note)
            ws.merge_cells(start_row=row_cursor, start_column=1,
                           end_row=row_cursor, end_column=nc)
            c = ws.cell(row_cursor, 1)
            c.font, c.alignment = F_META, A_META_W
            ws.row_dimensions[row_cursor].height = 24
            row_cursor += 1

        for col, label in enumerate(_table_headers(nc), start=1):
            cell = ws.cell(row_cursor, col, label)
            cell.font, cell.fill = F_HEAD, FILL_HEAD
            cell.alignment, cell.border = A_HEAD, BORDER_HEAD
        ws.row_dimensions[row_cursor].height = 28
        row_cursor += 1

        for _seq, _page, row, _block_no, _nc in items:
            cell = ws.cell(row_cursor, 1, row[0])
            cell.font, cell.alignment = F_DATA, A_CENTER
            cell = ws.cell(row_cursor, 2, row[1])
            cell.font, cell.alignment, cell.number_format = F_DATA, A_LEFT, "@"
            for col in range(3, 8):
                cell = ws.cell(row_cursor, col, row[col - 1])
                cell.font, cell.alignment = F_DATA, A_LEFT
            for col in range(8, nc + 1):
                cell = ws.cell(row_cursor, col, row[col - 1])
                cell.font, cell.alignment, cell.number_format = F_DATA, A_RIGHT, "0.00"
            row_cursor += 1

    for col in range(1, max_cols + 1):
        key = chr(64 + col)
        ws.column_dimensions[key].width = COL_WIDTHS[key]
    # 单工作表模式下多段表格共用一张表，冻结表头只对第一段有效，
    # 滚动到后面几段时表头会对不上，所以这里不冻结。
    ws.sheet_view.showGridLines = False
    wb.properties.title = sheet_name
    wb.save(out_path)
    return row_cursor - 1


def build_workbook(view, header, out_path, sheet_name="商品清单", layout_mode=None):
    mode = layout_mode or OUTPUT_MODE
    if mode == "single":
        return _build_workbook_single_sheet(view, header, out_path, sheet_name)
    if mode == "separate":
        return _build_workbook_separate_sheets(view, header, out_path, sheet_name)
    raise ConvertError(f"未知的输出方式：{mode}")


# ============================ 文件名 ============================
def guess_output_name(pdf_path, header=None):
    """导出文件名 = 导入文件名（仅把扩展名改为 .xlsx）

    例：欢迎访问天天采配--参考7月数据_2026年8月执行.pdf
        → 欢迎访问天天采配--参考7月数据_2026年8月执行.xlsx
    """
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    return base + ".xlsx"


# ============================ 对外主入口 ============================
def convert(pdf_path, out_dir, progress=None, log=None, layout_mode=None):
    """转换单个 PDF，返回结果字典"""
    log = log or (lambda *_: None)
    log(f"开始解析：{os.path.basename(pdf_path)}")
    recs, blocks, header = extract_records(pdf_path, progress, log)

    for bi, b in enumerate(blocks, 1):
        note = "序号完整" if not b["missing"] else f"缺号 {b['missing']}"
        log(f"  区块{bi}：{b['ncol']} 列 / {b['rows']} 行 / 序号 {b['lo']}-{b['hi']} / {note}")

    view = compose_rows(recs, blocks)
    out_name = guess_output_name(pdf_path, header)
    out_path = os.path.join(out_dir, out_name)
    mode = layout_mode or OUTPUT_MODE
    last_row = build_workbook(view, header, out_path, layout_mode=mode)

    log(f"  已导出：{out_name}（数据 {len(view)} 行）")
    return {
        "pdf": pdf_path,
        "output": out_path,
        "rows": len(view),
        "blocks": blocks,
        "header": header,
        "layout_mode": mode,
        "warnings": [f"区块{i+1} 序号缺号 {b['missing']}" for i, b in enumerate(blocks) if b["missing"]],
    }

# ============================ 第二部分：界面 HTML ============================
UI_HTML = r'''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>天天采配清单-PDF转Excel</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-user-select: none; user-select: none; }

  :root {
    --ink: #1b2230;
    --ink-2: #5a6577;
    --ink-3: #8d97a8;
    --accent: #3b6cf6;
    --accent-soft: rgba(59,108,246,.10);
    --ok: #16a37b;
    --warn: #d98324;
    --err: #d9564f;
    --line: rgba(255,255,255,.78);
    --glass: rgba(255,255,255,.55);
    --glass-2: rgba(255,255,255,.46);
    --radius: 16px;
  }

  html, body { height: 100%; overflow: hidden; }
  body {
    background: #eef1f7;
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", system-ui, sans-serif;
    color: var(--ink);
    -webkit-font-smoothing: antialiased;
  }

  /* 背景色块 */
  .backdrop { position: fixed; inset: 0; overflow: hidden; z-index: 0; }
  .blob { position: absolute; border-radius: 50%; filter: blur(78px); opacity: .52; }
  .b1 { width: 520px; height: 520px; left: -4%;  top: -16%; background: #a8c4ff; animation: drift 22s ease-in-out infinite; }
  .b2 { width: 480px; height: 480px; right: -2%; top: 2%;   background: #ffcfdd; animation: drift 26s ease-in-out infinite reverse; }
  .b3 { width: 560px; height: 560px; left: 22%;  bottom: -30%; background: #b8efdd; animation: drift 30s ease-in-out infinite; }
  .b4 { width: 380px; height: 380px; right: 12%; bottom: -12%; background: #ddd0ff; animation: drift 24s ease-in-out infinite reverse; }
  @keyframes drift {
    0%,100% { transform: translate3d(0,0,0) scale(1); }
    50%     { transform: translate3d(26px,-22px,0) scale(1.08); }
  }

  .app {
    position: relative; z-index: 1;
    height: 100vh;
    display: flex; flex-direction: column;
    background: var(--glass);
    backdrop-filter: blur(30px) saturate(180%);
    border: 1px solid rgba(255,255,255,.5);
    overflow: hidden;
  }

  /* 标题栏（可拖动） */
  .titlebar {
    height: 50px; flex: 0 0 50px;
    display: flex; align-items: center; gap: 12px;
    padding: 0 8px 0 18px;
    border-bottom: 1px solid rgba(255,255,255,.55);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .mark {
    width: 26px; height: 26px; border-radius: 8px;
    background: linear-gradient(135deg, #4f7dff, #2f5be0);
    display: grid; place-items: center;
    box-shadow: 0 4px 12px rgba(59,108,246,.30);
  }
  .mark svg { width: 14px; height: 14px; }
  .brand b { font-size: 13.5px; font-weight: 600; letter-spacing: .2px; }
  .brand span { font-size: 11.5px; color: var(--ink-3); }

  .spacer { flex: 1; }

  .status-pill {
    display: inline-flex; align-items: center; gap: 7px;
    height: 26px; padding: 0 12px; border-radius: 999px;
    background: rgba(255,255,255,.6);
    border: 1px solid rgba(255,255,255,.9);
    font-size: 11.5px; color: var(--ink-2);
  }
  .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--ink-3); transition: background .3s; }
  .dot.busy { background: var(--accent); animation: pulse 1.1s ease-in-out infinite; }
  .dot.done { background: var(--ok); }
  @keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.35;transform:scale(.75)} }

  .win-btns { display: flex; gap: 1px; }
  .win-btns i {
    width: 38px; height: 32px; border-radius: 7px;
    display: grid; place-items: center; cursor: pointer;
    color: var(--ink-2); font-style: normal; font-size: 12px;
    transition: background .15s, color .15s;
  }
  .win-btns i:hover { background: rgba(255,255,255,.8); }
  .win-btns i.close:hover { background: #e5484d; color: #fff; }

  /* 主体 */
  .body { flex: 1; display: flex; flex-direction: column; gap: 8px; padding: 9px 12px 11px; min-height: 0; }

  .card {
    background: var(--glass-2);
    border: 1px solid rgba(255,255,255,.72);
    border-radius: var(--radius);
    box-shadow: 0 2px 10px rgba(26,38,66,.04);
    display: flex; flex-direction: column;
    min-height: 0;
  }
  .card-hd { display: flex; align-items: center; gap: 8px; padding: 10px 16px 6px; }
  .card-hd h2 { font-size: 12.5px; font-weight: 600; letter-spacing: .3px; }
  .card-hd .hint { font-size: 11.5px; color: var(--ink-3); }
  .card-hd .right { margin-left: auto; display: flex; gap: 4px; }

  .ghost {
    background: none; border: 0; cursor: pointer;
    font: inherit; font-size: 11.5px; color: var(--ink-3);
    padding: 4px 8px; border-radius: 7px; transition: all .15s;
  }
  .ghost:hover { background: rgba(255,255,255,.8); color: var(--accent); }
  .ghost:disabled { opacity: .4; cursor: default; }

  /* 导入区 */
  .drop {
    margin: 3px 16px 8px; flex: 0 0 auto; height: 84px;
    border-radius: 13px;
    border: 1.5px dashed rgba(59,108,246,.30);
    background: linear-gradient(180deg, rgba(255,255,255,.5), rgba(255,255,255,.24));
    display: flex; align-items: center; justify-content: center; gap: 14px;
    cursor: pointer; transition: all .18s;
  }
  .drop:hover { border-color: rgba(59,108,246,.62); background: rgba(255,255,255,.62); }
  .drop:active { background: var(--accent-soft); }
  .drop .icon { width: 30px; height: 30px; color: var(--accent); opacity: .85; flex: 0 0 auto; }
  .drop .txt p { font-size: 12.5px; color: var(--ink-2); }
  .drop .txt p em { font-style: normal; color: var(--accent); font-weight: 600; }
  .drop .txt small { display: block; margin-top: 4px; font-size: 11px; color: var(--ink-3); }
  .drop.mini { height: 38px; gap: 9px; margin: 2px 16px 6px; }
  .drop.mini .icon { width: 17px; height: 17px; }
  .drop.mini .txt small { display: none; }
  .drop.mini .txt p { font-size: 11.5px; }

  /* 文件列表 */
  .files { flex: 1; overflow-y: scroll; padding: 0 4px 0 10px; min-height: 0; }
  .file { display: flex; align-items: center; gap: 9px; padding: 6px 8px; border-radius: 10px; transition: background .15s, opacity .15s; }
  .file:hover { background: rgba(255,255,255,.66); }
  .file .cbx { width: 15px; height: 15px; flex: 0 0 auto; margin: 0; cursor: pointer; accent-color: var(--accent); }
  .file.off { opacity: .5; }
  .file.off:hover { opacity: .8; }
  .file .fi {
    width: 24px; height: 24px; border-radius: 6px; flex: 0 0 auto;
    background: rgba(217,86,79,.10); color: var(--err);
    display: grid; place-items: center; font-size: 8.5px; font-weight: 700;
  }
  .file .fname { font-size: 12px; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file .fmeta { font-size: 11px; color: var(--ink-3); flex: 0 0 auto; }
  .file .fstate { font-size: 11px; flex: 0 0 auto; padding: 2px 8px; border-radius: 999px; background: rgba(141,151,168,.14); color: var(--ink-3); }
  .file .fstate.run  { background: var(--accent-soft); color: var(--accent); }
  .file .fstate.done { background: rgba(22,163,123,.13); color: var(--ok); }
  .file .fstate.fail { background: rgba(217,86,79,.12); color: var(--err); }
  .file .del { opacity: 0; transition: opacity .15s; }
  .file:hover .del { opacity: 1; }
  .empty { padding: 4px 10px 0; font-size: 11.5px; color: var(--ink-3); }

  /* 设置条 */
  .settings {
    flex: 0 0 auto;
    display: flex; flex-direction: column; gap: 6px;
    padding: 7px 16px 8px;
    border-top: 1px solid rgba(255,255,255,.7);
    background: rgba(255,255,255,.26);
    border-radius: 0 0 16px 16px;
  }
  .set-row { display: flex; align-items: center; gap: 10px; min-width: 0; }
  .set-item { display: flex; align-items: center; gap: 8px; min-width: 0; }
  .set-item > label { font-size: 11.5px; color: var(--ink-3); flex: 0 0 auto; }
  .dirwrap { display: flex; align-items: center; gap: 6px; min-width: 0; flex: 1; }
  .dirwrap input, .namewrap input {
    width: 100%; min-width: 80px; height: 26px; padding: 0 9px;
    border-radius: 8px; border: 1px solid rgba(255,255,255,.9);
    background: rgba(255,255,255,.62); font: inherit; font-size: 11.5px; color: var(--ink);
    outline: none; transition: all .16s;
    -webkit-user-select: text; user-select: text;
  }
  .dirwrap input:focus, .namewrap input:focus { border-color: var(--accent); background: #fff; box-shadow: 0 0 0 3px var(--accent-soft); }
  .dirwrap input[readonly] { color: var(--ink-2); }
  .namewrap { flex: 0 0 220px; min-width: 120px; }
  .namewrap input::placeholder { color: #a8b1c1; }
  .chips { display: flex; gap: 5px; flex: 0 0 auto; }
  .chip {
    height: 26px; padding: 0 10px; border-radius: 8px;
    border: 1px solid rgba(255,255,255,.9);
    background: rgba(255,255,255,.55);
    font: inherit; font-size: 11.5px; color: var(--ink-2);
    cursor: pointer; transition: all .16s; white-space: nowrap;
  }
  .chip:hover { background: rgba(255,255,255,.85); }
  .chip.on { background: var(--accent); border-color: var(--accent); color: #fff; box-shadow: 0 3px 10px rgba(59,108,246,.26); }
  .mini-btn {
    height: 26px; padding: 0 10px; border-radius: 8px; flex: 0 0 auto;
    border: 1px solid rgba(255,255,255,.9);
    background: rgba(255,255,255,.62);
    font: inherit; font-size: 11.5px; color: var(--ink-2);
    cursor: pointer; transition: all .16s; white-space: nowrap;
  }
  .mini-btn:hover { background: #fff; color: var(--ink); }

  /* 进度 */
  .progress-card {
    flex: 0 0 auto;
    display: flex; align-items: center; gap: 16px;
    padding: 10px 16px;
    background: var(--glass-2);
    border: 1px solid rgba(255,255,255,.72);
    border-radius: var(--radius);
  }
  .progress-main { flex: 1; min-width: 0; }
  .progress-top { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin-bottom: 7px; }
  .progress-top b {
    font-size: 12px; font-weight: 600; min-width: 0;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  .progress-top em { font-style: normal; font-size: 11.5px; color: var(--ink-3); font-variant-numeric: tabular-nums; flex: 0 0 auto; }
  .track { height: 6px; border-radius: 999px; background: rgba(141,151,168,.16); overflow: hidden; }
  .bar {
    height: 100%; width: 0%; border-radius: 999px;
    background: linear-gradient(90deg, #6d95ff, #3b6cf6);
    transition: width .25s ease; position: relative; overflow: hidden;
  }
  .bar::after {
    content: ""; position: absolute; inset: 0;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,.55), transparent);
    transform: translateX(-100%);
  }
  .bar.active::after { animation: shimmer 1.2s linear infinite; }
  @keyframes shimmer { to { transform: translateX(100%); } }

  .actions { display: flex; gap: 9px; flex: 0 0 auto; }
  .btn {
    height: 36px; padding: 0 18px; border-radius: 9px;
    font: inherit; font-size: 12.5px; font-weight: 500;
    cursor: pointer; transition: all .16s; white-space: nowrap;
    border: 1px solid rgba(255,255,255,.9);
    background: rgba(255,255,255,.62); color: var(--ink-2);
  }
  .btn:hover:not(:disabled) { background: #fff; color: var(--ink); box-shadow: 0 4px 14px rgba(26,38,66,.08); }
  .btn:active:not(:disabled) { transform: translateY(1px); }
  .btn.primary {
    background: linear-gradient(180deg, #4c7bff, #3b6cf6);
    border-color: transparent; color: #fff;
    box-shadow: 0 6px 18px rgba(59,108,246,.30);
  }
  .btn.primary:hover:not(:disabled) { background: linear-gradient(180deg, #5b87ff, #3f6ef8); color: #fff; }
  .btn:disabled { opacity: .42; cursor: not-allowed; }

  /* 日志 */
  .log-card { height: 146px; flex: 0 0 146px; display: flex; flex-direction: column; }
  .log-hd { display: flex; align-items: center; gap: 8px; padding: 8px 16px 6px; border-bottom: 1px solid rgba(255,255,255,.6); }
  .log-hd h2 { font-size: 11.5px; font-weight: 600; }
  .log-hd .count { font-size: 11px; color: var(--ink-3); }
  .log-hd .ghost { margin-left: auto; }
  .log-body {
    flex: 1; overflow-y: scroll; padding: 6px 10px 8px 16px;
    font-family: "Cascadia Mono", "Consolas", monospace;
    font-size: 11px; line-height: 1.6;
    -webkit-user-select: text; user-select: text;
  }
  .log-body div { display: flex; gap: 10px; }
  .log-body time { color: #a8b1c1; flex: 0 0 auto; }
  .log-body .msg { min-width: 0; white-space: pre-wrap; word-break: break-all; }
  .l-info .msg { color: var(--ink-2); }
  .l-ok   .msg { color: var(--ok); }
  .l-warn .msg { color: var(--warn); }
  .l-err  .msg { color: var(--err); }
  .l-dim  .msg { color: var(--ink-3); }

  .log-body::-webkit-scrollbar, .files::-webkit-scrollbar { width: 12px; height: 12px; }
  .log-body::-webkit-scrollbar-track, .files::-webkit-scrollbar-track {
    background: rgba(120,132,150,.14); border-radius: 6px; margin: 2px;
  }
  .log-body::-webkit-scrollbar-thumb, .files::-webkit-scrollbar-thumb {
    background: rgba(70,82,102,.72); border-radius: 6px;
    border: 3px solid transparent; background-clip: padding-box;
  }
  .log-body::-webkit-scrollbar-thumb:hover, .files::-webkit-scrollbar-thumb:hover {
    background: rgba(90,101,119,.75); background-clip: padding-box;
  }

  .toast {
    position: absolute; left: 50%; bottom: 24px; transform: translate(-50%, 12px);
    padding: 9px 17px; border-radius: 10px;
    background: rgba(27,34,48,.92); color: #fff; font-size: 12px;
    opacity: 0; visibility: hidden; transition: all .24s; z-index: 9;
  }
  .toast.show { opacity: 1; visibility: visible; transform: translate(-50%, 0); }
</style>
</head>
<body>
<div class="backdrop">
  <div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div><div class="blob b4"></div>
</div>

<div class="app">
  <div class="titlebar pywebview-drag-region">
    <div class="brand">
      <div class="mark">
        <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>
        </svg>
      </div>
      <b>天天采配清单-PDF转Excel</b><span>v1.5</span>
    </div>
    <div class="spacer"></div>
    <div class="status-pill"><i class="dot" id="dot"></i><span id="statusText">就绪</span></div>
    <div class="win-btns">
      <i id="btnMin" title="最小化">—</i>
      <i id="btnMax" title="最大化">▢</i>
      <i id="btnClose" class="close" title="关闭">✕</i>
    </div>
  </div>

  <div class="body">
    <div class="card" style="flex:1;">
      <div class="card-hd">
        <h2>待转换文件</h2><span class="hint" id="fileHint">未选择</span>
        <div class="right">
          <button class="ghost" id="btnCheckAll">全选</button>
          <button class="ghost" id="btnFolder">添加文件夹</button>
          <button class="ghost" id="btnClear">清空</button>
        </div>
      </div>

      <div class="drop" id="drop">
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5-5 5 5"/><path d="M12 5v12"/>
        </svg>
        <div class="txt">
          <p>点击选择 PDF 文件，或 <em>添加整个文件夹</em></p>
          <small>支持多选 · 仅处理文字版 PDF（天天采配导出）</small>
        </div>
      </div>

      <div class="files" id="fileList">
        <div class="empty" id="emptyTip">尚未添加文件</div>
      </div>

      <div class="settings">
        <div class="set-row">
          <div class="set-item" style="flex:1; min-width:0;">
            <label>导出目录</label>
            <div class="dirwrap">
              <input id="outDir" placeholder="选择 Excel 保存位置" readonly>
              <button class="mini-btn" id="btnBrowse">浏览…</button>
              <button class="mini-btn" id="btnOpen">打开目录</button>
            </div>
          </div>
        </div>
        <div class="set-row">
          <div class="set-item">
            <label>文件名</label>
            <div class="namewrap">
              <input id="outName" placeholder="留空 = 用原文件名" spellcheck="false">
            </div>
          </div>
          <div class="set-item" style="margin-left:auto;">
            <label>输出</label>
            <div class="chips" id="outChips">
              <button class="chip on" data-mode="single">单工作表</button>
              <button class="chip" data-mode="separate">多工作表</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="progress-card">
      <div class="progress-main">
        <div class="progress-top">
          <b id="taskLabel">等待开始</b>
          <em id="pct">0%</em>
        </div>
        <div class="track"><div class="bar" id="bar"></div></div>
      </div>
      <div class="actions">
        <button class="btn primary" id="btnStart" disabled>开始转换</button>
      </div>
    </div>

    <div class="card log-card">
      <div class="log-hd">
        <h2>处理日志</h2><span class="count" id="logCount">0 条</span>
        <button class="ghost" id="logClear">清空日志</button>
      </div>
      <div class="log-body" id="log"></div>
    </div>
  </div>

  <div class="toast" id="toast"></div>
</div>

<script>
  const $ = s => document.querySelector(s);
  const api = () => window.pywebview.api;
  let files = [];
  let busy = false;
  let layout = 'single';
  let runQueue = [];

  const now = () => new Date().toTimeString().slice(0, 8);
  const esc = s => String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));

  function log(text, level) {
    const el = document.createElement('div');
    el.className = 'l-' + (level || 'info');
    el.innerHTML = '<time>' + now() + '</time><span class="msg">' + esc(text) + '</span>';
    const box = $('#log');
    box.appendChild(el);
    while (box.children.length > 500) box.removeChild(box.firstChild);
    box.scrollTop = box.scrollHeight;
    $('#logCount').textContent = box.children.length + ' 条';
  }
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('show'), 2400);
  }
  function setStatus(text, cls) {
    $('#statusText').textContent = text;
    $('#dot').className = 'dot' + (cls ? ' ' + cls : '');
  }
  function setProgress(v, text) {
    $('#bar').style.width = v + '%';
    $('#pct').textContent = (Math.round(v * 10) / 10) + '%';
    if (text) $('#taskLabel').textContent = text;
    $('#bar').classList.toggle('active', v > 0 && v < 100);
  }

  const STATE_TEXT = { wait: '待处理', run: '处理中', done: '已转换', fail: '失败' };

  const checkedCount = () => files.filter(f => f.checked).length;

  function updateStartState() {
    const n = checkedCount();
    $('#fileHint').textContent = files.length ? ('已勾选 ' + n + ' / ' + files.length) : '未选择';
    $('#btnStart').disabled = busy || n === 0;
    $('#btnStart').title = n === 0 ? '请先勾选要转换的文件' : '';
  }

  function render() {
    const list = $('#fileList');
    list.querySelectorAll('.file').forEach(n => n.remove());
    $('#emptyTip').style.display = files.length ? 'none' : 'block';
    $('#drop').classList.toggle('mini', files.length > 0);
    files.forEach((f, i) => {
      const el = document.createElement('div');
      el.className = 'file' + (f.checked ? '' : ' off');
      el.title = f.path;
      el.innerHTML =
        '<input type="checkbox" class="cbx" data-i="' + i + '"' + (f.checked ? ' checked' : '') + '>' +
        '<div class="fi">PDF</div>' +
        '<div class="fname">' + esc(f.name) + '</div>' +
        '<div class="fmeta">' + esc(f.size) + '</div>' +
        '<div class="fstate ' + f.state + '">' + STATE_TEXT[f.state] + '</div>' +
        '<button class="ghost del" data-i="' + i + '">移除</button>';
      list.appendChild(el);
    });
    list.querySelectorAll('.cbx').forEach(cb => {
      cb.onchange = () => {
        if (busy) { cb.checked = !cb.checked; return; }
        files[+cb.dataset.i].checked = cb.checked;
        render();
      };
    });
    list.querySelectorAll('.del').forEach(b => {
      b.onclick = () => {
        if (busy) return;
        files.splice(+b.dataset.i, 1);
        render();
      };
    });
    updateStartState();
  }

  function addFiles(incoming) {
    let added = 0;
    (incoming || []).forEach(f => {
      if (!f || !f.path) return;
      if (files.some(x => x.path === f.path)) return;
      files.push({ path: f.path, name: f.name, size: f.size, state: 'wait', checked: true });
      added++;
    });
    render();
    if (added) {
      setStatus('待处理');
      log('已添加 ' + added + ' 个文件（默认勾选），当前共 ' + files.length + ' 个', 'info');
    }
  }

  async function call(method, ...args) {
    try {
      return await api()[method](...args);
    } catch (e) {
      log('调用失败：' + e, 'err');
      return { ok: false, msg: String(e) };
    }
  }

  async function pickFiles() {
    if (busy) return;
    const r = await call('pick_files');
    if (r && r.ok) addFiles(r.files);
    else if (r && r.msg && r.msg !== '已取消') toast(r.msg);
  }
  $('#drop').onclick = pickFiles;

  $('#btnFolder').onclick = async () => {
    if (busy) return;
    const r = await call('pick_folder');
    if (r && r.ok) addFiles(r.files);
    else if (r && r.msg && r.msg !== '已取消') toast(r.msg);
  };

  $('#btnClear').onclick = () => {
    if (busy) return;
    files = []; render(); setStatus('就绪');
    setProgress(0, '等待开始');
    log('已清空文件列表', 'dim');
  };

  $('#btnCheckAll').onclick = () => {
    if (busy || !files.length) return;
    const all = files.every(f => f.checked);
    files.forEach(f => f.checked = !all);
    render();
    log(all ? '已取消全部勾选' : '已勾选全部文件', 'dim');
  };

  $('#btnBrowse').onclick = async () => {
    if (busy) return;
    const r = await call('pick_out_dir');
    if (r && r.ok) { $('#outDir').value = r.dir; log('导出目录：' + r.dir, 'dim'); }
  };

  document.querySelectorAll('#outChips .chip').forEach(c => {
    c.onclick = () => {
      if (busy) return;
      document.querySelectorAll('#outChips .chip').forEach(x => x.classList.remove('on'));
      c.classList.add('on');
      layout = c.dataset.mode;
    };
  });

  $('#btnOpen').onclick = async () => {
    const d = $('#outDir').value.trim();
    if (!d) { toast('请先选择导出目录'); log('打开目录失败：尚未选择导出目录', 'err'); return; }
    const r = await call('open_dir', d);
    if (r && !r.ok) { toast(r.msg || '无法打开目录'); log('打开目录失败：' + (r.msg || '未知错误'), 'err'); }
    else log('已打开导出目录：' + d, 'dim');
  };

  $('#btnStart').onclick = async () => {
    if (busy) return;
    if (!files.length) { toast('请先导入 PDF 文件'); log('没有待处理文件', 'err'); return; }
    const chosen = files.filter(f => f.checked);
    if (!chosen.length) { toast('请先勾选要转换的文件'); log('未勾选任何文件，无法开始转换', 'err'); return; }
    const out = $('#outDir').value.trim();
    if (!out) { toast('请先选择导出目录'); log('无法开始：尚未选择导出目录', 'err'); $('#btnBrowse').focus(); return; }

    busy = true;
    $('#btnStart').disabled = true;
    $('#btnStart').textContent = '转换中…';
    $('#btnFolder').disabled = true;
    chosen.forEach(f => f.state = 'wait');
    render();
    setStatus('处理中', 'busy');
    setProgress(0, '准备中');
    log('='.repeat(40), 'dim');
    log('本次转换 ' + chosen.length + ' 个文件', 'dim');

    runQueue = chosen;
    const fileName = $('#outName').value.trim();
    const r = await call('start', chosen.map(f => f.path), out, layout, fileName);
    if (!r || !r.ok) {
      busy = false;
      $('#btnStart').disabled = false;
      $('#btnStart').textContent = '开始转换';
      $('#btnFolder').disabled = false;
      setStatus('就绪');
      updateStartState();
      toast((r && r.msg) || '启动失败');
      log((r && r.msg) || '启动失败', 'err');
    }
  };

  function onDone(ev) {
    busy = false;
    $('#btnStart').textContent = '开始转换';
    $('#btnFolder').disabled = false;
    $('#bar').classList.remove('active');
    updateStartState();
    runQueue = [];
    if (ev.fail === 0) {
      setStatus('已完成', 'done');
      setProgress(100, '转换完成：成功 ' + ev.ok + ' 个');
      log('转换完成：成功 ' + ev.ok + ' 个', 'ok');
      toast('转换完成，成功 ' + ev.ok + ' 个');
    } else if (ev.ok === 0) {
      setStatus('失败', '');
      $('#taskLabel').textContent = '转换失败';
      log('全部失败，请查看日志', 'err');
      toast('转换失败，请查看日志');
    } else {
      setStatus('部分完成');
      $('#taskLabel').textContent = '完成：成功 ' + ev.ok + '，失败 ' + ev.fail;
      log('部分完成：成功 ' + ev.ok + '，失败 ' + ev.fail, 'warn');
      toast('部分完成，成功 ' + ev.ok + ' 个');
    }
    (ev.errs || []).forEach(e => log('失败明细：' + e, 'err'));
    const d = $('#outDir').value.trim();
    if (d) log('全部结束，输出目录：' + d, 'dim');
  }

  function handleEvents(events) {
    (events || []).forEach(ev => {
      if (ev.type === 'log') log(ev.text, ev.level);
      else if (ev.type === 'progress') setProgress(ev.value, ev.text);
      else if (ev.type === 'file_state') {
        if (runQueue[ev.index]) { runQueue[ev.index].state = ev.state; render(); }
      } else if (ev.type === 'done') onDone(ev);
    });
  }

  let polling = false;
  async function tick() {
    if (polling) return;
    polling = true;
    try {
      const r = await api().poll();
      if (r && r.events && r.events.length) handleEvents(r.events);
    } catch (e) { /* 忽略抖动 */ }
    polling = false;
  }

  // 窗口按钮
  $('#btnMin').onclick = () => call('win_min');
  $('#btnMax').onclick = () => call('win_max');
  $('#btnClose').onclick = () => call('win_close');
  $('#logClear').onclick = () => { $('#log').innerHTML = ''; $('#logCount').textContent = '0 条'; };

  window.addEventListener('pywebviewready', () => {
    log('解析引擎就绪', 'ok');
    log('请添加天天采配导出的 PDF 清单文件', 'dim');
    updateStartState();
    setInterval(tick, 150);
  });
</script>
</body>
</html>

'''

def _ui_file():
    """把内嵌的界面 HTML 释放到临时目录，返回文件路径。"""
    d = os.path.join(tempfile.gettempdir(), "pdf2xlsx_ui")
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, "index.html")
    with open(p, "w", encoding="utf-8") as f:
        f.write(UI_HTML)
    return p



# ============================ 第三部分：程序主体 ============================
APP_NAME = "天天采配清单-PDF转Excel"
APP_VER = "v1.5"







def _fmt_size(n):
    if n >= 1048576:
        return "%.1f MB" % (n / 1048576.0)
    return "%d KB" % max(1, round(n / 1024.0))


def _safe_name(s):
    """去掉 Windows 文件名非法字符，并限制长度。"""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", s or "").strip(" .")
    return s[:80]


_WIN = None  # 主窗口，避免挂在 Api 上被 pywebview 递归序列化


class Api:
    """暴露给 JS 的接口。所有返回值必须是 JSON 可序列化对象。

    注意：属性一律用下划线开头，pywebview 会遍历非下划线属性做接口发现，
    把 Window / Queue 这类对象挂上去会导致无限递归。
    """

    def __init__(self):
        self._q = queue.Queue()
        self._busy = False
        self._lock = threading.Lock()

    # ---------------- 窗口 ----------------
    def win_min(self):
        if _WIN:
            _WIN.minimize()
        return {"ok": True}

    def win_max(self):
        if _WIN:
            try:
                if _WIN.state == "maximized":
                    _WIN.restore()
                else:
                    _WIN.maximize()
            except Exception:
                _WIN.maximize()
        return {"ok": True}

    def win_close(self):
        if _WIN:
            _WIN.destroy()
        return {"ok": True}

    # ---------------- 选择文件 ----------------
    def pick_files(self):
        try:
            r = _WIN.create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=True,
                file_types=("PDF 文件 (*.pdf)", "所有文件 (*.*)"))
        except Exception as e:
            return {"ok": False, "msg": "打开文件对话框失败：%s" % e}
        if not r:
            return {"ok": False, "msg": "已取消"}
        return {"ok": True, "files": [self._info(p) for p in r]}

    def pick_folder(self):
        try:
            d = _WIN.create_file_dialog(webview.FOLDER_DIALOG)
        except Exception as e:
            return {"ok": False, "msg": "打开文件夹对话框失败：%s" % e}
        if not d:
            return {"ok": False, "msg": "已取消"}
        folder = d[0] if isinstance(d, (list, tuple)) else d
        found = []
        for name in sorted(os.listdir(folder)):
            if name.lower().endswith(".pdf") and not name.startswith("~$"):
                found.append(os.path.join(folder, name))
        if not found:
            return {"ok": False, "msg": "该文件夹下没有 PDF 文件"}
        return {"ok": True, "files": [self._info(p) for p in found]}

    def pick_out_dir(self):
        try:
            d = _WIN.create_file_dialog(webview.FOLDER_DIALOG)
        except Exception as e:
            return {"ok": False, "msg": "打开文件夹对话框失败：%s" % e}
        if not d:
            return {"ok": False, "msg": "已取消"}
        folder = d[0] if isinstance(d, (list, tuple)) else d
        return {"ok": True, "dir": folder}

    def _info(self, p):
        p = os.path.abspath(p)
        try:
            size = _fmt_size(os.path.getsize(p))
        except OSError:
            size = "—"
        return {"path": p, "name": os.path.basename(p), "size": size}

    # ---------------- 打开目录 ----------------
    def open_dir(self, path):
        path = (path or "").strip()
        if path and os.path.isdir(path):
            try:
                os.startfile(path)  # noqa: S606
                return {"ok": True}
            except Exception as e:
                return {"ok": False, "msg": str(e)}
        return {"ok": False, "msg": "目录不存在"}

    # ---------------- 转换 ----------------
    def start(self, paths, out_dir, layout, name=""):
        if self._busy:
            return {"ok": False, "msg": "正在转换中，请稍候"}
        paths = [p for p in (paths or []) if os.path.isfile(p)]
        if not paths:
            return {"ok": False, "msg": "请先添加要转换的 PDF 文件"}
        out_dir = (out_dir or "").strip()
        if not out_dir:
            return {"ok": False, "msg": "请先选择导出目录"}
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            return {"ok": False, "msg": "无法创建导出目录：%s" % e}
        if layout not in ("single", "separate"):
            layout = "single"
        name = _safe_name((name or "").strip())

        self._busy = True
        threading.Thread(target=self._worker,
                         args=(paths, out_dir, layout, name), daemon=True).start()
        return {"ok": True}

    def _emit(self, ev):
        self._q.put(ev)

    @staticmethod
    def _final_path(tmp_file, out_dir, custom, multi, used):
        """把临时产物命名到最终文件名，遇到重名自动加序号，绝不覆盖已有文件。"""
        stem = os.path.splitext(os.path.basename(tmp_file))[0]
        if custom:
            stem = custom + "_" + stem if multi else custom
        stem = _safe_name(stem) or "商品清单"
        path = os.path.join(out_dir, stem + ".xlsx")
        root, ext = os.path.splitext(path)
        k = 1
        while os.path.normcase(os.path.abspath(path)) in used or os.path.exists(path):
            path = "%s(%d)%s" % (root, k, ext)
            k += 1
        used.add(os.path.normcase(os.path.abspath(path)))
        return path

    def _worker(self, paths, out_dir, layout, custom=""):
        ok = fail = 0
        outs, errs, used = [], [], set()
        n = len(paths)
        multi = n > 1
        tmpdir = os.path.join(out_dir, ".pdf2xlsx_tmp")
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
            os.makedirs(tmpdir, exist_ok=True)
        except Exception:
            tmpdir = out_dir

        self._emit({"type": "log", "level": "info",
                    "text": "导出目录：" + out_dir})
        if custom:
            self._emit({"type": "log", "level": "dim",
                        "text": "文件名：" + (custom + "_原文件名" if multi else custom + ".xlsx")})
        for i, pdf in enumerate(paths, 1):
            name = os.path.basename(pdf)
            self._emit({"type": "file_state", "index": i - 1, "state": "run"})
            self._emit({"type": "log", "level": "info",
                        "text": "[%d/%d] %s" % (i, n, name)})

            def prog(cur, total, txt, _i=i, _name=name):
                frac = (_i - 1 + cur / float(max(total, 1))) / n
                self._emit({"type": "progress", "value": round(frac * 100, 1),
                            "text": "%s：%s" % (_name, txt)})

            try:
                r = convert(
                    pdf, tmpdir, progress=prog,
                    log=lambda t: self._emit({"type": "log", "level": "dim",
                                              "text": "    " + t}),
                    layout_mode=layout)
                final = self._final_path(r["output"], out_dir, custom, multi, used)
                shutil.move(r["output"], final)
                r["output"] = final
                for w in r["warnings"]:
                    self._emit({"type": "log", "level": "warn", "text": "    ⚠ " + w})
                self._emit({"type": "log", "level": "ok",
                            "text": "    ✓ %s（%d 行）" % (
                                os.path.basename(r["output"]), r["rows"])})
                self._emit({"type": "file_state", "index": i - 1, "state": "done"})
                outs.append(r["output"])
                ok += 1
            except Exception as e:
                fail += 1
                errs.append("%s：%s" % (name, e))
                self._emit({"type": "file_state", "index": i - 1, "state": "fail"})
                self._emit({"type": "log", "level": "err",
                            "text": "    ✗ 失败：%s" % e})
                if not isinstance(e, ConvertError):
                    tb = traceback.format_exc().replace("\n", "\n    ")
                    self._emit({"type": "log", "level": "dim", "text": "    " + tb})

        shutil.rmtree(tmpdir, ignore_errors=True)
        self._emit({"type": "done", "ok": ok, "fail": fail,
                    "outs": outs, "errs": errs})
        self._busy = False

    def poll(self):
        events = []
        try:
            while True:
                events.append(self._q.get_nowait())
        except queue.Empty:
            pass
        return {"events": events, "busy": self._busy}


# ============================ 命令行模式 ============================
def _run_cli(argv):
    pdfs, out, layout = [], "", "single"
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--convert", "-c"):
            i += 1
            while i < len(argv) and not argv[i].startswith("--"):
                pdfs.append(argv[i]); i += 1
            continue
        if a in ("--out", "-o"):
            i += 1
            if i < len(argv):
                out = argv[i]; i += 1
            continue
        if a in ("--layout", "-l"):
            i += 1
            if i < len(argv):
                layout = argv[i].strip().lower(); i += 1
            continue
        i += 1
    if not pdfs or not out:
        return 2
    os.makedirs(out, exist_ok=True)
    lines = ["转换时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "=" * 46]
    ok = fail = 0
    for pdf in pdfs:
        lines.append("[%s]" % os.path.basename(pdf))
        try:
            r = convert(pdf, out, log=lambda t: lines.append("  " + t),
                           layout_mode=layout)
            for w in r["warnings"]:
                lines.append("  ⚠ " + w)
            lines.append("  ✓ %s（%d 行）\n" % (os.path.basename(r["output"]), r["rows"]))
            ok += 1
        except Exception as e:
            lines.append("  ✗ 失败：%s\n" % e)
            fail += 1
    lines.append("完成：成功 %d，失败 %d" % (ok, fail))
    with open(os.path.join(out, "转换日志.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return 0 if fail == 0 else 1


def main():
    if len(sys.argv) > 1 and any(a in ("--convert", "-c") for a in sys.argv[1:]):
        sys.exit(_run_cli(sys.argv[1:]))

    # 调试：设置 PDF2XLSX_DEBUG_PORT=9222 可开启远程调试端口
    dbg_port = os.environ.get("PDF2XLSX_DEBUG_PORT")
    if dbg_port:
        try:
            webview.settings["REMOTE_DEBUGGING_PORT"] = int(dbg_port)
        except Exception:
            pass

    global _WIN
    api = Api()
    window = webview.create_window(
        APP_NAME,
        url=_ui_file(),
        js_api=api,
        width=960, height=620,
        min_size=(880, 560),
        frameless=True,
        easy_drag=False,
        text_select=False,   # 关闭全局文本选择，pywebview 才会挂载标题栏拖动
        background_color="#EEF1F7",
    )
    _WIN = window
    webview.start(debug=bool(os.environ.get("PDF2XLSX_DEBUG") or dbg_port))
    return 0


if __name__ == "__main__":
    sys.exit(main())
