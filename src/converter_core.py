# -*- coding: utf-8 -*-
"""
天天采配商品清单 PDF → 标准格式 Excel  —— 转换核心

自动完成：
  1. 基于矩形网格 + 文字坐标重建表格（不依赖 pdf 表格识别，页尾行不会丢）
  2. 跨页断裂行自动合并（品名等长文本被翻页截断时自动拼接）
  3. 多区块自动识别（中农/含费价 11 列区块、品牌价 9 列区块）
  4. 每份源表输出到独立工作表，保留原列数与原序号
  5. 套用《欢迎访问天天采配》模板的完整单元格样式
"""

import os
import re
from collections import Counter

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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
