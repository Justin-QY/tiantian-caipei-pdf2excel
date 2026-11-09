# -*- coding: utf-8 -*-
r"""
天天采配商品清单 PDF → Excel   图形界面（pywebview 版）

· 界面：ui/index.html（HTML/CSS 渲染，磨玻璃视觉）
· 核心：converter_core.py（表结构识别、跨页断行复原、模板排版）
· 命令行：PDF清单转换器.exe --convert a.pdf b.pdf --out D:\out --layout single
"""

import os
import re
import sys
import queue
import shutil
import threading
import traceback
import datetime

import webview

APP_NAME = "天天采配清单-PDF转Excel"
APP_VER = "v1.5"


def res_path(rel):
    """打包后资源在 _MEIPASS 临时目录，源码运行时在当前目录。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


if getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import converter_core as cc  # noqa: E402


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
                r = cc.convert(
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
                if not isinstance(e, cc.ConvertError):
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
            r = cc.convert(pdf, out, log=lambda t: lines.append("  " + t),
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
        url=res_path(os.path.join("ui", "index.html")),
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
