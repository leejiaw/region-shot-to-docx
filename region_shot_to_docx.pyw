# -*- coding: utf-8 -*-
"""
区域截图 → Word

用法：
  1. 启动后按 F7（或点“框选区域”）在桌面上拖出一块区域。
  2. 之后每次按 F8（或点“立即截图”）就抓取这块固定区域。
  3. 全部截完后按 F9（或点“结束并生成 Word”），选择保存位置，
     每张截图自动排成 Word 里的一页。

快捷键（全局，切到别的窗口也生效）：F7 框选 / F8 截图 / F9 结束。
"""

from __future__ import annotations

import ctypes
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

APP_TITLE = "区域截图 → Word"
APP_BUILD = "v1.0"
APP_DIR_NAME = "RegionShotToDocx"
SETTINGS_FILE = (
    Path(os.environ.get("APPDATA") or tempfile.gettempdir())
    / APP_DIR_NAME
    / "settings.json"
)

MOD_NOREPEAT = 0x4000
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
WM_HOTKEY = 0x0312

VK_F7, VK_F8, VK_F9 = 0x76, 0x77, 0x78
HK_SELECT, HK_CAPTURE, HK_FINISH = 1, 2, 3

# (热键 id, 名称, ((修饰键, 虚拟键码, 显示名), 备用...))  —— 前一个被占用就自动退到后面
HOTKEY_SPECS = (
    (HK_SELECT, "框选", ((0, VK_F7, "F7"), (MOD_CONTROL | MOD_ALT, 0x31, "Ctrl+Alt+1"))),
    (HK_CAPTURE, "截图", ((0, VK_F8, "F8"), (MOD_CONTROL | MOD_ALT, 0x32, "Ctrl+Alt+2"))),
    (HK_FINISH, "结束", ((0, VK_F9, "F9"), (MOD_CONTROL | MOD_ALT, 0x33, "Ctrl+Alt+3"))),
)

SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77
SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN = 78, 79

A4_SHORT, A4_LONG = 21.0, 29.7
MARGIN_CM = 1.5


# --------------------------------------------------------------------------
# Win32 基础
# --------------------------------------------------------------------------
def enable_dpi_awareness() -> None:
    """在创建任何窗口前调用，保证 Tk 坐标与屏幕物理像素一致。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def virtual_screen() -> tuple[int, int, int, int]:
    u = ctypes.windll.user32
    return (
        u.GetSystemMetrics(SM_XVIRTUALSCREEN),
        u.GetSystemMetrics(SM_YVIRTUALSCREEN),
        u.GetSystemMetrics(SM_CXVIRTUALSCREEN),
        u.GetSystemMetrics(SM_CYVIRTUALSCREEN),
    )


def system_dpi() -> int:
    try:
        u = ctypes.windll.user32
        u.GetDpiForSystem.restype = ctypes.c_uint
        return int(u.GetDpiForSystem())
    except Exception:
        return 96


def grab_region(rect: tuple[int, int, int, int], delay_ms: int = 0):
    """按虚拟屏幕坐标抓取一块区域，返回 PIL.Image。"""
    from PIL import ImageGrab

    x, y, w, h = rect
    if delay_ms > 0:
        time.sleep(delay_ms / 1000.0)
    return ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)


def region_on_screen(rect: tuple[int, int, int, int]) -> bool:
    vx, vy, vw, vh = virtual_screen()
    x, y, w, h = rect
    return (
        w > 0
        and h > 0
        and x >= vx
        and y >= vy
        and x + w <= vx + vw
        and y + h <= vy + vh
    )


# --------------------------------------------------------------------------
# Word 生成
# --------------------------------------------------------------------------
def build_docx(
    out_path: str | Path,
    entries: list[dict],
    *,
    orientation: str = "auto",
    captions: bool = True,
) -> None:
    """entries: [{path, width, height, time}]，一张截图一页。"""
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt, RGBColor

    if not entries:
        raise ValueError("没有可写入的截图")

    aspect = entries[0]["width"] / entries[0]["height"]
    if orientation == "auto":
        landscape = aspect > 1.15
    else:
        landscape = orientation == "landscape"

    page_w, page_h = (A4_LONG, A4_SHORT) if landscape else (A4_SHORT, A4_LONG)
    content_w = page_w - 2 * MARGIN_CM
    content_h = page_h - 2 * MARGIN_CM
    # 给图注和行距留余量，避免“图片顶满 + 图注”把内容挤到下一页
    reserve_h = 1.0 if captions else 0.3
    max_h = content_h - reserve_h

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Cm(page_w)
    section.page_height = Cm(page_h)
    section.top_margin = Cm(MARGIN_CM)
    section.bottom_margin = Cm(MARGIN_CM)
    section.left_margin = Cm(MARGIN_CM)
    section.right_margin = Cm(MARGIN_CM)

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10)

    for i, e in enumerate(entries):
        w_px, h_px = e["width"], e["height"]
        w_cm = content_w
        h_cm = content_w * h_px / w_px
        if h_cm > max_h:  # 太高就让高度贴合可用区域
            h_cm = max_h
            w_cm = max_h * w_px / h_px

        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.page_break_before = i > 0  # 第一页不加分页
        para.add_run().add_picture(str(e["path"]), width=Cm(w_cm))

        if captions:
            cap = doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_before = Pt(3)
            cap.paragraph_format.space_after = Pt(0)
            run = cap.add_run(
                f"#{i + 1:02d}   {e['time']}   ·   {w_px} × {h_px}"
            )
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    props = doc.core_properties
    props.title = "区域截图"
    props.comments = f"共 {len(entries)} 张，每张一页"
    doc.save(str(out_path))


# --------------------------------------------------------------------------
# 全局热键
# --------------------------------------------------------------------------
class HotkeyManager(threading.Thread):
    """在独立线程里注册全局热键，命中后投递到队列。"""

    def __init__(self, out_queue: "queue.Queue[int]"):
        super().__init__(daemon=True)
        self.queue = out_queue
        self.resolved: dict[int, str] = {}   # id -> 实际生效的按键名
        self.failed: list[str] = []          # 连备用键都注册不上的动作名
        self.ready = threading.Event()

    def run(self) -> None:
        u = ctypes.windll.user32
        for hk_id, action, options in HOTKEY_SPECS:
            for mods, vk, label in options:
                if u.RegisterHotKey(None, hk_id, mods | MOD_NOREPEAT, vk):
                    self.resolved[hk_id] = label
                    break
            else:
                self.failed.append(action)
        self.ready.set()

        msg = wintypes.MSG()
        while True:
            ret = u.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret <= 0:
                break
            if msg.message == WM_HOTKEY:
                self.queue.put(int(msg.wParam))

    @staticmethod
    def shutdown() -> None:
        u = ctypes.windll.user32
        for hk_id, _action, _options in HOTKEY_SPECS:
            try:
                u.UnregisterHotKey(None, hk_id)
            except Exception:
                pass


# --------------------------------------------------------------------------
# 主程序
# --------------------------------------------------------------------------
def run_app() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    enable_dpi_awareness()

    root = tk.Tk()
    root.title(f"{APP_TITLE} · {APP_BUILD}")
    root.minsize(520, 700)

    dpi = system_dpi()
    try:
        root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass

    def center_window(win, w: int, h: int) -> None:
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        win.geometry(f"{w}x{h}+{x}+{y}")

    hotkey_queue: "queue.Queue[int]" = queue.Queue()
    hotkeys = HotkeyManager(hotkey_queue)
    hotkeys.start()

    state = {
        "region": None,          # (x, y, w, h)
        "shots": [],             # [{path,width,height,time}]
        "session_dir": None,     # Path
        "auto_job": None,
        "busy": False,
    }

    # ---------------- 变量 ----------------
    var_region = tk.StringVar(value="未选择区域")
    var_delay = tk.IntVar(value=200)
    var_interval = tk.DoubleVar(value=3.0)
    var_auto = tk.BooleanVar(value=False)
    var_captions = tk.BooleanVar(value=True)
    var_copy_images = tk.BooleanVar(value=False)
    var_orientation = tk.StringVar(value="自动")
    var_status = tk.StringVar(value="就绪：按 F7 框选区域")
    var_count = tk.StringVar(value="已截 0 张")
    var_hotkeys = tk.StringVar(value="快捷键：注册中…")

    # ---------------- 设置读写 ----------------
    def load_settings() -> None:
        try:
            data = json.loads(SETTINGS_FILE.read_text("utf-8"))
        except Exception:
            return
        r = data.get("region")
        if isinstance(r, list) and len(r) == 4:
            rect = tuple(int(v) for v in r)
            if region_on_screen(rect):
                state["region"] = rect
                refresh_region_label()
        var_delay.set(int(data.get("delay_ms", 200)))
        var_interval.set(float(data.get("interval_s", 3.0)))
        var_captions.set(bool(data.get("captions", True)))
        var_copy_images.set(bool(data.get("copy_images", False)))
        if data.get("orientation") in ("自动", "纵向", "横向"):
            var_orientation.set(data["orientation"])

    def save_settings() -> None:
        try:
            SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
            SETTINGS_FILE.write_text(
                json.dumps(
                    {
                        "region": list(state["region"]) if state["region"] else None,
                        "delay_ms": int(var_delay.get()),
                        "interval_s": float(var_interval.get()),
                        "captions": bool(var_captions.get()),
                        "copy_images": bool(var_copy_images.get()),
                        "orientation": var_orientation.get(),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                "utf-8",
            )
        except Exception:
            pass

    # ---------------- 界面 ----------------
    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    box_region = ttk.LabelFrame(frm, text=" 1. 区域 ", padding=10)
    box_region.pack(fill="x")
    ttk.Label(box_region, textvariable=var_region).pack(side="left", fill="x", expand=True)
    ttk.Button(box_region, text="框选区域 (F7)", command=lambda: choose_region()).pack(side="right")

    box_shot = ttk.LabelFrame(frm, text=" 2. 截图 ", padding=10)
    box_shot.pack(fill="x", pady=(8, 0))

    row = ttk.Frame(box_shot)
    row.pack(fill="x")
    ttk.Button(row, text="立即截图 (F8)", command=lambda: capture()).pack(side="left")
    ttk.Label(row, textvariable=var_count).pack(side="right")

    row2 = ttk.Frame(box_shot)
    row2.pack(fill="x", pady=(8, 0))
    ttk.Label(row2, text="截图前延迟(ms)").pack(side="left")
    ttk.Spinbox(row2, from_=0, to=3000, increment=50, width=6, textvariable=var_delay).pack(side="left", padx=(6, 16))
    ttk.Checkbutton(row2, text="自动连拍，每", variable=var_auto, command=lambda: toggle_auto()).pack(side="left")
    ttk.Spinbox(row2, from_=0.5, to=600, increment=0.5, width=6, textvariable=var_interval).pack(side="left", padx=6)
    ttk.Label(row2, text="秒").pack(side="left")

    box_out = ttk.LabelFrame(frm, text=" 3. 输出 ", padding=10)
    box_out.pack(fill="x", pady=(8, 0))
    line = ttk.Frame(box_out)
    line.pack(fill="x")
    ttk.Checkbutton(line, text="每页加时间戳说明", variable=var_captions).pack(side="left")
    ttk.Checkbutton(line, text="另存图片副本", variable=var_copy_images).pack(side="left", padx=12)
    ttk.Label(line, text="纸张").pack(side="left")
    ttk.Combobox(
        line, textvariable=var_orientation, values=("自动", "纵向", "横向"),
        width=5, state="readonly",
    ).pack(side="left", padx=6)

    ttk.Button(frm, text="结束并生成 Word (F9)", command=lambda: finalize()).pack(fill="x", pady=(10, 0), ipady=6)

    box_list = ttk.LabelFrame(frm, text=" 已截取的页面 ", padding=6)
    box_list.pack(fill="both", expand=True, pady=(8, 0))
    list_wrap = ttk.Frame(box_list)
    list_wrap.pack(fill="both", expand=True)
    lst = tk.Listbox(list_wrap, height=8, activestyle="none", font=("Consolas", 10))
    sb = ttk.Scrollbar(list_wrap, orient="vertical", command=lst.yview)
    lst.configure(yscrollcommand=sb.set)
    lst.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    line3 = ttk.Frame(box_list)
    line3.pack(fill="x", pady=(6, 0))
    ttk.Button(line3, text="撤销上一张", command=lambda: undo()).pack(side="left")
    ttk.Button(line3, text="全部清空", command=lambda: clear_all()).pack(side="left", padx=8)

    ttk.Label(frm, textvariable=var_hotkeys, foreground="#0a7a3d", wraplength=520).pack(fill="x", pady=(6, 0))
    ttk.Label(frm, textvariable=var_status, foreground="#666").pack(fill="x", pady=(4, 0))
    ttk.Label(
        frm,
        text=f"运行环境：{Path(sys.executable).name} · {APP_BUILD}",
        foreground="#999",
    ).pack(fill="x")

    # ---------------- 辅助 ----------------
    def refresh_region_label() -> None:
        r = state["region"]
        if not r:
            var_region.set("未选择区域")
        else:
            x, y, w, h = r
            var_region.set(f"区域：{w} × {h}  @ ({x}, {y})")

    def set_status(text: str) -> None:
        var_status.set(text)

    def refresh_count() -> None:
        var_count.set(f"已截 {len(state['shots'])} 张")

    def ensure_session() -> Path:
        if state["session_dir"] is None:
            base = Path(tempfile.gettempdir()) / APP_DIR_NAME / datetime.now().strftime("%Y%m%d_%H%M%S")
            base.mkdir(parents=True, exist_ok=True)
            state["session_dir"] = base
        return state["session_dir"]

    def window_covers_region() -> bool:
        try:
            if not root.winfo_viewable() or root.state() == "iconic":
                return False
            wx, wy = root.winfo_rootx(), root.winfo_rooty()
            ww, wh = root.winfo_width(), root.winfo_height()
        except Exception:
            return False
        x, y, w, h = state["region"]
        return not (wx + ww <= x or wx >= x + w or wy + wh <= y or wy >= y + h)

    # ---------------- 动作 ----------------
    def choose_region() -> None:
        if state["auto_job"]:
            stop_auto()
        root.withdraw()
        root.update_idletasks()
        time.sleep(0.08)
        try:
            rect = select_region_on_screen(root)
        finally:
            root.deiconify()
        if rect:
            state["region"] = rect
            refresh_region_label()
            save_settings()
            set_status("区域已设置：按 F8 截图，按 F9 结束")
        else:
            set_status("已取消框选")

    def capture() -> None:
        if state["busy"]:
            return
        if not state["region"]:
            messagebox.showinfo(APP_TITLE, "请先按 F7 框选一块区域。", parent=root)
            return
        if not region_on_screen(state["region"]):
            messagebox.showwarning(APP_TITLE, "区域不在当前屏幕范围内，请重新框选。", parent=root)
            return

        state["busy"] = True
        hidden = False
        try:
            if window_covers_region():
                root.withdraw()
                hidden = True
                root.update_idletasks()
                time.sleep(0.06)
            img = grab_region(state["region"], int(var_delay.get()))
        except Exception as exc:
            log_line(f"截图失败 exe={sys.executable} err={exc!r}")
            messagebox.showerror(
                APP_TITLE,
                f"截图失败：{exc}\n\n"
                f"当前 Python：\n{sys.executable}\n\n"
                f"（详细信息已写入同目录的「运行日志.txt」）",
                parent=root,
            )
            set_status("截图失败")
            return
        finally:
            if hidden:
                root.deiconify()
            state["busy"] = False

        now = datetime.now()
        folder = ensure_session()
        path = folder / f"{len(state['shots']) + 1:03d}_{now:%H%M%S}.png"
        try:
            img.save(path, "PNG")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"保存图片失败：{exc}", parent=root)
            return

        entry = {
            "path": path,
            "width": img.width,
            "height": img.height,
            "time": now.strftime("%Y-%m-%d %H:%M:%S"),
        }
        state["shots"].append(entry)
        lst.insert("end", f"#{len(state['shots']):02d}  {now:%H:%M:%S}   {img.width} × {img.height}")
        log_line(f"截图 OK #{len(state['shots'])} {img.width}x{img.height} -> {path.name}")
        lst.see("end")
        refresh_count()
        set_status(f"已截取第 {len(state['shots'])} 张")

    def undo() -> None:
        if not state["shots"]:
            return
        last = state["shots"].pop()
        try:
            Path(last["path"]).unlink(missing_ok=True)
        except Exception:
            pass
        if lst.size():
            lst.delete("end")
        refresh_count()
        set_status("已撤销上一张")

    def clear_all() -> None:
        if not state["shots"]:
            return
        if not messagebox.askyesno(APP_TITLE, "清空本次已截取的所有图片？", parent=root):
            return
        for item in state["shots"]:
            try:
                Path(item["path"]).unlink(missing_ok=True)
            except Exception:
                pass
        state["shots"].clear()
        lst.delete(0, "end")
        refresh_count()
        set_status("已清空")

    def toggle_auto() -> None:
        if var_auto.get():
            start_auto()
        else:
            stop_auto()

    def start_auto() -> None:
        if not state["region"]:
            var_auto.set(False)
            messagebox.showinfo(APP_TITLE, "请先按 F7 框选一块区域。", parent=root)
            return
        set_status("自动连拍中……再按一次取消")
        tick()

    def tick() -> None:
        if not var_auto.get():
            return
        capture()
        try:
            ms = max(300, int(float(var_interval.get()) * 1000))
        except Exception:
            ms = 3000
        state["auto_job"] = root.after(ms, tick)

    def stop_auto() -> None:
        if state["auto_job"]:
            try:
                root.after_cancel(state["auto_job"])
            except Exception:
                pass
            state["auto_job"] = None
        var_auto.set(False)
        set_status("已停止自动连拍")

    def finalize() -> None:
        stop_auto()
        if not state["shots"]:
            messagebox.showinfo(APP_TITLE, "还没有截图。按 F8 先截几张。", parent=root)
            return

        desktop = Path.home() / "Desktop"
        initial_dir = desktop if desktop.is_dir() else Path.home()
        default_name = f"区域截图_{datetime.now():%Y%m%d_%H%M%S}.docx"
        out = filedialog.asksaveasfilename(
            parent=root,
            title="保存 Word 文档",
            defaultextension=".docx",
            initialfile=default_name,
            initialdir=str(initial_dir),
            filetypes=[("Word 文档", "*.docx")],
        )
        if not out:
            set_status("已取消保存")
            return

        orientation = {"自动": "auto", "纵向": "portrait", "横向": "landscape"}[var_orientation.get()]
        try:
            build_docx(
                out,
                state["shots"],
                orientation=orientation,
                captions=bool(var_captions.get()),
            )
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"生成 Word 失败：{exc}\n\n{traceback.format_exc()}", parent=root)
            return

        if var_copy_images.get():
            try:
                img_dir = Path(out).with_suffix("").with_name(Path(out).stem + "_图片")
                img_dir.mkdir(parents=True, exist_ok=True)
                for item in state["shots"]:
                    shutil.copy2(item["path"], img_dir / Path(item["path"]).name)
            except Exception:
                pass

        set_status(f"已生成：{out}")
        try:
            os.startfile(out)  # noqa: S606 - 打开给用户看
        except Exception:
            pass
        reset_session()

    def reset_session() -> None:
        for item in state["shots"]:
            try:
                Path(item["path"]).unlink(missing_ok=True)
            except Exception:
                pass
        state["shots"].clear()
        lst.delete(0, "end")
        if state["session_dir"]:
            try:
                shutil.rmtree(state["session_dir"], ignore_errors=True)
            except Exception:
                pass
            state["session_dir"] = None
        refresh_count()
        set_status("就绪：按 F7 框选区域，F8 截图，F9 结束")

    # ---------------- 热键分发 ----------------
    def on_hotkey(hk: int) -> None:
        if hk == HK_SELECT:
            choose_region()
        elif hk == HK_CAPTURE:
            capture()
        elif hk == HK_FINISH:
            finalize()

    def poll_hotkeys() -> None:
        try:
            while True:
                on_hotkey(hotkey_queue.get_nowait())
        except queue.Empty:
            pass
        except Exception:
            traceback.print_exc()
        root.after(80, poll_hotkeys)

    def on_close() -> None:
        stop_auto()
        save_settings()
        HotkeyManager.shutdown()
        root.destroy()

    def apply_hotkey_status() -> None:
        if not hotkeys.ready.is_set():
            root.after(120, apply_hotkey_status)
            return
        parts = []
        fallback = []
        for hk_id, action, options in HOTKEY_SPECS:
            label = hotkeys.resolved.get(hk_id)
            parts.append(f"{label} {action}" if label else f"{action} 不可用")
            if label and label != options[0][2]:
                fallback.append(f"{action} → {label}")

        var_hotkeys.set("快捷键：" + "　·　".join(parts))
        if hotkeys.failed:
            set_status("部分快捷键不可用，可直接用界面按钮")
            if not os.environ.get("RSD_SMOKE"):
                messagebox.showwarning(
                    APP_TITLE,
                    "这些动作的快捷键都被其他程序占用了：\n  " + "、".join(hotkeys.failed) +
                    "\n\n如果之前打开过本程序，请先关掉那个旧窗口再重开。\n现在也可以直接用界面上的按钮操作。",
                    parent=root,
                )
        elif fallback and not os.environ.get("RSD_SMOKE"):
            set_status("F7/F8/F9 被占用，已自动改用备用键")
            messagebox.showinfo(
                APP_TITLE,
                "F7 / F8 / F9 已被其他程序占用，已自动改用备用键：\n  " +
                "\n  ".join(fallback) +
                "\n\n如果屏幕上还开着一个旧的本程序窗口，请关掉它，然后再重开本程序。",
                parent=root,
            )
        elif state["region"]:
            set_status("已恢复上次的区域，可以开始截图了")

    root.protocol("WM_DELETE_WINDOW", on_close)
    load_settings()
    refresh_count()
    root.after(120, poll_hotkeys)
    root.after(200, apply_hotkey_status)

    if os.environ.get("RSD_SMOKE"):  # 自检：不显示窗口，2 秒后自动退出
        root.withdraw()
        root.update_idletasks()

        def _smoke_done():
            lines = [
                "SMOKE OK hotkey_failures=" + repr(hotkeys.failed),
                "hotkeys=" + repr(hotkeys.resolved),
                f"python={sys.executable}",
                f"ui req size: {frm.winfo_reqwidth()} x {frm.winfo_reqheight()}",
            ]
            for name, widget in (("region", box_region), ("shot", box_shot), ("out", box_out), ("list", box_list)):
                lines.append(f"  {name}: {widget.winfo_reqwidth()} x {widget.winfo_reqheight()}")
            text = "\n".join(lines)
            print(text, flush=True)
            marker = os.environ.get("RSD_SMOKE_OUT")
            if marker:
                try:
                    Path(marker).write_text(text, "utf-8")
                except Exception:
                    pass
            on_close()
        root.after(800, _smoke_done)

    center_window(root, 580, 820)

    def bring_to_front() -> None:
        try:
            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            root.after(900, lambda: root.attributes("-topmost", False))
            root.focus_force()
        except Exception:
            pass
        log_line("UI 就绪，窗口已显示")

    root.after(150, bring_to_front)
    log_line("窗口已创建，进入事件循环")

    root.mainloop()


def select_region_on_screen(root) -> tuple[int, int, int, int] | None:
    """全屏半透明遮罩，拖拽选择矩形，返回屏幕坐标 (x, y, w, h)。"""
    import tkinter as tk

    vx, vy, vw, vh = virtual_screen()
    top = tk.Toplevel(root)
    top.overrideredirect(True)
    top.attributes("-topmost", True)
    top.attributes("-alpha", 0.30)
    top.geometry(f"{vw}x{vh}+{vx}+{vy}")
    top.configure(bg="black")

    canvas = tk.Canvas(top, bg="black", highlightthickness=0, cursor="crosshair")
    canvas.pack(fill="both", expand=True)

    result: dict[str, tuple[int, int, int, int] | None] = {"rect": None}
    start = {"x": 0, "y": 0}
    rect_id = {"id": None}
    size_id = {"id": None}

    hint = canvas.create_text(
        vw // 2, 40,
        text="拖动鼠标框选区域　·　Esc 取消",
        fill="#ffffff", font=("Microsoft YaHei", 16), tags="instr",
    )
    canvas.create_rectangle(
        vw // 2 - 230, 18, vw // 2 + 230, 62,
        outline="#00d0ff", width=1, tags="instr",
    )

    def on_press(event) -> None:
        start["x"], start["y"] = event.x, event.y
        canvas.delete("instr")
        if rect_id["id"] is None:
            rect_id["id"] = canvas.create_rectangle(
                event.x, event.y, event.x, event.y,
                outline="#00d0ff", width=2,
            )

    def on_drag(event) -> None:
        if rect_id["id"] is None:
            return
        canvas.coords(rect_id["id"], start["x"], start["y"], event.x, event.y)
        w, h = abs(event.x - start["x"]), abs(event.y - start["y"])
        canvas.coords(hint, min(start["x"], event.x) + 8, min(start["y"], event.y) - 12)
        canvas.itemconfigure(hint, text=f"{w} × {h}", font=("Consolas", 12), fill="#00d0ff")

    def on_release(event) -> None:
        x0, y0 = start["x"], start["y"]
        x1, y1 = event.x, event.y
        left, top_ = min(x0, x1), min(y0, y1)
        w, h = abs(x1 - x0), abs(y1 - y0)
        if w < 8 or h < 8:
            close(None)
            return
        result["rect"] = (vx + left, vy + top_, w, h)
        close(None)

    def close(_event) -> None:
        try:
            top.grab_release()
        except Exception:
            pass
        top.destroy()

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    top.bind("<Escape>", close)

    top.grab_set()
    top.focus_force()
    canvas.focus_set()
    root.wait_window(top)
    return result["rect"]


# --------------------------------------------------------------------------
# 依赖自举：双击 .pyw 时可能落到没装 Pillow 的 Python 上
# --------------------------------------------------------------------------
REQUIRED_MODULES = (("PIL", "Pillow"), ("docx", "python-docx"))
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
LOG_FILE = Path(__file__).with_name("运行日志.txt")


def log_line(text: str) -> None:
    """把运行环境写进同目录的日志，出问题时可定位。"""
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {text}\n")
    except Exception:
        pass


def missing_dependencies() -> list[str]:
    missing = []
    for module, package in REQUIRED_MODULES:
        try:
            __import__(module)
        except Exception:
            missing.append(package)
    return missing


def candidate_interpreters():
    """依次产出 (探测用 exe, 启动用 exe)。"""
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    base = home / ".dsh" / "dsh-runtimes" / "dsh-primary-runtime" / "dependencies" / "python"
    py, pyw = base / "python.exe", base / "pythonw.exe"
    if py.is_file():
        yield str(py), str(pyw if pyw.is_file() else py)
    for name in ("python.exe", "python3.exe"):
        found = shutil.which(name)
        if found:
            guess = Path(found).with_name("pythonw.exe")
            yield found, str(guess if guess.is_file() else found)


def interpreter_has_deps(exe: str) -> bool:
    if os.path.normcase(os.path.abspath(exe)) == os.path.normcase(os.path.abspath(sys.executable)):
        return True
    try:
        result = subprocess.run(
            [exe, "-c", "import PIL, docx"],
            capture_output=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        )
        return result.returncode == 0
    except Exception:
        return False


def show_dependency_error(missing: list[str]) -> None:
    message = (
        "缺少依赖：" + "、".join(missing) + "\n\n"
        "当前使用的 Python：\n" + sys.executable + "\n\n"
        "解决办法（任选其一）：\n\n"
        "1）关闭这个窗口，改双击同目录下的「启动.bat」\n"
        "    它会自动挑一个已经装好依赖的 Python。\n\n"
        "2）给当前 Python 装上依赖：\n"
        f'    "{sys.executable}" -m pip install pillow python-docx\n'
    )
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, message)
        root.destroy()
    except Exception:
        print(message)


def bootstrap() -> bool:
    """依赖齐全返回 True；否则尝试换一个有依赖的 Python 重开自己。"""
    log_line(
        f"启动 build={APP_BUILD} exe={sys.executable} "
        f"relaunched={os.environ.get('RSD_RELAUNCHED', '0')}"
    )
    missing = missing_dependencies()
    log_line(f"依赖检查 missing={missing or '无'}")
    if not missing:
        return True

    if not os.environ.get("RSD_RELAUNCHED"):
        me = os.path.normcase(os.path.abspath(sys.executable))
        for probe, launch in candidate_interpreters():
            if os.path.normcase(os.path.abspath(probe)) == me:
                continue
            ok = interpreter_has_deps(probe)
            log_line(f"候选 {probe} 可用={ok}")
            if not ok:
                continue
            env = os.environ.copy()
            env["RSD_RELAUNCHED"] = "1"
            try:
                subprocess.Popen(
                    [launch, os.path.abspath(__file__)],
                    env=env,
                    cwd=str(Path(__file__).resolve().parent),
                    close_fds=True,
                )
                log_line(f"已用 {launch} 重新启动")
                return False
            except Exception as exc:
                log_line(f"重启失败 {launch}: {exc}")
                continue

    log_line("没有找到可用的 Python 环境，显示错误提示")
    show_dependency_error(missing)
    return False


def install_crash_logging() -> None:
    """pythonw 没有控制台；崩溃时把现场写进日志，避免“闪一下就没了”。"""
    import faulthandler

    try:
        handle = open(LOG_FILE, "a", encoding="utf-8")
        faulthandler.enable(handle)
    except Exception:
        pass

    def _hook(exc_type, exc, tb):
        log_line("未捕获异常：\n" + "".join(traceback.format_exception(exc_type, exc, tb)))

    sys.excepthook = _hook
    try:
        threading.excepthook = lambda args: log_line(
            "线程异常：\n" + "".join(traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback))
        )
    except Exception:
        pass


if __name__ == "__main__":
    install_crash_logging()
    try:
        log_line(f"===== 进程启动 pid={os.getpid()} =====")
        if bootstrap():
            run_app()
        log_line("主循环结束，进程退出")
    except Exception:
        log_line("启动失败：\n" + traceback.format_exc())
        try:
            import tkinter.messagebox as mb

            mb.showerror("启动失败", traceback.format_exc())
        except Exception:
            pass
        sys.exit(1)
