# -*- coding: utf-8 -*-
"""清掉任意卡住的 VBA 错误弹窗（自动找 Excel 进程 + #32770 对话框）。

用法：python dismiss_vba_dialog.py [excel_pid]
      不给 pid 时自动找所有 EXCEL.EXE 里的 VBA 错误框。
关法：优先 BM_CLICK“结束”按钮(id 4800)，无效再 WM_COMMAND，再无效真实鼠标点击，最后才杀进程。
"""
import sys
import time

import psutil
import win32api
import win32con
import win32gui
import win32process

END_ID, STATIC_ID = 4800, 4803


def excel_pids():
    out = []
    for p in psutil.process_iter(["pid", "name"]):
        if (p.info.get("name") or "").upper() == "EXCEL.EXE":
            out.append(p.info["pid"])
    return out


def find_dialogs(pids):
    hits = []

    def cb(h, _):
        try:
            _, wp = win32process.GetWindowThreadProcessId(h)
            if wp not in pids or not win32gui.IsWindowVisible(h):
                return True
            if win32gui.GetClassName(h) == "#32770" and "Visual Basic" in win32gui.GetWindowText(h):
                hits.append((wp, h))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return hits


def kids(h):
    out = []

    def cb(c, _):
        out.append({"hwnd": c, "id": win32gui.GetDlgCtrlID(c),
                    "class": win32gui.GetClassName(c), "text": win32gui.GetWindowText(c)})
        return True

    win32gui.EnumChildWindows(h, cb, None)
    return out


def gone(h, sec=4.0):
    t0 = time.time()
    while time.time() - t0 < sec:
        if not win32gui.IsWindow(h):
            return True
        time.sleep(0.1)
    return not win32gui.IsWindow(h)


targets = [int(sys.argv[1])] if len(sys.argv) > 1 else excel_pids()
dialogs = find_dialogs(targets) if targets else []
if not dialogs:
    print("没有发现卡住的 VBA 弹窗")
    sys.exit(0)

for pid, dlg in dialogs:
    ch = kids(dlg)
    err = next((c["text"] for c in ch if c["id"] == STATIC_ID), "")
    print("pid=%s 弹窗=%r 错误文本=%r" % (pid, win32gui.GetWindowText(dlg), err))
    btn = next((c for c in ch if c["id"] == END_ID), None)
    if btn:
        win32gui.PostMessage(btn["hwnd"], win32con.BM_CLICK, 0, 0)
        if gone(dlg):
            print("  -> BM_CLICK 关闭成功 ✓（进程存活）")
            continue
    if btn:
        win32gui.PostMessage(dlg, win32con.WM_COMMAND, btn["id"], btn["hwnd"])
        if gone(dlg, 3):
            print("  -> WM_COMMAND 关闭成功 ✓")
            continue
        l, t, r, b = win32gui.GetWindowRect(btn)
        win32api.SetCursorPos(((l + r) // 2, (t + b) // 2))
        time.sleep(0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        if gone(dlg, 3):
            print("  -> 鼠标点击 关闭成功 ✓")
            continue
    print("  -> 关不掉，杀进程 pid=%s" % pid)
    psutil.Process(pid).kill()
