# -*- coding: utf-8 -*-
"""VBA 宏执行守卫（统一版）：弹窗处理 + 无弹窗卡死判定 + 心跳 + 分级处置。

判定逻辑（每 interval 秒一轮，全用绝对墙钟）：
  1) 有弹窗（class #32770 或标题含 Visual Basic）→ 点 id 4800“结束”按钮；点不掉 → kill
  2) 无弹窗时做**活性探测**：SendMessageTimeout(hwnd, WM_NULL, SMTO_ABORTIFHUNG, 300ms)
     - 有响应         → OK（正常在跑），计数清零
     - 无响应 + 心跳文件新鲜      → BUSY_WITH_PROGRESS（在干活，继续等，直到硬预算）
     - 无响应 + CPU 在涨          → BUSY_CPU（可能在算；超硬预算判 RUNAWAY_CPU → kill）
     - 无响应 + CPU 不涨 + 超 hang_after → IDLE_HUNG（等 IO/锁/被挂起）→ kill
  3) 预算耗尽仍未判定异常 → BUDGET_EXCEEDED（不改状态，交由调用方决定）

用法：
  python vba_guard.py <excel_pid> [budget=60] [interval=0.1] [hang_after=6] [hard_budget=45]
                      [heartbeat=<进度文件路径>] [heartbeat_stale=10] [no_kill]
输出行（供日志留痕）：
  HELLO / STATE / DIALOG / ERRTEXT / DISMISSED_BY_CLICK / KILLED_* / VERDICT <判定> elapsed=<s> cpu_delta=<s>
退出码：0=已处置（弹窗或卡死）；3=预算内正常，未触发处置。
依赖：psutil、pywin32。
"""
import json
import os
import sys
import time

import psutil
import win32con
import win32gui
import win32process

END_BTN_ID = 4800          # 出错框“结束(&E)”
STATIC_ID = 4803           # 错误文本
WM_NULL = 0
SMTO_BLOCK, SMTO_ABORTIFHUNG = 0x0001, 0x0002
SMTO_FLAGS = SMTO_BLOCK | SMTO_ABORTIFHUNG


def parse_args(argv):
    pos = [a for a in argv[1:] if not a.startswith("--") and "=" not in a]
    opts = {}
    for a in argv[1:]:
        if a.startswith("--") and "=" in a:
            k, v = a[2:].split("=", 1)
            opts[k] = v
        elif a.startswith("--"):
            opts[a[2:]] = True
    pid = int(pos[0])
    budget = float(pos[1]) if len(pos) > 1 else 60.0
    interval = float(pos[2]) if len(pos) > 2 else 0.1
    return pid, budget, interval, opts


def dialogs(pid):
    hits = []

    def cb(hwnd, _):
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid != pid or not win32gui.IsWindowVisible(hwnd):
                return True
            cls, title = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if cls == "#32770" or "Visual Basic" in title:
                kids = []

                def cb2(c, _):
                    kids.append({"hwnd": c, "id": win32gui.GetDlgCtrlID(c),
                                 "class": win32gui.GetClassName(c),
                                 "text": win32gui.GetWindowText(c)})
                    return True

                win32gui.EnumChildWindows(hwnd, cb2, None)
                hits.append({"hwnd": hwnd, "class": cls, "title": title, "children": kids})
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return hits


def main_hwnd(pid):
    found = []

    def cb(hwnd, _):
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid == pid and win32gui.GetClassName(hwnd) == "XLMAIN":
                found.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return found[0] if found else None


def responsive(hwnd, timeout_ms=300):
    """消息泵探测；超时=无响应（不抛异常的写法）。"""
    try:
        win32gui.SendMessageTimeout(hwnd, WM_NULL, 0, 0, SMTO_FLAGS, timeout_ms)
        return True
    except Exception:
        return False


def cpu_total(p):
    t = p.cpu_times()
    return t.user + t.system


def kill(pid, verdict, elapsed, cpu_delta):
    try:
        psutil.Process(pid).kill()
    except Exception as e:
        print("KILL_FAILED %r" % (e,), flush=True)
        return 1
    print("VERDICT %s elapsed=%.1fs cpu_delta=%.1fs -> KILLED pid=%s" % (verdict, elapsed, cpu_delta, pid), flush=True)
    return 0


def main():
    pid, budget, interval, opts = parse_args(sys.argv)
    hang_after = float(opts.get("hang_after", 6))
    hard_budget = float(opts.get("hard_budget", 45))
    hb_path = opts.get("heartbeat")
    hb_stale = float(opts.get("heartbeat_stale", 10))
    do_kill = "no_kill" not in opts
    try:
        proc = psutil.Process(pid)
    except Exception as e:
        print("HELLO_FAILED %r" % (e,), flush=True)
        return 2
    hwnd = main_hwnd(pid)
    print("HELLO pid=%s hwnd=%s budget=%ss interval=%ss hang_after=%ss hard_budget=%ss heartbeat=%s"
          % (pid, hwnd, budget, interval, hang_after, hard_budget, hb_path or "-"), flush=True)

    t_start = time.time()
    cpu0 = cpu_total(proc)
    samples = [(t_start, cpu0)]
    t_last_state = 0.0
    idle_since = None
    while time.time() - t_start < budget:
        now = time.time()
        elapsed = now - t_start
        # ---- 1) 弹窗优先 ----
        d = dialogs(pid)
        if d:
            dlg = d[0]
            err = next((c["text"] for c in dlg["children"] if c["id"] == STATIC_ID), "")
            print("DIALOG " + json.dumps({"title": dlg["title"], "hwnd": dlg["hwnd"],
                                          "buttons": [(c["id"], c["text"]) for c in dlg["children"]
                                                      if c["class"] == "Button"]}, ensure_ascii=False), flush=True)
            if err:
                print("ERRTEXT " + json.dumps(err, ensure_ascii=False), flush=True)
            btn = next((c for c in dlg["children"] if c["id"] == END_BTN_ID), None)
            if btn:
                win32gui.PostMessage(btn["hwnd"], win32con.BM_CLICK, 0, 0)
                for _ in range(30):
                    time.sleep(0.1)
                    if not [x for x in dialogs(pid) if x["hwnd"] == dlg["hwnd"]]:
                        print("DISMISSED_BY_CLICK elapsed=%.2fs" % (time.time() - t_start), flush=True)
                        return 0
                print("CLICK_NO_EFFECT", flush=True)
            if do_kill:
                return kill(pid, "DIALOG_UNCLOSABLE", elapsed, cpu_total(proc) - cpu0)
            return 0
        # ---- 2) 活性判定 ----
        if hwnd is None:
            hwnd = main_hwnd(pid)
        resp = responsive(hwnd) if hwnd else True
        cpu_now = cpu_total(proc)
        samples.append((now, cpu_now))
        while len(samples) > 1 and now - samples[0][0] > 2.0:
            samples.pop(0)
        cpu_recent = cpu_now - samples[0][1]      # 近 2 秒真实 CPU 占用
        cpu_delta = cpu_now - cpu0
        hb_age = None
        if hb_path and os.path.exists(hb_path):
            hb_age = now - os.path.getmtime(hb_path)
        if resp:
            idle_since = None
        else:
            if idle_since is None:
                idle_since = now
            hung_for = now - idle_since
            if hb_age is not None and hb_age <= hb_stale:
                state = "BUSY_WITH_PROGRESS"
            elif cpu_recent > 0.3:
                state = "BUSY_CPU"
                if hung_for >= hard_budget:
                    if do_kill:
                        return kill(pid, "RUNAWAY_CPU", elapsed, cpu_delta)
            elif hung_for >= hang_after:
                state = "IDLE_HUNG"
                if do_kill:
                    return kill(pid, "IDLE_HUNG", elapsed, cpu_delta)
            else:
                state = "UNRESPONSIVE_WAIT"
            if now - t_last_state > 2.0:
                print("STATE %s elapsed=%.1fs hung_for=%.1fs cpu_recent=%.2fs cpu_total=%.2fs hb_age=%s"
                      % (state, elapsed, hung_for, cpu_recent, cpu_delta,
                         ("%.1fs" % hb_age) if hb_age is not None else "-"), flush=True)
                t_last_state = now
        time.sleep(interval)
    print("VERDICT BUDGET_EXCEEDED elapsed=%.1fs cpu_delta=%.1fs -> 未处置（交给调用方决定）"
          % (time.time() - t_start, cpu_total(proc) - cpu0), flush=True)
    return 3


if __name__ == "__main__":
    sys.exit(main())
