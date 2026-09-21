#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vba-kit MCP server —— 零依赖（只用 Python 标准库），stdio + 逐行 JSON-RPC 2.0。

暴露三个工具给任何 MCP 客户端：
  excel_status    只读：列出运行中的 Excel 实例，以及是否有卡住的弹窗
  run_vba         跑 run_vba.py：注入 → 守卫 → 运行 → 强制重算 → 断言 → 报告
  dismiss_dialog  清理卡住的 VBA 报错弹窗（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程）

启动： python tools/vba/mcp_server.py
说明：本 server 自身零依赖；run_vba 等脚本需要 pywin32 + psutil + 本机 Excel。
"""

import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PROTOCOL_FALLBACK = "2025-06-18"


def _run_tool_script(args, timeout=900):
    p = subprocess.run([PY] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=timeout)
    txt = (p.stdout or "").strip()
    if (p.stderr or "").strip():
        txt += "\n[stderr]\n" + p.stderr.strip()
    return p.returncode, txt


# ---------------- 工具实现 ----------------

def tool_excel_status(_args):
    """只读：Excel 实例 + 卡住的弹窗。"""
    try:
        import psutil
        import win32gui
        import win32process
    except ImportError as e:
        return "缺少依赖：%r —— 请先 pip install -r tools/vba/requirements.txt" % (e,)
    lines = []
    for p in psutil.process_iter(["pid", "name"]):
        try:
            if (p.info.get("name") or "").upper() != "EXCEL.EXE":
                continue
            pid = p.info["pid"]
            wins = []

            def cb(h, _):
                try:
                    _, wp = win32process.GetWindowThreadProcessId(h)
                    if wp == pid and win32gui.IsWindowVisible(h):
                        cls, title = win32gui.GetClassName(h), win32gui.GetWindowText(h)
                        if cls == "#32770" or "Visual Basic" in title:
                            wins.append("%s / %r" % (cls, title))
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(cb, None)
            lines.append("pid=%s 卡住的弹窗: %s" % (pid, "；".join(wins) if wins else "无"))
        except Exception:
            continue
    return "\n".join(lines) if lines else "当前没有运行中的 Excel 实例"


def tool_run_vba(args):
    """跑 run_vba.py。必填 workbook / code / macro。"""
    cmd = [os.path.join(HERE, "run_vba.py")]
    wb = args.get("workbook")
    if not wb:
        return "缺少参数 workbook"
    cmd += ["--workbook", str(wb)]
    for c in args.get("code") or []:
        cmd += ["--code", str(c)]
    macro = args.get("macro")
    if not macro:
        return "缺少参数 macro（要运行的宏名）"
    cmd += ["--run", str(macro)]
    for e in args.get("expect") or []:
        cmd += ["--expect", str(e)]
    if args.get("keep_open"):
        cmd.append("--keep-open")
    if args.get("save"):
        cmd.append("--save")
    if args.get("visible"):
        cmd.append("--visible")
    if args.get("hard_budget"):
        cmd += ["--hard-budget", str(args["hard_budget"])]
    if args.get("allow_unsafe"):
        cmd.append("--allow-unsafe")
    try:
        rc, out = _run_tool_script(cmd)
    except subprocess.TimeoutExpired:
        return "run_vba 超时（>900s），可能有未处理的卡死"
    return "exit=%s\n%s" % (rc, out)


def tool_dismiss_dialog(args):
    """清理卡住的 VBA 报错弹窗。可选 excel_pid。"""
    cmd = [os.path.join(HERE, "dismiss_vba_dialog.py")]
    if args.get("excel_pid"):
        cmd.append(str(args["excel_pid"]))
    try:
        rc, out = _run_tool_script(cmd, timeout=120)
    except subprocess.TimeoutExpired:
        return "dismiss_vba_dialog 超时"
    return "exit=%s\n%s" % (rc, out)


TOOLS = [
    {
        "name": "excel_status",
        "description": "只读：列出运行中的 Excel 实例，以及是否存在卡住的 VBA 报错弹窗（标题/类名）。跑宏前后想知道环境状态时用。",
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {"readOnlyHint": True, "title": "Excel 状态"},
    },
    {
        "name": "run_vba",
        "description": ("跑 Excel VBA 宏的开发测试跑道：自动剥 Attribute 行、拦截危险语句、起守卫"
                        "（弹窗点掉/假死/跑飞判定）、运行、强制重算、跑断言、返回报告。"
                        "开发或修改宏一律用它，不要手写 win32com 注入。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径（不存在则新建）"},
                "code": {"type": "array", "items": {"type": "string"},
                         "description": "代码文件，格式 模块名=路径 或 直接路径（模块名取文件名）"},
                "macro": {"type": "string", "description": "要运行的宏名（如 BuildReport）"},
                "expect": {"type": "array", "items": {"type": "string"},
                           "description": ("断言，如 cell:汇总!B2=623685.6 / sheet:汇总=存在 / "
                                           "pivot:汇总=1 / chart:汇总=1 / named:收入数据=存在")},
                "keep_open": {"type": "boolean", "description": "复用/保留 Excel 实例（迭代快，默认 false）"},
                "save": {"type": "boolean", "description": "跑完保存工作簿"},
                "visible": {"type": "boolean", "description": "让 Excel 可见（默认隐藏）"},
                "hard_budget": {"type": "number", "description": "跑飞判定的硬预算秒数（默认 45）"},
                "allow_unsafe": {"type": "boolean", "description": "放行危险语句（不推荐，很可能挂死）"},
            },
            "required": ["workbook", "code", "macro"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": False, "title": "跑 VBA 宏并测试"},
    },
    {
        "name": "dismiss_dialog",
        "description": "清理卡住的 VBA 报错弹窗（自动找 Excel 进程；依次尝试点“结束”按钮、WM_COMMAND、真实鼠标点击，最后才杀进程）。",
        "inputSchema": {
            "type": "object",
            "properties": {"excel_pid": {"type": "integer", "description": "可选，指定 Excel 进程号"}},
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": False, "title": "清掉卡住的弹窗"},
    },
]

HANDLERS = {"excel_status": tool_excel_status, "run_vba": tool_run_vba, "dismiss_dialog": tool_dismiss_dialog}


# ---------------- JSON-RPC 循环 ----------------

def send(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def handle(msg):
    method = msg.get("method")
    mid = msg.get("id")
    if method == "initialize":
        client_pv = (msg.get("params") or {}).get("protocolVersion") or PROTOCOL_FALLBACK
        send({"jsonrpc": "2.0", "id": mid, "result": {
            "protocolVersion": client_pv,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "vba-kit", "version": "1.0.0"},
        }})
    elif method in ("notifications/initialized", "notifications/cancelled"):
        return                       # 通知无需回复
    elif method == "ping":
        send({"jsonrpc": "2.0", "id": mid, "result": {}})
    elif method == "tools/list":
        send({"jsonrpc": "2.0", "id": mid, "result": {"tools": TOOLS}})
    elif method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name")
        args = params.get("arguments") or {}
        fn = HANDLERS.get(name)
        if fn is None:
            send({"jsonrpc": "2.0", "id": mid, "error": {"code": -32602, "message": "未知工具: %s" % name}})
            return
        try:
            text = fn(args)
            send({"jsonrpc": "2.0", "id": mid, "result": {
                "content": [{"type": "text", "text": str(text)}], "isError": False}})
        except Exception as e:
            send({"jsonrpc": "2.0", "id": mid, "result": {
                "content": [{"type": "text", "text": "工具执行异常: %r" % (e,)}], "isError": True}})
    elif mid is not None:
        send({"jsonrpc": "2.0", "id": mid, "error": {"code": -32601, "message": "未实现的方法: %s" % method}})


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue
        try:
            handle(msg)
        except Exception as e:
            if msg.get("id") is not None:
                send({"jsonrpc": "2.0", "id": msg["id"],
                      "error": {"code": -32603, "message": "内部错误: %r" % (e,)}})


if __name__ == "__main__":
    try:
        main()
    except (BrokenPipeError, KeyboardInterrupt):
        pass
