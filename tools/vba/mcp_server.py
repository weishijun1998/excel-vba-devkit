#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vba-kit MCP server —— 零依赖（只用 Python 标准库），stdio + 逐行 JSON-RPC 2.0。

暴露九类工具给任何 MCP 客户端：

  跑宏 / 验证（**唯一入口**，其他任何方式都别用来跑宏）
    excel_status      只读：列出运行中的 Excel 实例，以及是否有卡住的弹窗
    run_vba           跑 run_vba.py：注入 → 守卫 → 运行 → 强制重算 → 断言 → 报告
    dismiss_dialog    清理卡住的 VBA 报错弹窗（BM_CLICK → WM_COMMAND → 真实鼠标 → 杀进程）

  工程级读写（转发 vba_project.py；本 server 只做转发，COM 活在子进程里）
    vba_modules       只读：列模块 / 读模块源码
    vba_patch_module  行级编辑（不会丢 procedure attributes）
    vba_check_compile 真编译检查（注入空探针宏后 Run —— 工程级编译）
    vba_backup        导出/备份模块
    vba_import_module 从 .bas/.cls/.frm 整体替换模块（唯一安全方式，保留 Attribute）
    vba_list_references 列工程引用 / 断引用

职责边界（别越界，否则会踩 SKILL.md 里的真实事故）：
  * 跑宏 → 只用 run_vba（有守卫/断言/Attribute 剥离）；不要用别的 server 的 run。
  * 改代码 → 用 vba_patch_module（行级）；整体替换用 vba_import_module（Import 保留
    procedure attributes）；**永远不要"导出→整模块重写"**。
  * compile 是只读语义（探针跑完即移除，不保存就不落盘）。

启动： python tools/vba/mcp_server.py
说明：本 server 自身零依赖；run_vba / vba_project 等脚本需要 pywin32 + psutil + 本机 Excel。
"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PROJECT = os.path.join(HERE, "vba_project.py")
PROTOCOL_FALLBACK = "2025-06-18"


def _setup_stdio():
    """强制 stdio 为 UTF-8。

    MCP 规范要求 stdio 上的消息是 UTF-8。Windows 上 Python 的 sys.stdin/stdout 默认走
    locale 编码（中文机器 = cp936/GBK），于是本 server 写出 GBK 字节、客户端按 UTF-8
    解码 → 工具描述与中文参数全部乱码（实测：utf-8 解码失败 0xb6，cp936 解码成功）。
    本 kit 的 run_vba.py / vba_diagnose.py / vba_lint_refs.py 均已做同样处理，此处补齐。
    """
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(newline="\n")   # 别让 Windows 把 \n 翻成 \r\n
        except Exception:
            pass


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


# ---------------- 工程级读写（转发 vba_project.py）----------------
# 本 server 只拼命令行 + 转发；真正的 COM 活在子进程里（子进程挂了不会带走 server）。

def _project(argv, timeout=240):
    # 注意：_run_tool_script 自己会在最前面插解释器，这里只给脚本路径
    try:
        rc, out = _run_tool_script([PROJECT] + argv, timeout=timeout)
    except subprocess.TimeoutExpired:
        return ("vba_project 超时（>%ss）——通常是有模态框没被点掉。"
                "先调 excel_status 看环境，必要时 dismiss_dialog。" % timeout)
    return "exit=%s\n%s" % (rc, out)


def _wb(args):
    wb = args.get("workbook")
    return str(wb) if wb else None


def tool_vba_modules(args):
    wb = _wb(args)
    if not wb:
        return "缺少参数 workbook"
    action = (args.get("action") or "list").lower()
    if action == "list":
        return _project(["list", "--workbook", wb], timeout=120)
    if action in ("read_all", "all"):
        return _project(["read", "--workbook", wb, "--all"], timeout=120)
    mod = args.get("module_name")
    if not mod:
        return "action=read 需要 module_name"
    return _project(["read", "--workbook", wb, "--module", str(mod)], timeout=120)


def tool_vba_patch_module(args):
    wb, mod, op = _wb(args), args.get("module_name"), args.get("operation")
    if not wb:
        return "缺少参数 workbook"
    if not mod:
        return "缺少参数 module_name"
    if not op:
        return "缺少参数 operation（replace/insert/delete/append）"
    argv = ["patch", "--workbook", wb, "--module", str(mod), "--op", str(op)]
    if args.get("start_line") is not None:
        argv += ["--start", str(int(args["start_line"]))]
    if args.get("end_line") is not None:
        argv += ["--end", str(int(args["end_line"]))]
    tmp = None
    code = args.get("code")
    if code:
        fd, tmp = tempfile.mkstemp(prefix="vba_kit_patch_", suffix=".txt")
        os.close(fd)
        # 写 UTF-8：vba_project 会自动识别编码，中文不会丢
        with open(tmp, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(str(code))
        argv += ["--code-file", tmp]
    if args.get("save", True):
        argv.append("--save")
    if args.get("no_backup"):
        argv.append("--no-backup")
    try:
        return _project(argv)
    finally:
        if tmp:
            try:
                os.remove(tmp)
            except OSError:
                pass


def tool_vba_check_compile(args):
    wb = _wb(args)
    if not wb:
        return "缺少参数 workbook"
    return _project(["compile", "--workbook", wb, "--dialog-wait", "2"], timeout=300)


def tool_vba_backup(args):
    wb = _wb(args)
    if not wb:
        return "缺少参数 workbook"
    argv = ["backup", "--workbook", wb]
    if args.get("module_name"):
        argv += ["--module", str(args["module_name"])]
    if args.get("out_dir"):
        argv += ["--out-dir", str(args["out_dir"])]
    return _project(argv, timeout=300)


def tool_vba_import_module(args):
    wb, src = _wb(args), args.get("source_file")
    if not wb:
        return "缺少参数 workbook"
    if not src:
        return "缺少参数 source_file"
    argv = ["import", "--workbook", wb, "--source", str(src)]
    if args.get("module_name"):
        argv += ["--module", str(args["module_name"])]
    if args.get("save", True):
        argv.append("--save")
    if args.get("no_backup"):
        argv.append("--no-backup")
    return _project(argv, timeout=300)


def tool_vba_list_references(args):
    wb = _wb(args)
    if not wb:
        return "缺少参数 workbook"
    return _project(["refs", "--workbook", wb], timeout=120)


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
    {
        "name": "vba_modules",
        "description": ("只读：列出 .xlsm 里的 VBA 模块（名/类型/行数），或读某个模块 / 全部模块的源码。"
                        "CodeModule **不含 Attribute 行**（那只存在于导出的 .bas 里）——所以引用的行号"
                        "可以直接喂给 vba_patch_module 的 start_line。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
                "action": {"type": "string", "enum": ["list", "read", "read_all"],
                           "description": "默认 list；read 需要 module_name；read_all 读全部"},
                "module_name": {"type": "string", "description": "action=read 时必填"},
            },
            "required": ["workbook"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": True, "title": "读 VBA 模块"},
    },
    {
        "name": "vba_patch_module",
        "description": ("按行编辑 VBA 模块（replace / insert / delete / append）——**行级编辑**，"
                        "因此不会丢 procedure attributes（如 VB_ProcData 快捷键绑定；整模块重写才会丢）。"
                        "行号 1-based，来自 vba_modules 的 read。代码里若含 Attribute 行会被自动剥离"
                        "（带 Attribute 注入会让整个工程编译失败）。默认改前自动备份、默认落盘保存。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
                "module_name": {"type": "string", "description": "要改的模块名"},
                "operation": {"type": "string", "enum": ["replace", "insert", "delete", "append"],
                              "description": "append 忽略 start_line；insert 的 start_line=0 表示插到开头"},
                "start_line": {"type": "integer", "description": "1-based 起始行"},
                "end_line": {"type": "integer", "description": "1-based 结束行（含）；默认等于 start_line"},
                "code": {"type": "string", "description": "要写入的代码（replace/insert/append 用）"},
                "save": {"type": "boolean", "description": "是否落盘保存（默认 true）"},
                "no_backup": {"type": "boolean", "description": "true 则跳过改前自动备份（不推荐）"},
            },
            "required": ["workbook", "module_name", "operation"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": False, "title": "行级改 VBA 模块"},
    },
    {
        "name": "vba_check_compile",
        "description": ("真编译检查：注入一个空探针宏并 Run 它 —— VBA 的编译是**工程级**的，探针跑通即"
                        "整个 VBAProject 可编译。探针跑完立即移除，不保存就不落盘。"
                        "verdict 取值：CLEAN（可编译）/ COMPILE_ERRORS / BROKEN_REFERENCES / RUN_FAILED。"
                        "注意：工程编译不过时 Excel 报的是『无法运行宏』，与宏名歧义、宏被禁用**同一文案**，"
                        "报告里会写明判定前提。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
            },
            "required": ["workbook"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": True, "title": "编译检查"},
    },
    {
        "name": "vba_backup",
        "description": ("把 VBA 模块导出到磁盘（VBE 的 Export，写的是**系统 ANSI/GBK**）。"
                        "不指定 module_name 则备份全部模块并写 _manifest.json；"
                        "指定 module_name 则只导出那一个。导出文件里**含 Attribute 行**，"
                        "是改代码前留退路、以及对照 procedure attributes 的基准。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
                "module_name": {"type": "string", "description": "只导出该模块（默认全部）"},
                "out_dir": {"type": "string", "description": "输出目录（默认 <工作簿目录>/VBA_Backup/<时间戳>）"},
            },
            "required": ["workbook"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": True, "title": "备份 VBA 模块"},
    },
    {
        "name": "vba_import_module",
        "description": ("从 .bas/.cls/.frm 文件**整体替换**一个模块。这是唯一安全的整体替换方式："
                        "VBComponents.Import 会解析并保留 Attribute 行（含 VB_ProcData 快捷键绑定），"
                        "所以**文件里必须保留 Attribute 行、且按系统 ANSI(GBK) 编码**，否则中文注释会乱码。"
                        "同名的 Document 模块（Sheet/ThisWorkbook）不能这样替换。默认改前备份、默认保存。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
                "source_file": {"type": "string", "description": "源文件绝对路径（.bas/.cls/.frm/.txt）"},
                "module_name": {"type": "string", "description": "导入后的模块名（默认取文件名）"},
                "save": {"type": "boolean", "description": "是否落盘保存（默认 true）"},
                "no_backup": {"type": "boolean", "description": "true 则跳过改前自动备份"},
            },
            "required": ["workbook", "source_file"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": False, "title": "导入替换模块"},
    },
    {
        "name": "vba_list_references",
        "description": ("列出 VBA 工程引用（Tools > References），标出 is_broken 的项。"
                        "断引用必然导致编译失败（VBE 里报『找不到工程或库』），换台机器/换环境必炸。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "目标 .xlsm 绝对路径"},
            },
            "required": ["workbook"],
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": True, "title": "列工程引用"},
    },
]

HANDLERS = {
    "excel_status": tool_excel_status,
    "run_vba": tool_run_vba,
    "dismiss_dialog": tool_dismiss_dialog,
    "vba_modules": tool_vba_modules,
    "vba_patch_module": tool_vba_patch_module,
    "vba_check_compile": tool_vba_check_compile,
    "vba_backup": tool_vba_backup,
    "vba_import_module": tool_vba_import_module,
    "vba_list_references": tool_vba_list_references,
}


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
    _setup_stdio()          # 必须先于任何输出：否则中文描述按 GBK 写出 → 客户端乱码
    try:
        main()
    except (BrokenPipeError, KeyboardInterrupt):
        pass
