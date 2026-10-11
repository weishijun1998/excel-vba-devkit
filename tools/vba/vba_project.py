# -*- coding: utf-8 -*-
"""vba_project.py —— VBA 工程级读写与**真**编译检查（CLI）

为什么有它
----------
"读模块 / 改模块 / 编译检查 / 备份 / 看引用"这类**工程级**操作，原先依赖外部第三方
server（Downloads/vba-mcp-server-master）。它有两个硬伤：

  * `vba_write_module` 是"整模块重写"（DeleteLines 全部 + AddFromString）—— 会丢
    procedure attributes（SKILL.md 记录的事故：1874 行生产模块的快捷键绑定被静默抹掉）；
  * `vba_check_compile` 名不副实 —— 它的源码注释自己承认 "VBProject doesn't have a
    direct compile method exposed via COM / The best we can do is ..."，实际只查了断引用
    和行数，**从没编译过**。

本脚本按本 kit 的硬规则重做：

  * 改代码走**行级编辑**（DeleteLines / InsertLines），不做整模块重写。
    Attribute 行只存在于导出的 .bas 里、不在 CodeModule 里，行级编辑天然不碰它们。
  * 注入类操作（patch）若代码含 `Attribute` 行 → **自动剥离**（与 run_vba.py 一致；
    带 Attribute 的源码注入会让整个工程编译失败，规则 1）。
  * 整体替换只走 `import`（VBComponents.Import）—— 它会解析并保留 Attribute 行（规则 1）。
  * 编译检查是真的：注入一个空的探针宏并 Run 它 —— VBA 的编译是**工程级**的，Run 之前
    VBE 必须把整个 VBAProject 编译通过，所以"探针跑通"就等于"整个工程可编译"。跑完立即
    移除探针，磁盘不留痕。（网上通行的 `VBE.CommandBars.FindControl(Id:=578).Execute()`
    在本机中文 VBE 上**实测无效**：它返回的是"文件(&F)"菜单，扫遍菜单也没有编译项，
    Execute() 什么都不做却给"通过"的假结论 —— 比不检查更坏。）
  * 写操作默认**先自动备份**全部模块（--no-backup 关闭）。
  * 默认**不落盘**，要 `--save` 才写回工作簿。

子命令
------
  list    --workbook X                          列出全部模块（名/类型/行数）
  read    --workbook X (--module M | --all)     读模块源码（CodeModule 原文，含行号语义）
  patch   --workbook X --module M --op ...      行级编辑：replace/insert/delete/append
  compile --workbook X                          真编译检查（注入空探针宏后 Run）
  backup  --workbook X [--module M] [--out-dir] 备份/导出模块（含 manifest）
  import  --workbook X --source F [--module M]  从 .bas/.cls/.frm 导入（保留 Attribute）
  refs    --workbook X                          列出工程引用（Tools > References）

公共开关：--save（写回磁盘）/ --no-backup / --visible / --ascii / --allow-events

用法示例：
  python vba_project.py list --workbook "C:/path/x.xlsm"
  python vba_project.py read --workbook x.xlsm --module Module1
  python vba_project.py patch --workbook x.xlsm --module Module1 --op replace \
      --start 12 --end 14 --code-file fix.txt --save
  python vba_project.py compile --workbook x.xlsm
  python vba_project.py backup --workbook x.xlsm

退出码：0=成功；1=失败（编译错/找不到模块/COM 错）；2=参数或环境问题（缺依赖等）
"""
import argparse
import json
import locale
import os
import re
import sys
import tempfile
import threading
import time
from contextlib import contextmanager

HERE = os.path.dirname(os.path.abspath(__file__))

VB_COMP_TYPES = {1: "StandardModule", 2: "ClassModule", 3: "UserForm", 100: "Document"}
VB_COMP_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
DOC_TYPE = 100                      # Document 模块（Sheet/ThisWorkbook）不能删
VBA_XLSM_FORMAT = 52                # xlOpenXMLWorkbookMacroEnabled

# ---- 输出：永不因为终端编码而崩 -------------------------------------------
ASCII_MAP = {"✅": "[OK]", "❌": "[FAIL]", "⚠": "[!]", "ℹ": "[i]", "→": "->",
             "•": "-", "…": "..."}
ASCII_ONLY = False


def _setup_stdio():
    """stdio 一律 UTF-8（中文机器默认 cp936，会在管道里变乱码）。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
        try:
            stream.reconfigure(newline="\n")
        except Exception:
            pass


def emit(text):
    if ASCII_ONLY:
        for k, v in ASCII_MAP.items():
            text = text.replace(k, v)
        text = text.encode("ascii", "replace").decode("ascii")
    print(text, flush=True)


def jdump(obj):
    """统一 JSON 输出（--ascii 时也保证 JSON 本身可被任何终端读）。"""
    emit(json.dumps(obj, ensure_ascii=ASCII_ONLY, indent=2, default=str))


# ---- 源码文件编码：自动识别 ------------------------------------------------
# VBE 的 Export 写的是**系统 ANSI**（中文机器 = GBK）；用户手写的 .bas 可能是 UTF-8。
# 只认一种就会读不了或者读出乱码中文。
ANSI_CANDIDATES = ("cp936", "cp932", "cp949", "cp1252", "latin-1")


def read_source(path, forced=None):
    """读源码文件 → (文本, 实际编码)。"""
    with open(path, "rb") as f:
        raw = f.read()
    if forced:
        return raw.decode(forced), forced
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    order = ["utf-8"]
    try:
        pref = locale.getpreferredencoding(False)
        if pref and pref.lower().replace("-", "") != "utf8":
            order.append(pref)
    except Exception:
        pass
    order += [c for c in ANSI_CANDIDATES if c not in order]
    last = None
    for enc in order:
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, LookupError) as e:
            last = e
    raise last or UnicodeDecodeError("unknown", b"", 0, 1, "无法识别编码")


def system_ansi():
    try:
        return locale.getpreferredencoding(False) or "cp936"
    except Exception:
        return "cp936"


def write_source_text(path, text):
    """写"给 VBE 消费"的源码文件。

    **必须写成系统 ANSI（中文机器 = GBK）**：Import 按 ANSI 解析，写 UTF-8 会让每个
    中文注释变成乱码（SKILL.md 记录的坑）。不写 BOM。
    """
    enc = system_ansi()
    with open(path, "w", encoding=enc, newline="\r\n") as f:
        f.write(text)
    return enc


def crlf(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


# ---- Attribute 行（VBE 元数据）--------------------------------------------
# 只存在于导出的 .bas 里。注入到 CodeModule（AddFromString / InsertLines）会让
# 那个模块编译失败并**拖垮整个工程**（0x800A03EC +"宏可能被禁用"）—— 不是信任中心问题。
ATTR_LINE = re.compile(r"(?im)^[ \t]*Attribute\b.*(?:\r?\n|\Z)")


def strip_attributes(text):
    """剥掉 Attribute 行 → (干净文本, 剥掉的行数)。"""
    n = len(ATTR_LINE.findall(text))
    return ATTR_LINE.sub("", text), n


def count_proc_attribute(text):
    """数 procedure 级 Attribute（如 VB_ProcData 快捷键绑定）—— 用于改写前提示。"""
    return len(re.findall(r"(?im)^[ \t]*Attribute\s+\w+\.", text))


# ---- COM 会话 --------------------------------------------------------------
def _import_com():
    try:
        import psutil
        import win32com.client as w32
        import win32process
    except ImportError as e:
        emit("缺少依赖：%r —— 请先 pip install -r tools/vba/requirements.txt" % (e,))
        return None
    return w32, win32process, psutil


def _trust_hint():
    return ("需要在 Excel 里勾选『信任对 VBA 工程对象模型的访问』："
            "文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 勾选该项")


def explain_exception(exc):
    """把 COM 异常翻译成人话。绝不只报 hresult 数字 —— 同一个故障会伪装成不同码。"""
    s = str(exc)
    low = s.lower()
    if any(k in s for k in _IN_USE_MARKS):
        return ("工作簿正被另一个 Excel 实例占用（上一次调用的 Excel 退出有延迟，或该文件已在"
                "你打开的 Excel 里）。本脚本已自动退避重试；若仍失败：先关掉那个工作簿，"
                "或用 excel_status 看环境。原始错误：" + s)
    if ("not trusted" in low or "programmatic access" in low or "不信任" in s
            or "拒绝" in s and "vba" in low):
        return "程序化访问 VBA 工程被拒绝（信任中心未勾选）→ " + _trust_hint()
    if "0x800ac472" in low:
        return "程序化访问 VBA 工程被拒绝（0x800AC472）→ " + _trust_hint()
    if "0x800a03ec" in low:
        return ("0x800A03EC。两种常见原因：① 模块里混入了 Attribute 行（本项目已自动剥离）；"
                "② 工作簿窗口是隐藏态（装载器/宿主工作簿的常见设计）→ 用 run_vba.py 跑宏，别改文件")
    if "0x800a9c68" in low:
        return "0x800A9C68 = VBA 编译错误"
    return s


# "文件正被其他程序使用"——COM 打开工作簿时可能撞上的典型文案（中英文都列上）
_IN_USE_MARKS = ("其他程序使用", "正被其他程序", "不能访问文件", "被占用",
                 "in use", "being used", "Cannot access the file", "already open")


def _open_with_retry(xl, path, attempts=6):
    """打开工作簿；撞上"文件正被其他程序使用"就退避重试。

    为什么必须重试（实测）：上一次调用新建的 Excel 实例 `Quit()` 是**异步**的，进程要 1~3 秒
    才真正退出并释放文件锁。而 MCP 场景里 `vba_patch_module` → `vba_check_compile` 是**连续
    两次调用**（间隔可能只有几百毫秒），于是后一次会在文件仍被占用的瞬间失败，报出一条与
    代码毫无关系的错误（"Microsoft Excel 不能访问文件…文件正被其他程序使用"）。
    """
    last = None
    for i in range(attempts):
        try:
            return xl.Workbooks.Open(path)
        except Exception as e:
            if not any(k in str(e) for k in _IN_USE_MARKS):
                raise
            last = e
            time.sleep(0.6 + 0.6 * i)
    raise last


SESSION_ATTRS = ("DisplayAlerts", "AutomationSecurity", "EnableEvents", "Visible")


def get_excel(w32):
    """拿到 Excel 实例 → (xl, created)。**只负责拿，不改任何全局开关。**

    改开关与还原必须成对由调用方负责（见 excel_session）—— 复用用户正开着的实例时
    改动 Visible / AutomationSecurity 这类全局状态而不还原，就是破坏用户会话。
    """
    try:
        xl = w32.GetActiveObject("Excel.Application")
    except Exception:
        xl = None
    if xl is None:
        return w32.DispatchEx("Excel.Application"), True
    return xl, False


def _workbook_tail(path_or_url):
    """把本地路径或 OneDrive URL 归一成可比较的尾部分段（父目录/文件名）。"""
    s = str(path_or_url).replace("\\", "/")
    if "://" in s:
        import urllib.parse
        s = urllib.parse.unquote(s.split("://", 1)[1])
    parts = [p for p in s.split("/") if p]
    return "/".join(parts[-2:]).lower()


def _ensure_excel_exit(pid, timeout=4.0):
    """确认**本进程创建**的 Excel 真的退出；到点还在就 kill。

    为什么必须做（实测）：`Quit()` 只是"请求退出"。注入过 VBA 模块、触发过 VBE 编译的实例，
    实测会留下**无窗口的僵尸进程** —— 进程还在、仍持有目标工作簿，于是后续每次调用都报
    「Microsoft Excel 不能访问文件…文件正被其他程序使用」，而且退避重试十几秒也过不去。
    这个实例是本进程创建的，kill 它绝不碰用户自己开的 Excel。
    """
    if not pid:
        return None
    try:
        import psutil
    except ImportError:
        return None
    t0 = time.time()
    while time.time() - t0 < timeout:
        if not psutil.pid_exists(pid):
            return "exited"
        time.sleep(0.25)
    try:
        psutil.Process(pid).kill()
        return "killed"
    except Exception:
        return "kill_failed"


def _same_workbook(wb, path):
    """wb 是否就是目标文件。

    ⚠ OneDrive / SharePoint 同步目录里的工作簿，`wb.FullName` 返回的是 **https URL**
    （形如 https://tenant-my.sharepoint.com/personal/u/Documents/桌面/x.xlsm），而本地路径是
    C:/Users/…/OneDrive - X/桌面/x.xlsm —— 前缀完全不同，用 abspath 比较**永远不相等**，
    于是每次都会把同一个文件再打开一遍（可能弹出"保存冲突副本"）。
    这种情况退化为比较"父目录/文件名"尾部分段（并要求本地文件确实存在）。
    """
    try:
        full = str(wb.FullName)
    except Exception:
        return False
    if full.startswith(("http://", "https://")):
        return os.path.exists(path) and _workbook_tail(full) == _workbook_tail(path)
    return os.path.normcase(os.path.abspath(full)) == os.path.normcase(os.path.abspath(path))


def find_open_workbook(xl, path):
    try:
        n = xl.Workbooks.Count
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            wb = xl.Workbooks.Item(i)
            if _same_workbook(wb, path):
                return wb
        except Exception:
            pass
    return None


@contextmanager
def excel_session(w32, win32process, path, create=False, visible=False,
                  allow_events=False, save=False):
    """打开目标工作簿 → yield (xl, wb, info) → 收拾现场。

    收拾规则（别把用户的 Excel 弄乱）：
      * 工作簿本来没开 → 我们开的 → 关掉（除非 save 已经保住磁盘版本）
      * 工作簿本来开着 → 原样留着（用户的东西不动）
      * Excel 实例本进程新建的 → Quit；复用的实例绝不 Quit
    """
    path = os.path.abspath(path)
    if not os.path.exists(path) and not create:
        raise FileNotFoundError("工作簿不存在: %s" % path)
    if os.path.exists(path) and not path.lower().endswith((".xlsm", ".xlsb", ".xls", ".xlam", ".xla")):
        raise ValueError("不是可承载 VBA 的文件: %s" % path)

    xl, created_excel = get_excel(w32)
    info = {"reused_instance": not created_excel, "opened_by_us": False, "saved": False,
            "reused_app": not created_excel, "visible_untouched": not created_excel,
            "local_dir": os.path.dirname(os.path.abspath(path))}
    # 记录全局开关原值：复用用户的实例时，收尾必须原样放回去
    prev = {}
    for attr in SESSION_ATTRS:
        try:
            prev[attr] = getattr(xl, attr)
        except Exception:
            prev[attr] = None
    try:
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1          # 别让宏安全策略弹框打断
        if created_excel:
            xl.Visible = bool(visible)     # 自己建的实例：默认隐藏（快）
        if not allow_events:
            # 装载器/宿主工作簿常常"一打开就跑程序"——工程级读写绝不能顺带触发它
            xl.EnableEvents = False
    except Exception:
        pass
    wb = None
    pid = None
    try:
        pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    except Exception:
        pass
    info["excel_pid"] = pid

    try:
        wb = find_open_workbook(xl, path)
        if wb is None:
            if os.path.exists(path):
                wb = _open_with_retry(xl, path)
            else:
                wb = xl.Workbooks.Add()
                wb.SaveAs(path, FileFormat=VBA_XLSM_FORMAT)
            info["opened_by_us"] = True
        info["workbook"] = str(wb.Name)
        yield xl, wb, info
    finally:
        try:
            # info["saved"] 可能已被命令体显式置位（命令体里保存才能拿到真实的 "saved" 回执）
            if save and wb is not None and not info.get("saved"):
                wb.Save()
                info["saved"] = True
        except Exception as e:
            emit("⚠ 保存失败: %s" % explain_exception(e))
        if wb is not None and info.get("opened_by_us"):
            try:
                wb.Close(SaveChanges=False)   # 要存早存了；这里绝不再弹"是否保存"
            except Exception:
                pass
        if created_excel:
            try:
                xl.Quit()                     # 只关自己建出来的实例；复用用户实例时绝不 Quit
            except Exception:
                pass
            info["excel_exit"] = _ensure_excel_exit(pid)   # Quit 不等于退出，见函数说明
            if info["excel_exit"] == "killed":
                emit("ℹ 自建的 Excel 实例 Quit 后未自行退出，已强杀"
                     "（否则它会占住工作簿，让后续调用报『文件正被其他程序使用』）")
        else:
            # 复用用户的实例：把动过的全局开关全部还原（尤其 Visible —— 否则用户的
            # Excel 窗口会被静默藏掉，进程还在但用户以为程序被关了）
            for attr in SESSION_ATTRS:
                val = prev.get(attr)
                if val is None:
                    continue
                try:
                    setattr(xl, attr, val)
                except Exception:
                    pass


def get_module(wb, name):
    for c in wb.VBProject.VBComponents:
        if c.Name == name:
            return c
    return None


def module_names(wb):
    return [c.Name for c in wb.VBProject.VBComponents]


# ============================ 子命令 ========================================

def cmd_list(args, w32, win32process):
    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events) as (xl, wb, info):
        mods = []
        for c in wb.VBProject.VBComponents:
            cm = c.CodeModule
            n = cm.CountOfLines if cm is not None else 0
            mods.append({
                "name": c.Name,
                "type": VB_COMP_TYPES.get(c.Type, "Unknown(%s)" % c.Type),
                "type_id": c.Type,
                "lines": n,
                "deletable": c.Type != DOC_TYPE,
            })
        jdump({"ok": True, "workbook": str(wb.FullName), "count": len(mods),
               "modules": mods, "excel_pid": info.get("excel_pid"),
               "reused_instance": info["reused_instance"]})
    return 0


def cmd_read(args, w32, win32process):
    if not args.module and not args.all:
        emit("⛔ 要么给 --module <名>，要么给 --all")
        return 2
    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events) as (xl, wb, info):
        if args.all:
            out = []
            for c in wb.VBProject.VBComponents:
                cm = c.CodeModule
                n = cm.CountOfLines if cm is not None else 0
                out.append({
                    "name": c.Name,
                    "type": VB_COMP_TYPES.get(c.Type, "Unknown"),
                    "lines": n,
                    "code": (cm.Lines(1, n) if n else ""),
                })
            jdump({"ok": True, "workbook": str(wb.FullName), "modules": out,
                   "note": "行号从 1 开始；CodeModule 不包含 Attribute 行（它只在导出的文件里）"})
            return 0
        comp = get_module(wb, args.module)
        if comp is None:
            jdump({"ok": False, "error": "模块不存在: %s" % args.module,
                   "available_modules": module_names(wb)})
            return 1
        cm = comp.CodeModule
        n = cm.CountOfLines if cm is not None else 0
        jdump({"ok": True, "workbook": str(wb.FullName), "module": comp.Name,
               "type": VB_COMP_TYPES.get(comp.Type, "Unknown"),
               "lines": n, "code": (cm.Lines(1, n) if n else ""),
               "note": "行号从 1 开始；CodeModule 不包含 Attribute 行（它只在导出的文件里）"})
    return 0


def cmd_patch(args, w32, win32process):
    op = (args.op or "").lower().strip()
    if op not in ("replace", "insert", "delete", "append"):
        jdump({"ok": False, "error": "op 必须是 replace/insert/delete/append 之一"})
        return 2
    code = ""
    if args.code_file:
        if not os.path.exists(args.code_file):
            jdump({"ok": False, "error": "代码文件不存在: %s" % args.code_file})
            return 2
        code, src_enc = read_source(args.code_file, args.encoding)
    elif args.code:
        code, src_enc = args.code, "inline"
    else:
        code, src_enc = "", "none"

    stripped = 0
    if code and op != "delete":
        code, stripped = strip_attributes(code)
        code = crlf(code).strip("\r\n")
        if stripped:
            emit("ℹ 已剥离 %d 行 Attribute（VBE 元数据，注入会让整个工程编译失败）" % stripped)

    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events, save=args.save) as (xl, wb, info):
        comp = get_module(wb, args.module)
        if comp is None:
            jdump({"ok": False, "error": "模块不存在: %s" % args.module,
                   "available_modules": module_names(wb)})
            return 1
        cm = comp.CodeModule
        total = cm.CountOfLines if cm is not None else 0

        # 写前自动备份：生产工作簿改一行也要有退路
        backup_dir = None
        if args.backup:
            try:
                backup_dir = _do_backup(wb, None, args.out_dir, tag="before-patch",
                                        base_dir=pick_base_dir(wb, args.workbook))
            except Exception as e:
                emit("⚠ 自动备份失败（继续，但你少了一张安全网）: %s" % explain_exception(e))

        result = {"ok": True, "module": comp.Name, "operation": op,
                  "lines_before": total, "backup_dir": backup_dir,
                  "source_encoding": src_enc, "stripped_attribute_lines": stripped}

        if op == "append":
            if code:
                cm.InsertLines(total + 1, code)
            result.update(inserted_after_line=total)
        elif op == "insert":
            if args.start is None:
                jdump({"ok": False, "error": "insert 需要 --start（插到该行之后；0 = 插到开头）"})
                return 2
            pos = args.start
            if pos < 0 or pos > total:
                jdump({"ok": False, "error": "start=%s 超出模块行数(%s)" % (pos, total)})
                return 1
            if code:
                cm.InsertLines(pos + 1 if pos >= 1 else 1, code)
            result.update(inserted_after_line=pos)
        elif op == "delete":
            if args.start is None:
                jdump({"ok": False, "error": "delete 需要 --start [--end]"})
                return 2
            if args.start < 1 or args.start > total:
                jdump({"ok": False, "error": "start=%s 超出 1..%s" % (args.start, total)})
                return 1
            end = min(args.end or args.start, total)
            n = end - args.start + 1
            cm.DeleteLines(args.start, n)
            result.update(deleted_range=[args.start, end], lines_removed=n)
        else:  # replace
            if args.start is None:
                jdump({"ok": False, "error": "replace 需要 --start [--end]"})
                return 2
            if args.start > total + 1:
                jdump({"ok": False, "error": "start=%s 超出模块末尾(%s)，要追加请用 --op append"
                       % (args.start, total)})
                return 1
            if args.start > total:
                if code:
                    cm.InsertLines(total + 1, code)
                result.update(replaced_range="(beyond end -> appended)", lines_removed=0)
            else:
                end = min(args.end or args.start, total)
                n = end - args.start + 1
                cm.DeleteLines(args.start, n)
                if code:
                    cm.InsertLines(args.start, code)
                result.update(replaced_range=[args.start, end], lines_removed=n)

        result["lines_after"] = cm.CountOfLines
        if args.save:
            try:
                wb.Save()
                info["saved"] = True
            except Exception as e:
                jdump({"ok": False, "error": "保存失败：%s" % explain_exception(e),
                       "module": comp.Name, "lines_after": cm.CountOfLines})
                return 1
        result["saved"] = info["saved"]
        if not args.save:
            result["warning"] = "未落盘 —— 加 --save 才写回工作簿"
        jdump(result)
    return 0


def cmd_backup(args, w32, win32process):
    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events) as (xl, wb, info):
        dest = _do_backup(wb, args.module, args.out_dir, tag=None,
                          base_dir=pick_base_dir(wb, args.workbook))
        jdump({"ok": True, "workbook": str(wb.FullName), "backup_dir": dest,
               "encoding_note": "VBE 导出的是系统 ANSI(%s)；读回请用 vba_project 的自动识别"
                                % system_ansi()})
    return 0


def _is_url_path(p):
    return bool(re.match(r"(?i)^\s*(https?|ftp|file)://", str(p or "")))


def pick_base_dir(wb, workbook_arg):
    """挑一个**能写文件**的本地目录做备份根。

    ⚠ 实测坑：OneDrive / SharePoint 上打开的 .xlsm，`wb.FullName` 会是 **URL**，例如
    `https://carlsberggroup-my.sharepoint.com/personal/<user>/Documents/桌面/VBA MCP测试/x.xlsm`。
    拿它当路径拼备份目录 → `...\\_selftest\\https:` → `WinError 123 文件名、目录名或卷标语法不正确`，
    于是"备份"静默失败（这类工作簿恰恰是最需要备份的生产件）。
    顺序：调用方给的本地路径 → wb.Path（本地时）→ 系统临时目录。
    """
    cands = []
    if workbook_arg:
        cands.append(os.path.dirname(os.path.abspath(str(workbook_arg))))
    try:
        cands.append(str(wb.Path))
    except Exception:
        pass
    for c in cands:
        if c and not _is_url_path(c) and os.path.isdir(c):
            return c
    return tempfile.gettempdir()


def _do_backup(wb, only_module, out_dir, tag, base_dir=None):
    """导出模块到磁盘 → 返回目录。单个模块时直接放文件。

    `base_dir` 由调用方通过 `pick_base_dir(wb, <本地工作簿路径>)` 得到：
    OneDrive 工作簿的 `wb.FullName` 是 https URL，拿它拼目录会 `WinError 123` 让备份静默失败。
    """
    base = base_dir or pick_base_dir(wb, None)
    if out_dir:
        dest = os.path.abspath(out_dir)
    else:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        name = ("%s_%s" % (stamp, tag)) if tag else stamp
        dest = os.path.join(base, "VBA_Backup", name)
    os.makedirs(dest, exist_ok=True)
    exported = []
    for c in wb.VBProject.VBComponents:
        if only_module and c.Name != only_module:
            continue
        ext = VB_COMP_EXT.get(c.Type, ".txt")
        path = os.path.join(dest, "%s%s" % (c.Name, ext))
        c.Export(path)
        item = {"name": c.Name, "type": VB_COMP_TYPES.get(c.Type, "Unknown"),
                "file": path, "lines": c.CodeModule.CountOfLines if c.CodeModule else 0}
        try:
            with open(path, "rb") as f:
                head = f.read(4000)
            text, enc = _decode_head(head)
            item["file_encoding"] = enc
            item["proc_attributes"] = count_proc_attribute(text)
        except Exception:
            pass
        exported.append(item)
    if len(exported) > 1:
        manifest = {"source_file": str(wb.FullName), "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "modules": exported, "encoding": system_ansi()}
        with open(os.path.join(dest, "_manifest.json"), "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
    return dest


def _decode_head(raw):
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig", "replace"), "utf-8-sig"
    for enc in ("utf-8",) + ANSI_CANDIDATES:
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("latin-1", "replace"), "latin-1"


def cmd_import(args, w32, win32process):
    src = os.path.abspath(args.source)
    if not os.path.isfile(src):
        jdump({"ok": False, "error": "源文件不存在: %s" % src})
        return 2
    ext = os.path.splitext(src)[1].lower()
    if ext not in (".bas", ".cls", ".frm", ".txt"):
        jdump({"ok": False, "error": "不支持的扩展名 %s（要 .bas/.cls/.frm/.txt）" % ext})
        return 2
    mod_name = args.module or os.path.splitext(os.path.basename(src))[0]

    # 读一遍只为体检（Import 需要 Attribute 行；这里**不剥离**）
    try:
        text, enc = read_source(src, args.encoding)
    except UnicodeDecodeError as e:
        text, enc = "", "unknown(%s)" % e
    has_name_attr = bool(re.search(r'(?im)^\s*Attribute\s+VB_Name\s*=', text))
    proc_attrs = count_proc_attribute(text)

    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events, save=args.save) as (xl, wb, info):
        backup_dir = None
        if args.backup:
            try:
                backup_dir = _do_backup(wb, None if mod_name not in module_names(wb) else mod_name,
                                        args.out_dir, tag="before-import",
                                        base_dir=pick_base_dir(wb, args.workbook))
            except Exception as e:
                emit("⚠ 自动备份失败: %s" % explain_exception(e))
        existing = get_module(wb, mod_name)
        replaced = False
        if existing is not None:
            if existing.Type == DOC_TYPE:
                jdump({"ok": False, "error": "%s 是 Document 模块（Sheet/ThisWorkbook），"
                                            "不能用 import 替换" % mod_name})
                return 1
            wb.VBProject.VBComponents.Remove(existing)
            replaced = True
        comp = wb.VBProject.VBComponents.Import(src)
        try:
            comp.Name = mod_name          # 文件名即模块名；显式对齐
        except Exception:
            pass
        if args.save:
            try:
                wb.Save()
                info["saved"] = True
            except Exception as e:
                jdump({"ok": False, "error": "保存失败：%s" % explain_exception(e)})
                return 1
        jdump({"ok": True, "module": comp.Name, "imported_from": src,
               "source_encoding": enc, "attribute_vb_name_present": has_name_attr,
               "proc_attributes_in_file": proc_attrs, "replaced_existing": replaced,
               "lines": comp.CodeModule.CountOfLines if comp.CodeModule else 0,
               "backup_dir": backup_dir, "saved": info["saved"],
               "note": ("Import 会解析并保留 Attribute 行（这是唯一安全的整体替换方式）；"
                        "文件必须按系统 ANSI(%s) 编码，否则中文注释会乱码" % system_ansi())
               + ("" if has_name_attr else "；⚠ 文件里没有 Attribute VB_Name = \"...\" 行，"
                                           "模块名按文件名决定")})
    return 0


def cmd_refs(args, w32, win32process):
    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events) as (xl, wb, info):
        refs, broken = [], []
        for r in wb.VBProject.References:
            item = {"name": r.Name, "guid": r.Guid, "major": r.Major, "minor": r.Minor,
                    "is_broken": bool(r.IsBroken)}
            for attr in ("Description", "FullPath"):
                try:
                    item[attr.lower()] = getattr(r, attr)
                except Exception:
                    item[attr.lower()] = None
            refs.append(item)
            if item["is_broken"]:
                broken.append(r.Name)
        jdump({"ok": True, "workbook": str(wb.FullName), "count": len(refs),
               "references": refs, "broken": broken,
               "hint": ("断引用必然导致编译失败：VBE 里会报『找不到工程或库』。"
                        "处理办法：在 VBE 工具>引用里取消勾选，或装上对应的库/路径。") if broken else None})
    return 0


# ---- 真编译检查：VBE 578 "Compile" 命令 + 弹窗捕捉 --------------------------
# 要点：执行编译命令时若工程有语法错误，VBE 会弹**模态**对话框并阻塞 Execute()。
# 所以必须一边执行、一边点掉弹窗（点掉后 Execute 才会返回），否则整个进程卡死。
_BTN_TEXTS = ("确定", "OK", "结束", "End", "关闭", "Close")
_BTN_IDS = (4800, 1, 2)


def _dialog_watcher(stop, hits, pid):
    import win32con
    import win32gui
    import win32process

    while not stop["stop"]:
        found = []
        def cb(h, _):
            try:
                if not win32gui.IsWindowVisible(h):
                    return True
                if win32gui.GetClassName(h) != "#32770":
                    return True
                if pid:
                    _, wp = win32process.GetWindowThreadProcessId(h)
                    if wp != pid:
                        return True
                title = win32gui.GetWindowText(h)
                if "Visual Basic" in title or "VBA" in title:
                    found.append(h)
            except Exception:
                pass
            return True
        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass

        for h in found:
            kids = []
            def kcb(c, _):
                try:
                    kids.append((c, win32gui.GetDlgCtrlID(c),
                                 win32gui.GetClassName(c), win32gui.GetWindowText(c)))
                except Exception:
                    pass
                return True
            try:
                win32gui.EnumChildWindows(h, kcb, None)
            except Exception:
                pass
            msg = next((t for _, i, cl, t in kids if i == 4803), "")
            if not msg:
                msg = next((t for _, i, cl, t in kids if cl == "Static" and len(t) > 8), "")
            buttons = [(c, i, t) for c, i, cl, t in kids if cl == "Button"]
            hits.append({"title": win32gui.GetWindowText(h), "text": msg,
                         "buttons": [t for _, _, t in buttons]})
            target = None
            for c, i, t in buttons:                       # 先按文案（编译错误框是"确定"）
                if t in _BTN_TEXTS or any(t.startswith(x) for x in _BTN_TEXTS):
                    target = (c, i, t)
                    break
            if target is None:
                for c, i, t in buttons:                   # 再按 id（运行期错误框 4800=结束）
                    if i in _BTN_IDS:
                        target = (c, i, t)
                        break
            if target is None and buttons:
                target = buttons[0]
            if target:
                c, i, t = target
                try:
                    win32gui.PostMessage(c, win32con.BM_CLICK, 0, 0)
                except Exception:
                    try:
                        win32gui.PostMessage(h, win32con.WM_COMMAND, i, c)
                    except Exception:
                        pass
                hits[-1]["clicked"] = t
        time.sleep(0.1)


# ---- 真编译检查：注入空探针宏 + Run 触发 VBE 的工程级编译 -------------------
# 为什么不用 `VBE.CommandBars.FindControl(Id:=578).Execute()`（网上通行的做法）：
# **实测在中文 VBE 上它返回的是"文件(&F)"菜单**，扫遍 VBE 菜单也找不到任何编译项，
# 于是 Execute() 什么都不编译、也不报错 —— 会给出"编译通过"的假结论（比不检查更坏）。
#
# VBA 的编译是**工程级**的：Run 任何一个宏之前，VBE 必须先把整个 VBAProject 编译通过。
# 所以：注入一个空的探针宏 → Run 它 → 成功即"整个工程可编译"，失败即"工程编译不过"。
# 探针是空过程（无副作用），跑完立刻移除，磁盘上不留痕（除非 --save）。
#
# 注意：工程编译不过时 Excel 报的是「无法运行"X"宏…可能被禁用」—— 与"宏不存在"、
# "宏被禁用"共用同一文案（这正是 SKILL.md 规则 1/2 的坑）。我们已经确认探针模块确实
# 注入成功，所以这个文案在此处只能解释为编译/可用性失败；输出里会给出这个前提。
#
# ⚠ 实测坑（复现：`python tools/vba/vbe_compile_probe.py` 的 [3]）：**过程名不能与模块名相同**。
# 若模块 M 里有 `Sub M()`，`Run("M")` / `Run("book.xlsm!M")` / `Run("'book.xlsm'!M")` **全部失败**，
# 报的还是「无法运行…宏」那句文案 —— 看起来像编译错误，其实是名字解析歧义。所以探针的模块名与
# 过程名必须分开。
PROBE_NAME = "VBA_KIT_COMPILE_PROBE"      # 模块名
PROBE_PROC = "VBAPROBE_Run"               # 过程名：必须与模块名不同
_UNAVAILABLE_MARKS = ("无法运行", "Cannot run the macro", "宏在此工作簿中不可用",
                      "may be disabled", "所有的宏都被禁用")


def _dialog_watcher(stop, hits, pid):
    import win32con
    import win32gui
    import win32process

    while not stop["stop"]:
        found = []
        def cb(h, _):
            try:
                if not win32gui.IsWindowVisible(h):
                    return True
                if win32gui.GetClassName(h) != "#32770":
                    return True
                if pid:
                    _, wp = win32process.GetWindowThreadProcessId(h)
                    if wp != pid:
                        return True
                title = win32gui.GetWindowText(h)
                if "Visual Basic" in title or "VBA" in title:
                    found.append(h)
            except Exception:
                pass
            return True
        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass

        for h in found:
            kids = []
            def kcb(c, _):
                try:
                    kids.append((c, win32gui.GetDlgCtrlID(c),
                                 win32gui.GetClassName(c), win32gui.GetWindowText(c)))
                except Exception:
                    pass
                return True
            try:
                win32gui.EnumChildWindows(h, kcb, None)
            except Exception:
                pass
            msg = next((t for _, i, cl, t in kids if i == 4803), "")
            if not msg:
                msg = next((t for _, i, cl, t in kids if cl == "Static" and len(t) > 8), "")
            buttons = [(c, i, t) for c, i, cl, t in kids if cl == "Button"]
            hits.append({"title": win32gui.GetWindowText(h), "text": msg,
                         "buttons": [t for _, _, t in buttons]})
            target = None
            for c, i, t in buttons:                       # 先按文案（编译错误框是"确定"）
                if t in _BTN_TEXTS or any(t.startswith(x) for x in _BTN_TEXTS):
                    target = (c, i, t)
                    break
            if target is None:
                for c, i, t in buttons:                   # 再按 id（运行期错误框 4800=结束）
                    if i in _BTN_IDS:
                        target = (c, i, t)
                        break
            if target is None and buttons:
                target = buttons[0]
            if target:
                c, i, t = target
                try:
                    win32gui.PostMessage(c, win32con.BM_CLICK, 0, 0)
                except Exception:
                    try:
                        win32gui.PostMessage(h, win32con.WM_COMMAND, i, c)
                    except Exception:
                        pass
                hits[-1]["clicked"] = t
        time.sleep(0.1)


def cmd_compile(args, w32, win32process):
    with excel_session(w32, win32process, args.workbook, visible=args.visible,
                       allow_events=args.allow_events, save=args.save) as (xl, wb, info):
        result = {"ok": False, "compiled": False, "workbook": str(wb.FullName),
                  "method": "inject-empty-probe-then-run (VBA 编译是工程级的)",
                  "probe_module": PROBE_NAME, "probe_proc": PROBE_PROC,
                  "excel_pid": info.get("excel_pid"),
                  "visible_untouched": info.get("visible_untouched"),
                  "dialogs": [], "broken_refs": []}

        # 预检：断引用必然编译失败，能给出比"编译不过"更精确的结论
        try:
            for r in wb.VBProject.References:
                if r.IsBroken:
                    result["broken_refs"].append(r.Name)
        except Exception as e:
            result["preflight_error"] = explain_exception(e)

        vbp = wb.VBProject
        comp, preexisting = get_module(wb, PROBE_NAME), False
        if comp is not None:
            preexisting = True
        else:
            try:
                comp = vbp.VBComponents.Add(1)
                comp.Name = PROBE_NAME
                comp.CodeModule.AddFromString("Sub %s()\r\nEnd Sub\r\n" % PROBE_PROC)
            except Exception as e:
                result.update(error="注入探针模块失败：%s" % explain_exception(e),
                              hint="工程被密码保护 / 信任中心未勾选 / 工程已损坏时都会这样")
                jdump(result)
                return 1

        stop = {"stop": False}
        hits = []
        watcher = threading.Thread(target=_dialog_watcher,
                                   args=(stop, hits, info.get("excel_pid")), daemon=True)
        watcher.start()
        run_err = None
        t0 = time.time()
        try:
            try:
                xl.Run("%s!%s" % (wb.Name, PROBE_PROC))
            except Exception:
                xl.Run("'%s'!%s" % (wb.Name, PROBE_PROC))     # 工作簿名含空格时用引号形式
        except Exception as e:
            run_err = e
        time.sleep(args.dialog_wait)
        stop["stop"] = True
        watcher.join(timeout=2)

        # 先把探针摘掉，再让会话收尾保存 —— 磁盘上不能留探针
        if not preexisting and comp is not None:
            try:
                vbp.VBComponents.Remove(comp)
                result["probe_removed"] = True
            except Exception as e:
                result["probe_removed"] = False
                result["cleanup_error"] = explain_exception(e)

        result["dialogs"] = hits
        result["elapsed_s"] = round(time.time() - t0, 2)

        if hits:
            result.update(ok=False, compiled=True, verdict="COMPILE_ERRORS",
                          message="编译未通过：捕捉到 %d 个 VBE 对话框" % len(hits))
            jdump(result)
            return 1

        if run_err is not None:
            msg = str(run_err)
            if any(k in msg for k in _UNAVAILABLE_MARKS):
                result.update(
                    ok=False, compiled=True, verdict="COMPILE_ERRORS",
                    error=explain_exception(run_err),
                    premise="探针模块已确认注入成功，故该文案只能解释为工程编译不过 / 宏不可用",
                    hint=("若你确信工程本来是可编译的，做一次对照：在这个 Excel 里手动运行任意宏。"
                          "手动也跑不起来 ⇒ 是宏安全策略/信任中心，不是代码问题。"))
            else:
                result.update(ok=False, compiled=False, verdict="RUN_FAILED",
                              error=explain_exception(run_err),
                              hint="探针宏没能跑起来，且不是'无法运行宏'文案 —— 属环境/COM 问题")
            jdump(result)
            return 1

        if result["broken_refs"]:
            result.update(ok=False, compiled=True, verdict="BROKEN_REFERENCES",
                          message="编译未报错，但存在断引用（换台机器/换环境必炸）")
            jdump(result)
            return 1

        result.update(ok=True, compiled=True, verdict="CLEAN",
                      message="编译通过：整个 VBAProject 编译成功（空探针宏已运行并移除）")
        jdump(result)
        return 0


# ============================ main ==========================================

def main():
    # 不用子命令（subparsers）：那样 `list --workbook X` 会解析失败（全局选项只能在子命令之前）。
    # 位置参数 + 统一选项 ⇒ 选项放前面放后面都能用，模型抄文档命令时不会踩坑。
    ap = argparse.ArgumentParser(add_help=True, description="VBA 工程读写与真编译检查")
    ap.add_argument("cmd", nargs="?", choices=["list", "read", "patch", "compile",
                                              "backup", "import", "refs"],
                    help="子命令")
    ap.add_argument("--workbook", help="目标 .xlsm 绝对路径")
    ap.add_argument("--module", default=None, help="模块名")
    ap.add_argument("--op", default=None, choices=["replace", "insert", "delete", "append"],
                    help="patch 的操作类型")
    ap.add_argument("--start", type=int, default=None, help="1-based 起始行（insert 时 0=插到开头）")
    ap.add_argument("--end", type=int, default=None, help="1-based 结束行（含）")
    ap.add_argument("--code", default=None, help="要写入的代码（短代码用；长代码用 --code-file）")
    ap.add_argument("--code-file", default=None, help="要写入的代码文件")
    ap.add_argument("--source", default=None, help="import 的源文件（.bas/.cls/.frm）")
    ap.add_argument("--all", action="store_true", help="read: 读全部模块")
    ap.add_argument("--save", action="store_true", help="写操作后保存工作簿（默认不落盘）")
    ap.add_argument("--no-backup", dest="backup", action="store_false",
                    help="写操作前不自动备份（默认备份到 VBA_Backup/<时间戳>）")
    ap.add_argument("--out-dir", default=None, help="备份目录（默认 <工作簿目录>/VBA_Backup/<时间戳>）")
    ap.add_argument("--visible", action="store_true", help="让 Excel 可见（默认隐藏）")
    ap.add_argument("--allow-events", action="store_true",
                    help="允许触发工作簿事件（默认禁用，避免顺带跑起装载器的 Workbook_Open）")
    ap.add_argument("--encoding", default=None, help="源码文件编码（默认自动识别）")
    ap.add_argument("--dialog-wait", type=float, default=1.5,
                    help="Run 探针后再等几秒收集可能出现的对话框（默认 1.5s）")
    ap.add_argument("--ascii", action="store_true", help="输出只用 ASCII（任何终端都不乱码/不崩）")
    ap.add_argument("--selftest", action="store_true", help="离线自检（不碰 Excel）")

    args = ap.parse_args()
    global ASCII_ONLY
    ASCII_ONLY = bool(args.ascii)
    _setup_stdio()

    if args.selftest:
        return selftest()

    if not args.cmd:
        ap.print_help()
        return 2
    if not args.workbook:
        emit("⛔ 缺少 --workbook")
        return 2
    if args.cmd == "patch" and (not args.module or not args.op):
        emit("⛔ patch 需要 --module 与 --op")
        return 2
    if args.cmd == "import" and not args.source:
        emit("⛔ import 需要 --source")
        return 2

    com = _import_com()
    if com is None:
        return 2
    w32, win32process, _psutil = com

    handler = {"list": cmd_list, "read": cmd_read, "patch": cmd_patch,
               "compile": cmd_compile, "backup": cmd_backup,
               "import": cmd_import, "refs": cmd_refs}[args.cmd]
    try:
        return handler(args, w32, win32process)
    except FileNotFoundError as e:
        jdump({"ok": False, "error": str(e)})
        return 2
    except Exception as e:
        jdump({"ok": False, "error": explain_exception(e), "cmd": args.cmd,
               "hint": "若是 COM 错误，先确认 Excel 没有弹着模态对话框、且信任中心已勾选"})
        return 1


def selftest():
    """离线自检：不碰 Excel，验证编码/剥离/换行等纯逻辑。"""
    import tempfile
    ok = True
    checks = []

    raw = 'Attribute VB_Name = "M"\r\nSub A()\r\n    MsgBox "x" \' 注释\r\nEnd Sub\r\n'
    clean, n = strip_attributes(raw)
    checks.append(("strip_attributes 行数", n == 1))
    checks.append(("strip 后无 Attribute", "Attribute" not in clean))
    checks.append(("strip 保留代码", 'Sub A()' in clean and 'MsgBox' in clean))
    ok &= all(c[1] for c in checks)

    proc = 'Attribute VB_Name = "M"\r\nAttribute A.VB_ProcData.VB_Invoke_Func = "P\\n14"\r\nSub A()\r\nEnd Sub'
    checks.append(("procedure attribute 计数", count_proc_attribute(proc) == 1))
    ok &= checks[-1][1]

    checks.append(("crlf 规范化", crlf("a\nb\r\nc") == "a\r\nb\r\nc"))
    ok &= checks[-1][1]

    with tempfile.TemporaryDirectory() as d:
        f = os.path.join(d, "t.bas")
        enc = write_source_text(f, 'Sub 中文()\r\nEnd Sub\r\n')
        text, back = read_source(f)
        checks.append(("ANSI 写回读一致", "中文" in text))
        ok &= checks[-1][1]
        f2 = os.path.join(d, "u.bas")
        with open(f2, "wb") as fh:
            fh.write('Sub 中文()\r\nEnd Sub\r\n'.encode("utf-8-sig"))
        t2, e2 = read_source(f2)
        checks.append(("UTF-8 BOM 识别", e2 == "utf-8-sig" and "中文" in t2))
        ok &= checks[-1][1]

    jdump({"ok": bool(ok), "system_ansi": system_ansi(),
           "checks": [{"name": k, "pass": v} for k, v in checks]})
    return 0 if ok else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (BrokenPipeError, KeyboardInterrupt):
        pass
