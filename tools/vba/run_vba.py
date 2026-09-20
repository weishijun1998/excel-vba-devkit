# -*- coding: utf-8 -*-
"""run_vba.py —— VBA 宏的"开发测试跑道"（一步完成：注入 → 守卫 → 运行 → 断言 → 报告）

设计目标：让"写宏 → 跑 → 看结果 → 改 → 再跑"每轮又快又能自己说清哪儿错了。

一条命令做的事：
  1) 读代码文件，剥掉 Attribute 行（否则必报 0x800A03EC"宏不可用"）
  2) 静态检查危险语句（MsgBox / InputBox / Stop / Debug.Assert / .Show）→ 默认拒绝注入
  3) 复用已打开的 Excel（若目标工作簿已开着）或新建实例
  4) 注入模块（同名先删），运行前保存（被守卫杀掉也能从磁盘继续）
  5) 起独立守卫进程（弹窗点掉 / 假死判定 / 跑飞判定）
  6) 运行宏，判定结果（正常 / 报错弹窗 / 卡死被杀 / 超预算）
  7) 回读断言（单元格值、表/透视表/图表/命名区域存在性）
  8) 出一份报告（stdout 可读版 + JSON 机读版）

用法示例：
  python run_vba.py --workbook "C:/path/x.xlsm" --code Report=vba_report.txt \
      --run BuildSummaryReport \
      --expect "cell:校验!B4=0" --expect "cell:校验!B5=通过" \
      --expect "pivot:汇总=1" --expect "chart:汇总=1" --expect "named:收入数据=存在"

断言语法（--expect 可重复）：
  cell:表!A1=值        单元格等于（数字按数值比，其余按文本比）
  cell:表!A1!=值       不等于
  cell:表!A1~子串     包含
  cell:表!A1!~子串    不包含
  sheet:表名=存在/不存在
  pivot:表名=数量      该表透视表个数
  chart:表名=数量      该表图表个数
  named:名称=存在/不存在

退出码：0=全部通过；1=有失败（断言/运行/卡死）；2=被拦（危险代码或参数错）
"""
import argparse
import json
import locale
import os
import re
import subprocess
import sys
import time

import psutil

try:
    import win32com.client as w32
    import win32process
except ImportError:
    print("需要 pywin32（本机 venv 已有）")
    sys.exit(2)

HERE = os.path.dirname(os.path.abspath(__file__))
GUARD = os.path.join(HERE, "vba_guard.py")
STRIP_ATTR = re.compile(r"(?im)^\s*Attribute\b.*$")
DANGEROUS = [
    (re.compile(r"\bMsgBox\b"), "MsgBox（弹窗=挂死）"),
    (re.compile(r"\bInputBox\b"), "InputBox（弹窗=挂死）"),
    (re.compile(r"(?im)^\s*Stop\s*$"), "Stop 语句（会把自动化实例拖进 VBE 断点=假死）"),
    (re.compile(r"\bDebug\.Assert\b"), "Debug.Assert（同上）=假死）"),
    (re.compile(r"\.\s*Show\b"), "UserForm/.Show（模态窗口=挂死）"),
]

# ---- 通用能力①：输出永不崩 -------------------------------------------------
# Windows 控制台默认编码可能是 cp936/cp932 等；此时打印 ⚠ ✅ 这类符号会抛
# UnicodeEncodeError。屏幕上看着正常，一旦 `> out.txt` 或 `| Select-String`
# 就整个崩掉（极难排查），所以这里固定 UTF-8 + errors=replace。
ASCII_MAP = {"✅": "[OK]", "❌": "[FAIL]", "⚠": "[!]", "⛔": "[BLOCK]",
             "↻": "[retry]", "ℹ": "[i]", "•": "-", "–": "-", "…": "..."}
ASCII_ONLY = False


def _setup_stdio():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


# ---- 通用能力②：源码编码自动识别 -------------------------------------------
# VBA 导出的 .bas/.frm/.cls 编码取决于导出环境：新版 UTF-8(常带 BOM)、
# 老版中文 GBK(CP936)、日文 CP932、西欧 CP1252。只认 UTF-8 会让这些文件
# 直接读不了（UnicodeDecodeError）。顺序：显式指定 > BOM > UTF-8 > 系统 ANSI > 常见候选。
ANSI_CANDIDATES = ("cp936", "cp932", "cp949", "cp1252", "latin-1")


def read_source(path, forced=None):
    """读源码文件 → (文本, 实际使用的编码)。"""
    with open(path, "rb") as f:
        raw = f.read()
    if forced:
        return raw.decode(forced), forced
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    order = ["utf-8"]
    try:
        pref = locale.getpreferredencoding(False)
        if pref and pref.lower().replace("-", "") not in ("utf8",):
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
    raise last


# ---- 通用能力③：只听"正事"，不被注释/字符串误伤 ----------------------------
def code_only(src):
    """剥掉注释与字符串字面量，只留真正的代码。

    为什么需要：注释里写一句 `' 千万别用 Me.Show`、或字符串里写 "(Me.Show 0)"
    都不是真的调用，不该被危险语句拦截器误伤（假阳性会白白折腾一轮）。
    规则：双引号内整体丢弃（"" 视为转义的一个引号）；行内 ' 之后丢弃；
    行首 Rem 之后丢弃。够用即止，不追求完整词法分析。
    """
    out = []
    for line in src.splitlines():
        if re.match(r"(?i)^\s*rem\b", line):
            out.append("")
            continue
        buf, i, n, in_str = [], 0, len(line), False
        while i < n:
            ch = line[i]
            if in_str:
                if ch == '"':
                    if i + 1 < n and line[i + 1] == '"':
                        i += 2
                        continue
                    in_str = False
                i += 1
                continue
            if ch == '"':
                in_str = True
                buf.append(" ")        # 占位，避免两侧 token 粘连
                i += 1
                continue
            if ch == "'":
                break
            buf.append(ch)
            i += 1
        out.append("".join(buf))
    return "\n".join(out)


def log(msg):
    if ASCII_ONLY:
        for k, v in ASCII_MAP.items():
            msg = msg.replace(k, v)
        msg = msg.encode("ascii", "replace").decode("ascii")
    print(msg, flush=True)


def parse_expect(spec):
    """拆成 (kind, target, op, value)。"""
    m = re.match(r"^\s*(\w+)\s*:\s*(.+?)\s*(!=|!~|~|=)\s*(.*)$", spec)
    if not m:
        raise ValueError("断言格式不对: %r" % spec)
    return m.group(1).lower(), m.group(2).strip(), m.group(3), m.group(4).strip()


def same(a, b):
    if a is None:
        return b in ("", "None", "None值")
    sa, sb = str(a).strip(), str(b).strip()
    try:
        return abs(float(sa) - float(sb)) < 1e-6
    except ValueError:
        return sa == sb


def get_active_excel():
    try:
        return w32.GetActiveObject("Excel.Application")
    except Exception:
        return None


def find_open_workbook(xl, path):
    try:
        n = xl.Workbooks.Count
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            wb = xl.Workbooks.Item(i)
            if os.path.normcase(os.path.abspath(str(wb.FullName))) == os.path.normcase(os.path.abspath(path)):
                return wb
        except Exception:
            pass
    return None


def inject(vbp, mod, code):
    try:
        vbp.VBComponents.Remove(vbp.VBComponents(mod))
    except Exception:
        pass
    c = vbp.VBComponents.Add(1)
    c.Name = mod
    c.CodeModule.AddFromString(code)
    return c


def resolve_run_name(xl, wb, name):
    """裸宏名是拿「当前活动工作簿」解析的。

    如果工作簿窗口在磁盘上就是隐藏的（常见的"装载器/宿主工作簿"设计：打开就跑程序、
    不让用户被这个簿子干扰），打开后 `ActiveWorkbook` 为 None，`Run("裸名")` 必然失败；
    而报错文本与 Attribute 行事故**一字不差**（「无法运行"X"宏。可能是因为该宏在此工作簿
    中不可用，或者所有的宏都被禁用」），照 Attribute 的方向查会一无所获。

    这种情况**是设计，不是故障**，不要为了迁就去改造工作簿；这里自动补成
    "工作簿名!宏名" 即可。返回 (实际使用的名字, 是否被自动限定)。
    """
    if "!" in name:
        return name, False
    try:
        if xl.ActiveWorkbook is not None:
            return name, False
    except Exception:
        pass
    try:
        wb_name = wb.Name
    except Exception:
        return name, False
    return "%s!%s" % (wb_name, name), True


def read_cell(wb, spec):
    """spec 形如 表!A1"""
    if "!" not in spec:
        raise ValueError("单元格断言要写成 表!A1")
    sheet, addr = spec.split("!", 1)
    return wb.Sheets(sheet).Range(addr).Value


def count_collection(wb, sheet, attr):
    """取透视表/图表的数量。

    注意：win32com 动态分发下 `ws.PivotTables` 拿到的是 function（不是集合），
    直接 `.Count` 会报 "'function' object has no attribute 'Count'"；
    所以依次尝试 [.Count, 调用后再 .Count, len(...)]。
    """
    ws = wb.Sheets(sheet)
    raw = getattr(ws, attr)
    last = None
    for fn in (lambda: raw.Count,
               lambda: raw().Count,
               lambda: len(raw()),
               lambda: len(raw)):
        try:
            return fn()
        except Exception as e:
            last = e
    raise last


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--workbook", required=True, help="目标 .xlsm（不存在则新建）")
    ap.add_argument("--code", action="append", required=True, metavar="[模块名=]文件",
                    help="VBA 代码文件；可重复。模块名默认取文件名（去扩展名）")
    ap.add_argument("--run", required=True, help="要运行的宏名")
    ap.add_argument("--expect", action="append", default=[], help="断言，可重复")
    ap.add_argument("--budget", type=float, default=60, help="守卫总预算秒（默认 60）")
    ap.add_argument("--hang-after", type=float, default=6, help="无响应多久判假死（默认 6s）")
    ap.add_argument("--hard-budget", type=float, default=45, help="跑飞判定的硬预算（默认 45s）")
    ap.add_argument("--visible", action="store_true", help="Excel 可见（默认隐藏）")
    ap.add_argument("--allow-events", action="store_true",
                    help="允许触发工作簿事件（默认禁用：装载器工作簿的 Workbook_Open 不会被顺带跑起来）")
    ap.add_argument("--save", action="store_true", help="运行后保存工作簿")
    ap.add_argument("--keep-open", action="store_true", help="结束后保留 Excel 与工作簿")
    ap.add_argument("--allow-unsafe", action="store_true", help="放行危险语句（默认拒绝）")
    ap.add_argument("--report", default=None, help="JSON 报告输出路径")
    ap.add_argument("--no-guard", action="store_true", help="不挂守卫（不推荐）")
    ap.add_argument("--encoding", default=None,
                    help="源码文件编码（默认自动识别：BOM→UTF-8→系统ANSI→cp936/cp932/cp949/cp1252）")
    ap.add_argument("--cleanup", action="store_true",
                    help="结束前移除本次注入的模块（处理正式工作簿时建议加）")
    ap.add_argument("--ascii", action="store_true", help="输出只用 ASCII 字符（任何终端都不乱码/不崩）")
    args = ap.parse_args()
    global ASCII_ONLY
    ASCII_ONLY = bool(args.ascii)
    _setup_stdio()

    wb_path = os.path.abspath(args.workbook)
    report = {"workbook": wb_path, "macro": args.run, "verdict": None, "run_return": None,
              "elapsed_s": None, "error": None, "guard": [], "assertions": [], "code_scan": [],
              "reused_instance": False, "saved": False, "excel_alive": None,
              "run_name": None, "run_name_qualified": None, "run_retried_qualified": False,
              "presaved": False, "cleaned_modules": None, "left_modules": None}
    problems = []

    # ---- 1) 读代码 + 剥 Attribute + 静态检查 ----
    modules = {}
    for spec in args.code:
        if "=" in spec and not os.path.exists(spec):
            mod, path = spec.split("=", 1)
        else:
            path = spec
            base = os.path.basename(path)
            mod = re.sub(r"\.(bas|vba|txt)$", "", base, flags=re.I)
            mod = re.sub(r"\.bas$", "", mod, flags=re.I)
        if not os.path.exists(path):
            log("✗ 代码文件不存在: %s" % path)
            return 2
        try:
            raw_text, enc = read_source(path, args.encoding)
        except UnicodeDecodeError as e:
            log("✗ 读不了源码文件（编码识别失败）: %s" % e)
            log("  → 用 --encoding 指定编码，例如 --encoding gbk / cp932 / utf-8-sig")
            return 2
        if enc not in ("utf-8", "utf-8-sig"):
            log("ℹ 源码编码识别为 %s（非 UTF-8）: %s" % (enc, os.path.basename(path)))
        raw = raw_text.replace("\r\n", "\n").replace("\n", "\r\n")
        clean = STRIP_ATTR.sub("", raw)
        code = code_only(clean)                  # 只听"正事"：注释/字符串不算
        hits = [why for rx, why in DANGEROUS if rx.search(code)]
        raw_hits = [why for rx, why in DANGEROUS if rx.search(clean)]
        report["code_scan"].append({"module": mod, "file": path, "encoding": enc,
                                    "dangerous": hits,
                                    "dangerous_in_comments_only": sorted(set(raw_hits) - set(hits))})
        if STRIP_ATTR.search(raw):
            report["code_scan"][-1]["stripped_attribute_lines"] = len(STRIP_ATTR.findall(raw))
        if hits and not args.allow_unsafe:
            problems.append("模块 %s 含危险语句: %s" % (mod, "、".join(hits)))
        elif raw_hits and not hits:
            log("ℹ 危险字样只出现在注释/字符串里 → 放行: %s" % "、".join(sorted(set(raw_hits))))
        modules[mod] = clean
    if problems:
        log("⛔ 拒绝注入（用 --allow-unsafe 可强制，但很可能挂死）：")
        for p in problems:
            log("   - " + p)
        report["verdict"] = "blocked"
        report["error"] = "；".join(problems)
        _write_report(args, report)
        return 2

    # ---- 2) 拿 Excel 实例（复用优先）----
    xl = get_active_excel()
    wb = None
    if xl is not None:
        wb = find_open_workbook(xl, wb_path)
        report["reused_instance"] = wb is not None
    if xl is None:
        xl = w32.DispatchEx("Excel.Application")
    prev_enable_events = None
    try:
        prev_enable_events = bool(xl.EnableEvents)
    except Exception:
        pass
    try:
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1
        xl.Visible = bool(args.visible)
        if not args.allow_events:
            # 装载器/宿主工作簿"一打开就跑程序"：测试工具绝不能顺带把它跑起来
            xl.EnableEvents = False
            log("事件已禁用（EnableEvents=False）：打开工作簿不会触发 Workbook_Open；要触发请加 --allow-events")
    except Exception:
        pass
    if wb is None:
        if os.path.exists(wb_path):
            wb = xl.Workbooks.Open(wb_path)
        else:
            wb = xl.Workbooks.Add()
            wb.SaveAs(wb_path, FileFormat=52)
    xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    log("Excel: %s | pid=%s | 工作簿=%s" % ("复用已开实例" if report["reused_instance"] else "新建实例",
                                           xl_pid, wb.Name))

    # ---- 2.5) 先保存"用户数据"，再注入 ----
    # 顺序很关键：保存放在注入之前，测试模块就不会被动写进磁盘。
    # （原来放在注入之后 → 即使不加 --save，注入的模块也会被写进用户文件）
    try:
        wb.Save()
        report["saved"] = True
        report["presaved"] = True
    except Exception as e:
        log("⚠ 运行前保存失败: %r" % (e,))

    # ---- 3) 注入 ----
    for mod, code in modules.items():
        inject(wb.VBProject, mod, code)
        log("注入模块: %s" % mod)

    # ---- 4) 起守卫 ----
    guard = None
    if not args.no_guard:
        guard = subprocess.Popen(
            [sys.executable, "-u", GUARD, str(xl_pid), str(args.budget), "0.1",
             "--hang_after=%s" % args.hang_after, "--hard_budget=%s" % args.hard_budget],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
            encoding="utf-8", errors="replace")
        time.sleep(0.6)          # 预热，确保 Run 之前已在扫描

    # ---- 5) 运行前保存已上移到 2.5（放在注入之前，避免测试模块被动写盘）----

    # ---- 6) 运行 ----
    t0 = time.time()
    run_exc = None
    run_name, auto_qualified = resolve_run_name(xl, wb, args.run)
    report["run_name"] = run_name
    if auto_qualified:
        report["run_name_qualified"] = run_name
        log("⚠ 工作簿窗口为隐藏态（ActiveWorkbook 为空）→ 自动改用限定名: %s" % run_name)
        log("  （装载器/宿主类工作簿常故意隐藏窗口，属设计而非故障，不改文件）")
    try:
        xl.Run(run_name)
        report["run_return"] = "ok"
    except Exception as e:
        msg = str(e)
        if (not auto_qualified) and any(
                k in msg for k in ("无法运行", "Cannot run the macro",
                                   "宏在此工作簿中不可用", "may be disabled")):
            retry = "%s!%s" % (wb.Name, args.run)
            log("↻ 裸宏名解析失败 → 自动重试限定名: %s" % retry)
            try:
                xl.Run(retry)
                report["run_name"] = retry
                report["run_name_qualified"] = retry
                report["run_retried_qualified"] = True
                report["run_return"] = "ok（自动重试限定名成功）"
            except Exception as e2:
                run_exc = e2
                report["run_return"] = "exception: %s" % (repr(e2)[:300],)
        else:
            run_exc = e
            report["run_return"] = "exception: %s" % (repr(e)[:300],)
    report["elapsed_s"] = round(time.time() - t0, 3)

    # ---- 7) 收守卫（给宽限期，让它把 DISMISSED_BY_CLICK / VERDICT 打完）----
    guard_out = ""
    if guard is not None:
        if guard.poll() is None:
            for _ in range(15):
                if guard.poll() is not None:
                    break
                time.sleep(0.1)
        if guard.poll() is None:
            guard.terminate()
            try:
                guard_out = guard.communicate(timeout=10)[0]
            except Exception:
                guard.kill()
                guard_out = guard.communicate()[0]
        else:
            guard_out = guard.communicate(timeout=5)[0]
        report["guard"] = [l for l in guard_out.splitlines()]
    gtxt = "\n".join(report["guard"])
    dialog = ("DISMISSED_BY_CLICK" in gtxt) or ("DIALOG " in gtxt)
    errtext = None
    for l in report["guard"]:
        if l.startswith("ERRTEXT "):
            try:
                errtext = json.loads(l[len("ERRTEXT "):])
            except Exception:
                errtext = l[len("ERRTEXT "):]
    killed = "KILLED" in gtxt
    run_ok = run_exc is None

    # 无弹窗的「无法运行宏」= 宏名解析不到 / 工程编译不过，与"运行时错误弹窗"不是一类
    macro_unavailable = bool(run_exc) and any(
        k in str(run_exc) for k in ("无法运行", "Cannot run the macro",
                                    "宏在此工作簿中不可用", "宏可能被禁用", "may be disabled"))
    if dialog:
        report["verdict"] = "vba_error_dialog_dismissed"
    elif macro_unavailable and not report.get("run_retried_qualified"):
        report["verdict"] = "macro_not_runnable"
    elif killed:
        kind = "idle_hung" if "IDLE_HUNG" in gtxt else ("runaway_cpu" if "RUNAWAY_CPU" in gtxt else "killed")
        report["verdict"] = kind
    elif run_exc is not None:
        report["verdict"] = "run_exception"
    else:
        report["verdict"] = "ok"
    report["error"] = errtext or (str(run_exc)[:300] if run_exc else None)
    try:
        report["excel_alive"] = any(p.pid == xl_pid for p in psutil.process_iter(["pid"]))
    except Exception:
        report["excel_alive"] = None

    # ---- 7.5) 断言前强制重算：否则新写的公式可能读到未计算的陈旧值（假失败）----
    if report["excel_alive"]:
        done = False
        for meth in ("CalculateFull", "CalculateFullRebuild", "Calculate"):
            try:
                getattr(xl, meth)()
                done = True
                break
            except Exception:
                continue
        if not done:
            try:
                wb.Application.Calculate()
                done = True
            except Exception as e:
                if psutil.pid_exists(xl_pid):      # 进程已死（刚被守卫杀掉）就别报噪音
                    log("⚠ 强制重算失败: %r（公式断言可能读到陈旧值）" % (e,))
        if done:
            time.sleep(0.15)

    # ---- 8) 断言 ----
    fails = []
    if report["excel_alive"] and wb is not None:
        for spec in args.expect:
            try:
                kind, target, op, val = parse_expect(spec)
            except ValueError as e:
                report["assertions"].append({"spec": spec, "ok": False, "note": str(e)})
                fails.append(spec)
                continue
            ok, actual, note = None, None, ""
            try:
                if kind == "cell":
                    actual = read_cell(wb, target)
                    if op == "=":
                        ok = same(actual, val)
                    elif op == "!=":
                        ok = not same(actual, val)
                    elif op == "~":
                        ok = val in str(actual)
                    else:
                        ok = val not in str(actual)
                elif kind in ("sheet", "named"):
                    if kind == "sheet":
                        names = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
                        exists = target in names
                        if not exists and val in ("存在", "exists", "yes", "1"):
                            note = "现有工作表: %s" % names
                    else:
                        names = [wb.Names.Item(i + 1).Name for i in range(wb.Names.Count)]
                        exists = target in names
                    want = val in ("存在", "exists", "yes", "1", "true", "True")
                    ok = (exists == want)
                    actual = "存在" if exists else "不存在"
                elif kind in ("pivot", "chart"):
                    attr = {"pivot": "PivotTables", "chart": "ChartObjects"}[kind]
                    actual = count_collection(wb, target, attr)
                    ok = same(actual, val)
                else:
                    note = "未知断言类型 %s" % kind
                    ok = False
            except Exception as e:
                ok = False
                note = "读取失败: %r" % (e,)
            report["assertions"].append({"spec": spec, "ok": bool(ok), "actual": actual, "note": note})
            if not ok:
                fails.append(spec)

    # ---- 9) 清理 / 保存 / 收尾 ----
    if prev_enable_events is not None:
        try:
            xl.EnableEvents = prev_enable_events      # 复用实例时把事件开关还原，别留在用户会话里
        except Exception:
            pass
    injected = list(modules.keys())
    if args.cleanup and report["excel_alive"]:
        removed = []
        for mod in injected:
            try:
                wb.VBProject.VBComponents.Remove(wb.VBProject.VBComponents(mod))
                removed.append(mod)
            except Exception:
                pass
        report["cleaned_modules"] = removed
        log("已移除本次注入的模块: %s" % ("、".join(removed) or "（无）"))
    else:
        report["left_modules"] = injected
    if args.save and report["excel_alive"]:
        try:
            wb.Save()
            report["saved"] = True
        except Exception as e:
            log("⚠ 收尾保存失败: %r" % (e,))
    if not args.keep_open and report["excel_alive"] and not report["reused_instance"]:
        try:
            try:
                wb.Close(bool(args.save))     # 说好不存就不存（原来是无条件 Close(True)，会静默写入测试模块）
            except Exception:
                pass
            xl.Quit()
        except Exception:
            pass

    # ---- 10) 打印报告 ----
    log("")
    log("=== 判定 ===")
    verdict_cn = {
        "ok": "✅ 正常运行完成",
        "vba_error_dialog_dismissed": "❌ VBA 报错（弹窗已自动点掉，Excel 存活）",
        "vba_error_dialog": "❌ VBA 报错",
        "macro_not_runnable": "❌ 宏无法运行（无弹窗：宏名解析不到 / 工程编译不过）",
        "idle_hung": "❌ 卡死（无弹窗、零 CPU）→ 已杀进程",
        "runaway_cpu": "❌ 跑飞（无弹窗、CPU 拉满）→ 已杀进程",
        "killed": "❌ 被守卫终止",
        "run_exception": "❌ COM 异常",
        "blocked": "⛔ 被拦（危险代码）",
    }.get(report["verdict"], report["verdict"])
    log("结论: %s" % verdict_cn)
    log("耗时: %ss | 宏返回: %s" % (report["elapsed_s"], report["run_return"]))
    if report.get("run_name_qualified"):
        log("宏名: 已自动限定为 %s（%s）" % (
            report["run_name_qualified"],
            "裸名失败后重试成功" if report.get("run_retried_qualified") else "窗口隐藏态，属设计，未改文件"))
    if report["error"]:
        log("错因: %s" % report["error"])
    if dialog:
        log("守卫: 弹窗已点掉（Excel 存活=%s）" % report["excel_alive"])
    elif killed:
        log("守卫: %s" % [l for l in report["guard"] if "VERDICT" in l][-1:])
    if report["assertions"]:
        log("")
        log("=== 断言 ===")
        for a in report["assertions"]:
            log("  %s %s%s" % ("✅" if a["ok"] else "❌", a["spec"],
                               "" if a["ok"] else "  (实际: %r %s)" % (a.get("actual"), a.get("note", ""))))
    ok_all = (report["verdict"] == "ok") and not fails
    if ok_all:
        summary = "全部通过 ✅"
    elif report["verdict"] != "ok":
        summary = "运行未成功 ❌（判定: %s）%s" % (
            report["verdict"], "，另有 %d 条断言未过" % len(fails) if fails else "")
    else:
        summary = "断言未通过 ❌（%d 条）" % len(fails)
    log("")
    log("=== 总结: %s ===" % summary)
    if report.get("left_modules"):
        log("注意: 本工作簿现留有以下测试模块: %s（加 --cleanup 可在结束后自动移除）"
            % "、".join(report["left_modules"]))
    report["ok"] = ok_all
    _write_report(args, report)
    return 0 if ok_all else 1


def _write_report(args, report):
    p = args.report or (os.path.splitext(os.path.abspath(args.workbook))[0] + ".runreport.json")
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        log("报告: %s" % p)
    except Exception as e:
        log("⚠ 报告写入失败: %r" % (e,))


if __name__ == "__main__":
    sys.exit(main())
