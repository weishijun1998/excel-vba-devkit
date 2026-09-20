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


def log(msg):
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
    ap.add_argument("--save", action="store_true", help="运行后保存工作簿")
    ap.add_argument("--keep-open", action="store_true", help="结束后保留 Excel 与工作簿")
    ap.add_argument("--allow-unsafe", action="store_true", help="放行危险语句（默认拒绝）")
    ap.add_argument("--report", default=None, help="JSON 报告输出路径")
    ap.add_argument("--no-guard", action="store_true", help="不挂守卫（不推荐）")
    args = ap.parse_args()

    wb_path = os.path.abspath(args.workbook)
    report = {"workbook": wb_path, "macro": args.run, "verdict": None, "run_return": None,
              "elapsed_s": None, "error": None, "guard": [], "assertions": [], "code_scan": [],
              "reused_instance": False, "saved": False, "excel_alive": None,
              "run_name": None, "run_name_qualified": None, "run_retried_qualified": False}
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
        with open(path, encoding="utf-8") as f:
            raw = f.read().replace("\r\n", "\n").replace("\n", "\r\n")
        hits = []
        for rx, why in DANGEROUS:
            if rx.search(raw):
                hits.append(why)
        report["code_scan"].append({"module": mod, "file": path, "dangerous": hits})
        clean = STRIP_ATTR.sub("", raw)
        if STRIP_ATTR.search(raw):
            report["code_scan"][-1]["stripped_attribute_lines"] = len(STRIP_ATTR.findall(raw))
        if hits and not args.allow_unsafe:
            problems.append("模块 %s 含危险语句: %s" % (mod, "、".join(hits)))
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
    try:
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 1
        xl.Visible = bool(args.visible)
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

    # ---- 5) 运行前保存（被杀也能从磁盘继续）----
    try:
        wb.Save()
        report["saved"] = True
    except Exception as e:
        log("⚠ 运行前保存失败: %r" % (e,))

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

    # ---- 9) 保存/收尾 ----
    if args.save and report["excel_alive"]:
        try:
            wb.Save()
            report["saved"] = True
        except Exception as e:
            log("⚠ 收尾保存失败: %r" % (e,))
    if not args.keep_open and report["excel_alive"] and not report["reused_instance"]:
        try:
            try:
                wb.Close(True)
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
