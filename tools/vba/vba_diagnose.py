# -*- coding: utf-8 -*-
"""vba_diagnose.py — read-only diagnosis: why won't this workbook's VBA run?
（只读诊断：这个工作簿的 VBA 为什么跑不起来）

`run_vba.py` 报「无法运行"X"宏。可能是因为该宏在此工作簿中不可用，或者所有的宏都被禁用」时，
先用它分清病因 —— 这**同一句话**至少对应两类完全不同的原因：

  1) 工程层面：模块里有 `Attribute` 行、坏模块、坏引用、编译错误
     -> 剥掉 Attribute、坏模块就地覆写或删除、修引用
  2) 窗口层面：工作簿窗口**在磁盘上就是隐藏的**
     这是"装载器 / 宿主工作簿"的常见**设计**（打开就跑程序、不让用户被这个簿子干扰）。
     后果：打开后 `ActiveWorkbook` 为 None，而 `Run("裸宏名")` 是拿"活动工作簿"解析的
     -> 必然失败。**不要去改造这个文件**；自动化侧一律用限定名 `"工作簿名!宏名"`。

用法：
    python tools/vba/vba_diagnose.py "C:/path/x.xlsm"     # 单文件：静态检查 + COM 深查（要开 Excel）
    python tools/vba/vba_diagnose.py "C:/path/to/dir"     # 目录：只做静态检查（不开 Excel，秒级）
    python tools/vba/vba_diagnose.py "C:/path/x.xlsm" --json

只读：不注入、不修改、不保存。
退出码：0 未发现问题 ｜ 1 发现问题（窗口隐藏、Attribute 行、坏引用等）｜ 2 用法/路径错误
"""
import argparse
import json
import os
import re
import sys
import zipfile

ASCII_ONLY = False
ATTR_RX = re.compile(r"(?im)^\s*Attribute\b")
VIEW_RX = re.compile(r"<workbookView\b[^>]*>")
HIDDEN_RX = re.compile(r'visibility\s*=\s*"hidden"')
LOADER_RX = re.compile(r"(?im)^\s*(Private\s+|Public\s+)?Sub\s+(Workbook_Open|Auto_Open)\b")
MACRO_EXTS = (".xlsm", ".xltm", ".xlam")

# 未限定引用：窗口隐藏（没有活动工作簿）时，"裸"写法的对象引用会失败。
# 这些引用都经由 VBA 的 _Global 解析，而 _Global 依赖"当前活动的那个工作簿/工作表"。
UNQUALIFIED = [
    (re.compile(r"(?<!\.)\bSheets\s*\("), "Sheets("),
    (re.compile(r"(?<!\.)\bWorksheets\s*\("), "Worksheets("),
    (re.compile(r"(?<!\.)\bRange\s*\("), "Range("),
    (re.compile(r"(?<!\.)\bCells\s*\("), "Cells("),
    (re.compile(r"(?<!\.)\bColumns\s*\("), "Columns("),
    (re.compile(r"(?<!\.)\bRows\s*\("), "Rows("),
    (re.compile(r"(?<!\.)\bSelection\b"), "Selection"),
    (re.compile(r"\bActiveSheet\b"), "ActiveSheet"),
    (re.compile(r"\bActiveWorkbook\b"), "ActiveWorkbook"),
    (re.compile(r"\bActiveWindow\b"), "ActiveWindow"),
    (re.compile(r"\bApplication\s*\.\s*(?:Sheets|Worksheets|Range|Cells|ActiveSheet|ActiveWorkbook)\b"),
     "Application.xxx（同样依赖活动工作簿）"),
]


def _setup_stdio(ascii_only=False):
    """输出永不崩：Windows 控制台可能是 cp936/cp932，打印 ⚠ ✅ 会把整个程序打挂
    （屏幕上正常，一旦 `> out.txt` 或 `| Select-String` 就崩）。--ascii 时退化为纯 ASCII。"""
    enc = "ascii" if ascii_only else "utf-8"
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding=enc, errors="replace")
        except Exception:
            pass


def code_only(src):
    """剥掉注释与字符串字面量，只留真正的代码（避免把注释里的字样算成真调用）。"""
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
                buf.append(" ")
                i += 1
                continue
            if ch == "'":
                break
            buf.append(ch)
            i += 1
        out.append("".join(buf))
    return "\n".join(out)


def find_unqualified(code):
    """返回 [(行号, 命中文本)]。"""
    hits = []
    clean = code_only(code)
    for rx, what in UNQUALIFIED:
        for m in rx.finditer(clean):
            hits.append((clean[:m.start()].count("\n") + 1, what))
    return sorted(set(hits))


# --------------------------------------------------------------------------- 静态
def static_visibility(path):
    """不开 Excel，直接从 zip 里读 xl/workbook.xml 判断窗口隐藏态。

    返回 "hidden" / "visible" / None（读不到，例如 .xlsb 不是 zip、或加了密码）。
    对照实验已验证：隐藏保存后会出现 visibility="hidden"，恢复可见再保存就消失。
    """
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except Exception:
        return None
    m = VIEW_RX.search(xml)
    if not m:
        return None
    return "hidden" if HIDDEN_RX.search(m.group(0)) else "visible"


def scan_dir(d):
    """目录：只做静态扫描（不启动 Excel），列出 .xlsm/.xltm，标出隐藏窗口的。"""
    rows = []
    for root, dirs, files in os.walk(d):
        dirs[:] = [x for x in dirs if x not in (".git", "__pycache__", "$RECYCLE.BIN")]
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() not in MACRO_EXTS:
                continue
            p = os.path.join(root, f)
            rows.append({"file": p, "size_kb": round(os.path.getsize(p) / 1024, 1),
                         "window": static_visibility(p) or "unknown"})
    return rows


# --------------------------------------------------------------------------- COM 深查
def com_diagnose(path, name_hint=None):
    import win32com.client as w32

    findings = []
    info = {"active_workbook": None, "window_visible": None, "is_addin": None,
            "read_only": None, "sheets": [], "components": [], "attr_modules": [],
            "broken_refs": [], "loader": False}

    xl = w32.DispatchEx("Excel.Application")
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    xl.Visible = False
    xl.EnableEvents = False          # 诊断绝不能顺带触发装载器工作簿的 Workbook_Open（打开即跑程序）
    wb = None
    try:
        wb = xl.Workbooks.Open(path)
        print("=" * 78)
        print("文件:", path)

        print("\n[1] 工作簿 / 窗口状态        <-- 决定裸宏名能不能解析")
        try:
            info["read_only"] = bool(wb.ReadOnly)
            print("    ReadOnly            :", info["read_only"])
        except Exception:
            pass
        try:
            print("    ProtectedViewWindows:", xl.ProtectedViewWindows.Count)
        except Exception:
            pass
        try:
            info["is_addin"] = bool(wb.IsAddin)
            print("    IsAddin             :", info["is_addin"])
        except Exception:
            pass
        try:
            aw = xl.ActiveWorkbook
            info["active_workbook"] = None if aw is None else str(aw.Name)
            print("    ActiveWorkbook      :", info["active_workbook"] or "<None>")
        except Exception as e:  # noqa: BLE001
            print("    ActiveWorkbook      : <读取失败> %s" % type(e).__name__)
        try:
            n = wb.Windows.Count
            print("    Windows.Count       :", n)
            for i in range(1, n + 1):
                v = bool(wb.Windows(i).Visible)
                if i == 1:
                    info["window_visible"] = v
                print("      window[%d].Visible: %s" % (i, v))
        except Exception as e:  # noqa: BLE001
            print("    Windows             : 读取失败 %s" % type(e).__name__)

        stat = static_visibility(path)
        print("    文件内标注(静态)    :", {"hidden": "visibility=hidden", "visible": "无 visibility 属性"}
              .get(stat, "未知（非 zip 或读不到）"))

        hidden = info["window_visible"] is False or info["active_workbook"] is None
        info["hidden"] = hidden
        if hidden:
            print("    -> ⚠ 窗口隐藏（裸宏名会失败；宏内部未限定的引用也会失败，见结论）")
        else:
            print("    -> 窗口可见、有活动工作簿，裸宏名解析正常。")

        print("\n[2] 工作表")
        try:
            info["sheets"] = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
            print("    数量:", len(info["sheets"]), "->", info["sheets"])
        except Exception as e:  # noqa: BLE001
            print("    读取失败:", e)

        print("\n[3] VBA 工程")
        try:
            vbp = wb.VBProject
        except Exception as e:  # noqa: BLE001
            print("    不可访问:", e)
            print("    -> 需在信任中心勾选「信任对 VBA 工程对象模型的访问」")
            findings.append("无法访问 VBA 工程（信任中心未授权），工程层面的病因无法判断。")
            return info, findings
        try:
            print("    工程名:", vbp.Name, "| Protection:", vbp.Protection, "| Mode:", vbp.Mode)
        except Exception:
            pass

        print("\n[4] 引用（坏引用会拖垮整个工程）")
        try:
            total = vbp.References.Count
            for i in range(1, total + 1):
                try:
                    r = vbp.References.Item(i)
                    if r.IsBroken:
                        info["broken_refs"].append(str(r.Name))
                        print("    [坏] %-40s %s" % (r.Name, r.FullPath))
                except Exception:
                    pass
            print("    共 %d 个，坏引用 %d 个 %s" % (total, len(info["broken_refs"]),
                                                 info["broken_refs"] or ""))
        except Exception as e:  # noqa: BLE001
            print("    读取失败:", e)
        if info["broken_refs"]:
            findings.append("存在坏引用 %s -> 取消勾选或修复后工程才能编译。" % info["broken_refs"])

        print("\n[5] 组件 + Attribute 行 + 装载入口 + 未限定引用扫描")
        try:
            kind_map = {1: "标准模块", 2: "类模块", 3: "窗体", 100: "文档"}
            for i in range(1, vbp.VBComponents.Count + 1):
                c = vbp.VBComponents.Item(i)
                try:
                    cm = c.CodeModule
                    n = cm.CountOfLines
                    code = cm.Lines(1, n) if n else ""
                except Exception:
                    n, code = -1, ""
                lines = code.splitlines()
                hits = [k + 1 for k in range(len(lines)) if ATTR_RX.match(lines[k])]
                unq = find_unqualified(code)
                is_loader = bool(LOADER_RX.search(code))
                if hits:
                    info["attr_modules"].append({"module": str(c.Name), "lines": hits})
                if unq:
                    info.setdefault("unqualified", []).extend(
                        [{"module": str(c.Name), "line": ln, "what": w} for ln, w in unq])
                if is_loader:
                    info["loader"] = True
                info["components"].append({"module": str(c.Name),
                                           "type": kind_map.get(c.Type, str(c.Type)),
                                           "lines": n})
                print("    %-28s %-8s lines=%-5s %s%s%s" % (
                    c.Name, kind_map.get(c.Type, str(c.Type)), n,
                    "<== Attribute 行 @%s" % hits if hits else "",
                    "  <== 装载入口" if is_loader else "",
                    "  <== 未限定引用 %d 处" % len(unq) if unq else ""))
        except Exception as e:  # noqa: BLE001
            print("    读取失败:", e)
        if info["attr_modules"]:
            findings.append(
                "以下模块含 Attribute 行（只在导出的 .bas 里合法，当源码注入会编译失败，"
                "并拖垮整个工程）: %s -> 剥掉；坏模块必须**就地覆写或删除**。"
                % info["attr_modules"])

        if info.get("hidden"):
            wb_name = name_hint or getattr(wb, "Name", "<工作簿名>")
            if info.get("loader"):
                why = "本工作簿含 Workbook_Open/Auto_Open（打开即跑），隐藏窗口**很可能是有意设计**。"
            else:
                why = ("未发现「打开即跑」的入口，**无法断定**隐藏窗口是有意设计还是历史遗留"
                       "（例如用过「视图→隐藏」后保存）；两种情况的处置方式相同。")
            findings.append(
                "窗口隐藏 -> 没有「活动工作簿」-> \n"
                "      ① 外部用裸宏名调用必然失败；② 宏**内部**未限定的 Sheets/Range/Cells/ActiveSheet 也会失败。\n"
                "      " + why + "\n"
                "      处置：外部一律用限定名 \"%s!宏名\"（run_vba.py 已自动补）；"
                "脚本内部一律写 ThisWorkbook.Worksheets(...)；**不要为迁就去改造该文件**。" % wb_name)

        if info.get("unqualified"):
            sample = "、".join("%s:%d(%s)" % (u["module"], u["line"], u["what"])
                              for u in info["unqualified"][:10])
            findings.append(
                "发现 %d 处**未限定引用**（窗口隐藏时必失败；窗口可见时也可能因「活动对象不是你以为的那个」而出错）: %s"
                " -> 改成 ThisWorkbook.Worksheets(\"表名\") 或显式的工作簿/工作表变量。"
                % (len(info["unqualified"]), sample))

        print("\n" + "=" * 78)
        if findings:
            print("诊断结论:")
            for v in findings:
                print("  * " + v)
        else:
            print("未发现窗口 / 引用 / Attribute 层面的问题。")
            print("若仍报「无法运行宏」，依次排查：")
            print("  1) 工程内有编译错误（VBE 里点「调试->编译」看）")
            print("  2) 宏所在的模块被删/改名，或宏本身不存在")
            print("  3) 用全新空工作簿跑同一个宏，确认工具链本身正常（这一步能一刀排除工具链）")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            xl.Quit()
        except Exception:
            pass
    return info, findings


# --------------------------------------------------------------------------- 入口
def main():
    ap = argparse.ArgumentParser(add_help=True,
                                 description="只读诊断 Excel VBA 跑不起来的原因（窗口隐藏 / Attribute / 坏引用）")
    ap.add_argument("path", help=".xlsm/.xltm 文件，或一个目录（目录只做静态扫描）")
    ap.add_argument("--json", action="store_true", help="额外输出 JSON 摘要")
    ap.add_argument("--ascii", action="store_true", help="输出只用 ASCII 字符（任何终端都不乱码/不崩）")
    args = ap.parse_args()
    _setup_stdio(bool(args.ascii))

    path = os.path.abspath(args.path)
    if not os.path.exists(path):
        print("路径不存在:", path)
        return 2

    if os.path.isdir(path):
        rows = scan_dir(path)
        if not rows:
            print("目录下没有 .xlsm/.xltm 文件:", path)
            return 0
        hidden = [r for r in rows if r["window"] == "hidden"]
        print("=" * 78)
        print("静态扫描（未启动 Excel）:", path)
        print("  共 %d 个宏工作簿，其中窗口隐藏 %d 个" % (len(rows), len(hidden)))
        for r in rows:
            flag = "隐藏" if r["window"] == "hidden" else ("可见" if r["window"] == "visible" else "未知")
            print("    [%s] %8.1f KB  %s" % (flag, r["size_kb"], r["file"]))
        if hidden:
            print("\n  说明：隐藏窗口常见于**装载器/宿主工作簿**（设计如此），不是故障；")
            print("       但任何用裸宏名调用它们的脚本都会失败 —— 改用限定名 \"工作簿名!宏名\"，")
            print("       或直接用 run_vba.py（它会自动补限定名）。")
        if args.json:
            print(json.dumps({"mode": "dir", "rows": rows}, ensure_ascii=False, indent=2))
        return 1 if hidden else 0

    if os.path.splitext(path)[1].lower() not in MACRO_EXTS:
        print("⚠ 不是宏工作簿（期望 %s），仍会尝试诊断。" % "/".join(MACRO_EXTS))

    stat = static_visibility(path)
    print("静态检查（不开 Excel）:",
          {"hidden": "窗口隐藏（visibility=hidden）", "visible": "窗口可见"}
          .get(stat, "未知（非 zip、被加密或读不到）"))
    info, findings = com_diagnose(path)
    if info.get("loader"):
        print("提示：该工作簿含 Workbook_Open/Auto_Open —— 典型的「打开即跑」装载器，"
              "窗口隐藏多半是**有意设计**。")
    if args.json:
        print(json.dumps({"mode": "file", "file": path, "static_window": stat,
                          "info": info, "findings": findings}, ensure_ascii=False, indent=2))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
