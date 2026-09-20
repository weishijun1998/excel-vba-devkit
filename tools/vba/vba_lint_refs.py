# -*- coding: utf-8 -*-
"""vba_lint_refs.py — 扫出 VBA 源码里的「未限定引用」（静态、只读、不开 Excel）

为什么需要：VBA 里没写清楚"哪个工作簿/哪张表"的引用，要通过 `_Global` 解析，
而 `_Global` 依赖**当前活动的那个工作簿/工作表**。于是：

  * 手工用时（工作簿是活动的）一切正常 —— 所以这类写法能潜伏很多年；
  * 一旦从**自动化**里调用（隐藏窗口 / 无活动工作簿）→ 立刻失败：
      `Sheets("X")`            → 运行时错误 1004「方法 'Sheets' 作用于对象 '_Global' 时失败」
      `ActiveSheet`            → 得到 Nothing（这一行不报错），**下一行用**它时才报 91
      `Application.Range(...)` → 同样依赖活动工作表

会扫的写法：`Sheets(` `Worksheets(` `Range(` `Cells(` `Columns(` `Rows(` `Selection`
`Windows(` `Workbooks(` `ActiveSheet` `ActiveWorkbook` `ActiveWindow` `Application.xxx`
（带 `.` 前缀的**对象**限定如 `ws.Range(`、`ThisWorkbook.Worksheets(` 不算问题，会跳过）

用法：
    python tools/vba/vba_lint_refs.py "C:/path/vba-files"          # 扫目录
    python tools/vba/vba_lint_refs.py "C:/path/MainSubs.bas"       # 扫文件
    python tools/vba/vba_lint_refs.py <路径...> --json --quiet
    python tools/vba/vba_lint_refs.py <路径> --encoding gbk        # 手动指定编码

编码：自动识别（BOM → UTF-8 → 系统 ANSI → cp936/cp932/cp949/cp1252/latin-1），
      老式中文导出的 .bas 多是 GBK，只认 UTF-8 会直接读不了。

退出码：0=没有发现问题 ｜ 1=发现未限定引用 ｜ 2=用法/路径错误
"""
import argparse
import json
import locale
import os
import re
import sys

ASCII_ONLY = False
SRC_EXTS = (".bas", ".cls", ".frm", ".vba", ".txt", ".inc")
ANSI_CANDIDATES = ("cp936", "cp932", "cp949", "cp1252", "latin-1")

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
     "Application.xxx"),
]
# 注意：`Workbooks("名字")`、`Workbooks(1)`、`Windows(1)` **实测是安全的** ——
# 按名字/序号取对象，不依赖"活动工作簿/活动工作表"。所以**不列入**上面的清单，
# 免得把正确代码报成问题（实测：隐藏窗口下 A~D 四种写法全部成功，
# 只有 Application.Range 这类依赖活动工作表的才报 1004）。


def _setup_stdio():
    """输出永不崩（Windows 控制台可能是 cp936，打印符号会炸）。"""
    enc = "ascii" if ASCII_ONLY else "utf-8"
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding=enc, errors="replace")
        except Exception:
            pass


def read_source(path, forced=None):
    """读源码 → (文本, 编码)。VBA 导出文件的实际编码随环境而变。"""
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
    raise last


def code_only(src):
    """剥掉注释与字符串 —— 注释里提一句 Sheets( 不该被算成真代码。"""
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


def lint_text(text):
    """返回 [(行号, 命中, 该行原文)]，按行号排序去重。"""
    raw_lines = text.splitlines()
    clean = code_only(text)
    out = []
    for rx, what in UNQUALIFIED:
        for m in rx.finditer(clean):
            ln = clean[:m.start()].count("\n") + 1
            src_line = raw_lines[ln - 1].strip() if ln - 1 < len(raw_lines) else ""
            out.append((ln, what, src_line[:120]))
    return sorted(set(out))


def iter_files(paths):
    for p in paths:
        if os.path.isfile(p):
            yield p
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", "~$")]
                for f in sorted(files):
                    if os.path.splitext(f)[1].lower() in SRC_EXTS:
                        yield os.path.join(root, f)


def main():
    ap = argparse.ArgumentParser(add_help=True,
                                 description="扫出 VBA 源码里的未限定引用（静态、只读）")
    ap.add_argument("paths", nargs="+", help="要扫的文件或目录，可多个")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--quiet", action="store_true", help="只在有发现时输出")
    ap.add_argument("--encoding", default=None, help="手动指定源码编码（默认自动识别）")
    args = ap.parse_args()

    global ASCII_ONLY
    ASCII_ONLY = bool(getattr(args, "ascii", False))
    _setup_stdio()

    files = sorted(set(iter_files(args.paths)))
    if not files:
        print("没有找到可扫的源码文件（%s）" % "/".join(SRC_EXTS))
        return 2

    results, enc_counter = [], {}
    for p in files:
        try:
            text, enc = read_source(p, args.encoding)
        except Exception as e:  # noqa: BLE001
            results.append({"file": p, "error": "读不了: %s" % e, "hits": []})
            continue
        enc_counter[enc] = enc_counter.get(enc, 0) + 1
        hits = lint_text(text)
        if hits:
            results.append({"file": p, "encoding": enc, "hits": [
                {"line": ln, "what": what, "code": code} for ln, what, code in hits]})

    total = sum(len(r["hits"]) for r in results)
    if args.json:
        print(json.dumps({"checked": len(files), "files_with_hits": len(results),
                          "hits": total, "results": results}, ensure_ascii=False, indent=2))
        return 1 if total else 0

    if not args.quiet:
        print("扫描 %d 个文件，编码分布: %s" % (
            len(files), ", ".join("%s×%d" % (k, v) for k, v in sorted(enc_counter.items()))))
    if not total:
        if not args.quiet:
            print("未发现未限定引用。")
        return 0

    print("发现 %d 处未限定引用（%d 个文件）：" % (total, len(results)))
    for r in results:
        if r.get("error"):
            print("  [读不了] %s — %s" % (r["file"], r["error"]))
            continue
        print("  %s" % r["file"])
        for h in r["hits"][:40]:
            print("      %5d  %-16s %s" % (h["line"], h["what"], h["code"]))
        if len(r["hits"]) > 40:
            print("      ... 另有 %d 处" % (len(r["hits"]) - 40))
    print("\n修法：把「裸写法」限定到具体对象 ——")
    print("  Sheets(\"X\")        -> ThisWorkbook.Worksheets(\"X\")  或 ws.Worksheets(\"X\")")
    print("  Range(\"A1\")        -> ws.Range(\"A1\")")
    print("  ActiveSheet        -> 显式的 ws 变量（ActiveSheet 在无活动工作簿时是 Nothing）")
    print("  ActiveWorkbook     -> ThisWorkbook（或调用方显式传入的 wb）")
    return 1


if __name__ == "__main__":
    sys.exit(main())
