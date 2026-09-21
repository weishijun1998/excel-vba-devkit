"""只读批量巡检：.xlsm/.xlsx 是否处于"保存时窗口隐藏"状态。

为什么需要它：工作簿窗口隐藏会让 Application.Run("裸宏名") 必然失败，
报出与"代码带 Attribute 行"一字不差的"无法运行宏…可能所有的宏都被禁用"；
且该状态是文件级的，会随"另存为"传染给所有副本。

本脚本不启动 Excel、不修改任何文件，只看 zip 里的 xl/workbook.xml：
  正常保存  -> <workbookView ...>                     （无 visibility 属性）
  隐藏保存  -> <workbookView visibility="hidden" ...>

用法:
    python scan_hidden_windows.py <目录或文件> [更多路径...]

退出码: 0=没有隐藏文件  1=发现隐藏文件  2=用法错误
"""
import os
import re
import sys
import zipfile

VIEW_RX = re.compile(rb"<workbookView[^>]*>")
EXTS = (".xlsm", ".xlsx", ".xltm", ".xltx")


def check(path):
    """返回 (状态, 说明)，状态为 hidden / visible / unknown。"""
    try:
        with zipfile.ZipFile(path) as z:
            if "xl/workbook.xml" not in z.namelist():
                return "unknown", "缺 xl/workbook.xml（不是 xlsx/xlsm 容器）"
            xml = z.read("xl/workbook.xml")
    except zipfile.BadZipFile:
        return "unknown", "不是 zip 容器（可能是老的 .xls 二进制格式）"
    except Exception as e:  # noqa: BLE001
        return "unknown", "读取失败 %s" % type(e).__name__
    m = VIEW_RX.search(xml)
    if not m:
        return "unknown", "workbook.xml 里没有 workbookView 节点"
    if b'visibility="hidden"' in m.group(0):
        return "hidden", "visibility=hidden"
    return "visible", ""


def iter_targets(paths):
    for p in paths:
        if os.path.isfile(p):
            if p.lower().endswith(EXTS):
                yield p
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = [d for d in dirs if d not in ("~$", ".git", "__pycache__")]
                for f in files:
                    if f.startswith("~$") or not f.lower().endswith(EXTS):
                        continue
                    yield os.path.join(root, f)


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 2

    hidden, unknown, checked = [], [], 0
    for path in sorted(set(iter_targets(args))):
        state, note = check(path)
        checked += 1
        if state == "hidden":
            hidden.append(path)
            print("隐藏  %s" % path)
        elif state == "unknown":
            unknown.append((path, note))
        else:
            print("可见  %s" % path)

    print("\n扫描 %d 个文件：隐藏 %d 个" % (checked, len(hidden)))
    for path, note in unknown:
        print("  (跳过) %s — %s" % (path, note))

    if hidden:
        print("\n这些文件用裸宏名跑必然失败（报 0x800A03EC 文案，外层码可能是 0x80020009）：")
        print("  立刻绕过：Application.Run(\"文件名.xlsm!宏名\")，或 run_vba.py 的限定名/自动补限定名")
        print("  根治（会改文件，先取得用户同意）：wb.Windows(1).Visible = True 后 Save")
        print("  注意：该状态会随另存为传染，修复后要一并处理其副本/备份")
    return 1 if hidden else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
