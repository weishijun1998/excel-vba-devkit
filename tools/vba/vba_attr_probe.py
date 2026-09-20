# -*- coding: utf-8 -*-
"""VBA 注入探针：Attribute 行与 0x800A03EC 的关系（2026-09-20 实测）。

实测结论：
  A 干净代码                        -> 注入+运行成功
  B 代码带 Attribute VB_Name        -> 注入成功但 Run 报 0x800A03EC（文案甩锅"宏被禁用"）
  C 坏模块在场时再注入干净模块       -> 一样跑不了（整个 VBA 工程无法编译）
  D 把坏模块的代码就地覆写成干净代码 -> 恢复可运行
  另测：Import 带 attribute 的 .bas、过程级 Attribute(VB_ProcData) 同样必失败。

=> 规则：注入前剥掉所有 ^\\s*Attribute 行（模块名交给 API/参数）；
        已中招的模块必须就地覆写或删除，光在旁边加好模块没用。

用法：python vba_attr_probe.py            （输出到 %TEMP%）
注意：git-bash 下不要传路径参数，MSYS 会把它转成 C:\\\\//... 导致 SaveAs 失败。
前置：Excel + pywin32；信任中心已勾选"信任对 VBA 工程对象模型的访问"。
"""

import os
import re
import sys

import win32com.client as w32

STRIP = re.compile(r"(?im)^\s*Attribute\b.*$")
OUT = os.environ.get("TEMP", ".")
PATH = os.path.join(OUT, "vba_attr_probe.xlsm")

CLEAN = 'Sub WriteClean()\r\n    Range("A1").Value = "clean-ok"\r\nEnd Sub\r\n'
CLEAN2 = 'Sub WriteClean2()\r\n    Range("A3").Value = "clean2-ok"\r\nEnd Sub\r\n'
DIRTY = ('Attribute VB_Name = "ModAttr"\r\n\r\n'
         'Sub WriteAttr()\r\n    Range("A2").Value = "attr-ok"\r\nEnd Sub\r\n')
DIRTY_FIXED = 'Sub WriteAttr()\r\n    Range("A2").Value = "fixed-ok"\r\nEnd Sub\r\n'

xl = w32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = 1

if os.path.exists(PATH):
    os.remove(PATH)
wb = xl.Workbooks.Add()
wb.SaveAs(PATH, FileFormat=52)   # xlOpenXMLWorkbookMacroEnabled
wb.Close(False)
wb = xl.Workbooks.Open(PATH)     # 关键：从磁盘重开，模拟自动化实例
vbp = wb.VBProject
rows = []


def run_proc(proc):
    try:
        xl.Run(proc)
        return "运行成功"
    except Exception as e:
        return "运行失败(0x800A03EC)" if "0x800A03EC" in str(e) or "800A03EC" in repr(e) or "宏" in str(e) else "运行失败 %s" % (e,)


def add_module(name, code):
    m = vbp.VBComponents.Add(1)          # vbext_ct_StdModule
    m.Name = name
    m.CodeModule.AddFromString(code)
    return m


# A 干净代码
try:
    add_module("ModClean", CLEAN)
    rows.append(("A 干净代码", run_proc("WriteClean")))
except Exception as e:
    rows.append(("A 干净代码", "注入失败 %s" % (e,)))

# B 带 Attribute 行
try:
    dirty = add_module("ModAttr", DIRTY)
    rows.append(("B 带 Attribute", run_proc("WriteAttr")))
except Exception as e:
    dirty = None
    rows.append(("B 带 Attribute", "注入失败 %s" % (e,)))

# C 坏模块在场时新加的干净模块
try:
    add_module("ModClean2", CLEAN2)
    rows.append(("C 坏模块在场+新干净模块", run_proc("WriteClean2")))
except Exception as e:
    rows.append(("C 坏模块在场+新干净模块", "注入失败 %s" % (e,)))

# D 就地覆写坏模块为干净代码
if dirty is not None:
    try:
        dirty.CodeModule.DeleteLines(1, dirty.CodeModule.CountOfLines)
        dirty.CodeModule.AddFromString(STRIP.sub("", DIRTY_FIXED))
        rows.append(("D 就地覆写坏模块", run_proc("WriteAttr")))
    except Exception as e:
        rows.append(("D 就地覆写坏模块", "覆写失败 %s" % (e,)))

vals = [wb.Sheets(1).Range("A1").Value, wb.Sheets(1).Range("A2").Value, wb.Sheets(1).Range("A3").Value]
for n, v in rows:
    print("%-24s | %s" % (n, v))
print("单元格 A1/A2/A3 =", vals)
print("[done]", PATH)

try:
    wb.Close(False)
    xl.Quit()
except Exception:
    pass
