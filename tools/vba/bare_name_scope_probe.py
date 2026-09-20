# -*- coding: utf-8 -*-
"""bare_name_scope_probe.py —— 钉死"裸宏名到底哪种调用方式会失败"（可复现、不碰用户文件）

背景：Excel 的 `Application.Run` 用**裸宏名**时，是拿「当前活动工作簿」解析的。
工作簿窗口若是隐藏的（装载器/宿主工作簿常见设计），就没有活动工作簿。
但**同一个隐藏窗口**下，下面五种写法的结果完全不同 —— 这条边界很容易归因错
（"宏被禁用了" / "Attribute 行作祟" / "窗口隐藏" 会互相甩锅）。

本探针自己造一个隐藏窗口的临时装载器，依次跑五个案例并打印对照表：

  A) 外部调用 + 裸宏名          → 预期失败
  B) 外部调用 + 限定名          → 预期成功
  C) 内部 Application.Run 裸名  → 预期成功（走调用方工程的上下文）
  D) 内部 Application.Run 限定名 → 预期成功
  E) 内部直接 Call              → 预期成功

用法：python bare_name_scope_probe.py
退出码：0=结果与预期一致 ｜ 1=有偏差（说明 Excel 行为变了，需要重新研究）
"""
import os
import sys
import tempfile

import pythoncom
import win32com.client as win32

TMP = os.path.join(tempfile.gettempdir(), "bare_name_scope_probe.xlsm")

CODE = '''Sub Inner_Target()
    ThisWorkbook.Worksheets("S1").Range("C1").Value = "inner-ok"
End Sub

Sub Outer_RunBare()
    On Error GoTo EH
    Application.Run "Inner_Target"
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "bare:ok"
    Exit Sub
EH:
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "bare:FAIL " & Err.Number
End Sub

Sub Outer_RunQualified()
    On Error GoTo EH
    Application.Run ThisWorkbook.Name & "!Inner_Target"
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "qualified:ok"
    Exit Sub
EH:
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "qualified:FAIL " & Err.Number
End Sub

Sub Outer_DirectCall()
    On Error GoTo EH
    Call Inner_Target
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "call:ok"
    Exit Sub
EH:
    ThisWorkbook.Worksheets("S1").Range("C2").Value = "call:FAIL " & Err.Number
End Sub
'''


def main():
    if os.path.exists(TMP):
        os.remove(TMP)
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = True
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    try:
        wb = xl.Workbooks.Add()
        wb.Worksheets(1).Name = "S1"
        c = wb.VBProject.VBComponents.Add(1)
        c.Name = "ModScope"
        c.CodeModule.AddFromString(CODE)
        wb.SaveAs(TMP, FileFormat=52)
        wb.Windows(1).Visible = False        # 制造"装载器"状态：隐藏窗口
        wb.Save()
        wb.Close(SaveChanges=False)

        wb = xl.Workbooks.Open(TMP)
        aw = xl.ActiveWorkbook
        print("=" * 72)
        print("隐藏窗口的临时装载器:", TMP)
        print("ActiveWorkbook      :", "None（正是它导致裸名解析失败）" if aw is None else aw.Name)
        print("Windows(1).Visible  :", wb.Windows(1).Visible)
        print("=" * 72)
        name = os.path.basename(TMP)

        def clear():
            wb.Worksheets("S1").Range("C1:C2").ClearContents()

        def show(tag, macro):
            clear()
            note = ""
            try:
                xl.Run(macro)
            except Exception as e:  # noqa: BLE001
                note = "  ← Run 调用本身失败: %s" % str(e)[:60]
            got = (wb.Worksheets("S1").Range("C1").Value,
                   wb.Worksheets("S1").Range("C2").Value)
            print("  %-34s C1=%-12r C2=%r%s" % (tag, got[0], got[1], note))
            return got

        print("\n五种调用方式实测：")
        a = show("A) 外部调用 + 裸宏名", "Outer_RunBare")
        b = show("B) 外部调用 + 限定名", "%s!Outer_RunBare" % name)
        c = show("C) 内部 Run 裸名", "%s!Outer_RunBare" % name)   # 同上，看内层行为
        d = show("D) 内部 Run 限定名", "%s!Outer_RunQualified" % name)
        e = show("E) 内部直接 Call", "%s!Outer_DirectCall" % name)

        ok = (a[1] is None                     # A 应失败
              and b == ("inner-ok", "bare:ok")  # B 应成功
              and c == ("inner-ok", "bare:ok")  # C 内部裸名照样成功
              and d == ("inner-ok", "qualified:ok")
              and e == ("inner-ok", "call:ok"))
        print("\n结论：外部调用必须用限定名；VBA 内部互调不受影响。")
        print("预期一致：%s" % ("是 ✅" if ok else "否 ❌（Excel 行为可能变了，需重新研究）"))
        wb.Close(SaveChanges=False)
        return 0 if ok else 1
    finally:
        try:
            xl.Quit()
        except Exception:  # noqa: BLE001
            pass
        del xl
        if os.path.exists(TMP):
            try:
                os.remove(TMP)
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())
