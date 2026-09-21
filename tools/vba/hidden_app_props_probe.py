# -*- coding: utf-8 -*-
"""hidden_app_props_probe.py —— 钉死「窗口隐藏的装载器里，哪些 Application 设置会炸」

背景：工作簿窗口隐藏（装载器/宿主工作簿）时没有活动工作簿，除了「裸宏名解析失败」
和「宏内裸引用 1004」，还有**第三张面孔**：某些 Application 级设置也会 1004。
到底是哪些？凭记忆很容易张冠李戴，所以这里做同机 A/B 实测：

  同一份探针代码，先在**可见窗口**的正常工作簿上跑一遍（基线），
  再把同一个工作簿改成窗口隐藏，跑第二遍，逐条对照。

本探针自己造临时文件，不碰任何用户文件。

用法：python hidden_app_props_probe.py
退出码：0=结果与预期一致 ｜ 1=有偏差（Excel 行为可能变了，需重新研究）
"""
import os
import sys
import tempfile

import pythoncom
import win32com.client as win32

TMP = os.path.join(tempfile.gettempdir(), "hidden_app_props_probe.xlsm")

CODE = '''Sub Probe_AppProps()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    Dim r As Long
    r = 1
    On Error Resume Next

    Err.Clear
    Application.ScreenUpdating = False
    Record ws, r, "ScreenUpdating=False"
    r = r + 1
    Err.Clear
    Application.Calculation = xlCalculationManual
    Record ws, r, "Calculation=Manual"
    r = r + 1
    Err.Clear
    Application.DisplayAlerts = False
    Record ws, r, "DisplayAlerts=False"
    r = r + 1
    Err.Clear
    Application.StatusBar = "probe"
    Record ws, r, "StatusBar=text"
    r = r + 1
    Err.Clear
    Application.EnableEvents = False
    Record ws, r, "EnableEvents=False"
    r = r + 1
    Err.Clear
    Application.ScreenUpdating = True
    Record ws, r, "ScreenUpdating=True"
    r = r + 1
    Err.Clear
    Application.Calculation = xlCalculationAutomatic
    Record ws, r, "Calculation=Automatic"
    r = r + 1
    Err.Clear
    Application.StatusBar = False
    Record ws, r, "StatusBar=False"
    r = r + 1
    Err.Clear
    Application.DisplayAlerts = True
    Record ws, r, "DisplayAlerts=True"
    r = r + 1
    Err.Clear
    ThisWorkbook.Worksheets("S1").Range("Z1").Value = "qualified-write"
    Record ws, r, "qualified cell write"
    r = r + 1
    On Error GoTo 0
End Sub

Private Sub Record(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String)
    ws.Cells(r, 1).Value = label
    If Err.Number = 0 Then
        ws.Cells(r, 2).Value = "ok"
    Else
        ws.Cells(r, 2).Value = Err.Number & " | " & Left$(Err.Description, 60)
    End If
End Sub
'''

EXPECT_HIDDEN_DIFF = {"Calculation=Manual", "Calculation=Automatic"}


def read_table(wb):
    ws = wb.Worksheets("S1")
    out = {}
    for r in range(1, 12):
        k = ws.Cells(r, 1).Value
        v = ws.Cells(r, 2).Value
        if k:
            out[str(k)] = "" if v is None else str(v)
    return out


def run_probe(xl, wb, name):
    ws = wb.Worksheets("S1")
    ws.Range("A1:B12").ClearContents()
    try:
        xl.Run("%s!Probe_AppProps" % name)
    except Exception as e:  # noqa: BLE001
        print("  （Run 调用异常：%s）" % str(e)[:70])
    return read_table(wb)


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
        c.Name = "ModAppProps"
        c.CodeModule.AddFromString(CODE)
        wb.SaveAs(TMP, FileFormat=52)
        name = os.path.basename(TMP)

        print("=" * 78)
        print("第一遍：可见窗口（正常状态）—— 基线")
        print("=" * 78)
        a = run_probe(xl, wb, name)

        wb.Windows(1).Visible = False          # 制造「装载器」状态
        wb.Save()
        print("\n已把同一工作簿改成窗口隐藏（Windows(1).Visible=%s, ActiveWorkbook=%s）"
              % (wb.Windows(1).Visible, "None" if xl.ActiveWorkbook is None else xl.ActiveWorkbook.Name))

        print("\n" + "=" * 78)
        print("第二遍：窗口隐藏（装载器形态）—— 同一个工作簿、同一段代码")
        print("=" * 78)
        b = run_probe(xl, wb, name)

        print("\n" + "=" * 78)
        print("对照表")
        print("=" * 78)
        print("  %-24s %-34s %s" % ("设置项", "可见窗口", "窗口隐藏"))
        diffs = []
        for k in a:
            va, vb = a.get(k, "-"), b.get(k, "-")
            flag = ""
            if va != vb:
                diffs.append(k)
                flag = "   <== 差异"
            print("  %-24s %-34s %s%s" % (k, va[:34], vb[:40], flag))

        ok = set(diffs) == EXPECT_HIDDEN_DIFF
        print("\n结论：窗口隐藏时**只有 Application.Calculation 失败**（读或写都算），")
        print("      其余设置（ScreenUpdating / DisplayAlerts / StatusBar / EnableEvents /")
        print("      限定写法写单元格）全部照常工作。")
        print("      它只是提速开关，不影响结果 → 一律 On Error 兜住并按「自动」处理。")
        print("预期一致：%s" % ("是 ✅" if ok else "否 ❌（差异项=%s，Excel 行为可能变了，需重新研究）" % diffs))
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
