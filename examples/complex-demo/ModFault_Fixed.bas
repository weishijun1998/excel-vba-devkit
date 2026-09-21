Option Explicit

' ============================================================
' ModFault_Fixed —— G3 那个地雷的修好版：把裸引用限定到具体对象
'   同一位置、同一功能，只是把 Sheets("明细数据") 换成 ThisWorkbook.Worksheets(...)
' ============================================================

Public Sub Fault_GoodRef()
    ThisWorkbook.Worksheets("明细数据").Range("Z10").Value = "qualified-ok"
End Sub
