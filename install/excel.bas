Attribute VB_Name = "QuickCommands"
' ============================================================
' QuickCommands for Excel - 一键快捷指令
' ============================================================
' 安装方法（3步，1分钟搞定）：
'   1. 打开 Excel，按 Alt+F11 打开 VBA 编辑器
'   2. 左侧"VBAProject(PERSONAL.XLSB)"项目右键 → 导入文件 → 选 excel.bas
'      如果没有 PERSONAL.XLSB：先录一个宏（视图→宏→录制宏→存在位置选"个人宏工作簿"→停止录制），它会自动创建
'   3. 按 Ctrl+S 保存，关闭 Excel 重新打开
' 工具栏自动出现！
' ============================================================

' ---- 启动时自动创建工具栏 ----
Sub Auto_Open()
    AddToolbarButtons
End Sub

' ---- 工具栏创建 ----
Private Sub AddToolbarButtons()
    Dim cmdBar As Object, btn As Object
    On Error Resume Next
    
    Application.CommandBars("QuickCommands").Delete
    Set cmdBar = Application.CommandBars.Add(Name:="QuickCommands", Position:=msoBarTop, Temporary:=False)
    
    AddButton cmdBar, "日期", "InsertDate"
    AddButton cmdBar, "加粗", "ToggleBold"
    AddButton cmdBar, "斜体", "ToggleItalic"
    AddButton cmdBar, "下划线", "ToggleUnderline"
    AddButton cmdBar, "居中", "AlignCenter"
    AddButton cmdBar, "左对齐", "AlignLeft"
    AddButton cmdBar, "右对齐", "AlignRight"
    AddButton cmdBar, "高亮", "ToggleHighlight"
    AddButton cmdBar, "清格式", "ClearFormat"
    AddButton cmdBar, "清内容", "ClearContents"
    AddButton cmdBar, "插行", "InsertRow"
    AddButton cmdBar, "插列", "InsertColumn"
    AddButton cmdBar, "升序", "SortAsc"
    AddButton cmdBar, "降序", "SortDesc"
    AddButton cmdBar, "求和", "SumBelow"
    AddButton cmdBar, "求均值", "AverageBelow"
    AddButton cmdBar, "转值", "FormulaToValue"
    AddButton cmdBar, "去重", "RemoveDuplicates"
    AddButton cmdBar, "筛选", "FilterData"
    AddButton cmdBar, "冻结首行", "FreezeTopRow"
    AddButton cmdBar, "导出CSV", "ExportCSV"
    AddButton cmdBar, "导出PDF", "ExportPDF"
    
    cmdBar.Visible = True
    On Error GoTo 0
End Sub

Private Sub AddButton(ByVal cmdBar As Object, ByVal caption As String, ByVal action As String)
    Dim btn As Object
    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    btn.Caption = caption
    btn.OnAction = action
    btn.Style = msoButtonCaption
End Sub

' ============================================================
' 功能宏
' ============================================================

Sub InsertDate()
    Dim rng As Object
    Set rng = Selection
    If rng.Cells.Count = 1 Then
        rng.Value = Format(Date, "yyyy-mm-dd")
    Else
        Dim cell As Object
        For Each cell In rng.Cells
            cell.Value = Format(Date, "yyyy-mm-dd")
        Next
    End If
End Sub

Sub ToggleBold()
    Selection.Font.Bold = Not Selection.Font.Bold
End Sub

Sub ToggleItalic()
    Selection.Font.Italic = Not Selection.Font.Italic
End Sub

Sub ToggleUnderline()
    If Selection.Font.Underline = -4142 Then
        Selection.Font.Underline = 2
    Else
        Selection.Font.Underline = -4142
    End If
End Sub

Sub AlignCenter()
    Selection.HorizontalAlignment = -4108
End Sub

Sub AlignLeft()
    Selection.HorizontalAlignment = -4131
End Sub

Sub AlignRight()
    Selection.HorizontalAlignment = -4152
End Sub

Sub ToggleHighlight()
    Dim ci As Long
    ci = Selection.Interior.ColorIndex
    If ci = 6 Then
        Selection.Interior.ColorIndex = -4142
    Else
        Selection.Interior.ColorIndex = 6
    End If
End Sub

Sub ClearFormat()
    Selection.ClearFormats
End Sub

Sub ClearContents()
    Selection.ClearContents
End Sub

Sub InsertRow()
    Selection.EntireRow.Insert
End Sub

Sub InsertColumn()
    Selection.EntireColumn.Insert
End Sub

Sub SortAsc()
    Selection.Sort Key1:=Selection, Order1:=1, Header:=0
End Sub

Sub SortDesc()
    Selection.Sort Key1:=Selection, Order1:=2, Header:=0
End Sub

Sub SumBelow()
    Selection.Formula = "=SUM(" & Selection.Offset(-1, 0).Resize(1, Selection.Columns.Count).Address(False, False) & ")"
End Sub

Sub AverageBelow()
    Selection.Formula = "=AVERAGE(" & Selection.Offset(-1, 0).Resize(1, Selection.Columns.Count).Address(False, False) & ")"
End Sub

Sub FormulaToValue()
    Selection.Value = Selection.Value
End Sub

Sub RemoveDuplicates()
    Selection.RemoveDuplicates Columns:=1, Header:=0
End Sub

Sub FilterData()
    Selection.AutoFilter
End Sub

Sub FreezeTopRow()
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Sub UnprotectSheet()
    ActiveSheet.Unprotect
End Sub

Sub ProtectSheet()
    ActiveSheet.Protect
End Sub

Sub ExportCSV()
    Dim outPath As String
    outPath = Replace(ActiveWorkbook.FullName, ".xlsx", ".csv")
    If outPath = ActiveWorkbook.FullName Then
        outPath = Replace(ActiveWorkbook.FullName, ".xlsm", ".csv")
    End If
    ' 保存一份 CSV 副本，不改变当前工作簿
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs Filename:=outPath, FileFormat:=6
    ActiveWorkbook.Close SaveChanges:=False
End Sub

Sub ExportPDF()
    Dim outPath As String
    outPath = Replace(ActiveWorkbook.FullName, ".xlsx", ".pdf")
    If outPath = ActiveWorkbook.FullName Then
        outPath = Replace(ActiveWorkbook.FullName, ".xlsm", ".pdf")
    End If
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=outPath
End Sub
