Attribute VB_Name = "QuickCommands"
' ============================================================
' QuickCommands for Word - 一键快捷指令
' ============================================================
' 安装方法（3步，1分钟搞定）：
'   1. 打开 Word，按 Alt+F11 打开 VBA 编辑器
'   2. 左侧"Normal"项目右键 → 导入文件 → 选 word.bas
'   3. 按 Ctrl+S 保存，关闭 Word 重新打开
' 工具栏自动出现！
' ============================================================

' ---- 启动时自动创建工具栏 ----
Sub AutoExec()
    AddToolbarButtons
End Sub

' ---- 关闭时清理工具栏 ----
Sub AutoExit()
    On Error Resume Next
    Application.CommandBars("QuickCommands").Delete
    On Error GoTo 0
End Sub

' ---- 工具栏创建 ----
Private Sub AddToolbarButtons()
    Dim cmdBar As Object, btn As Object
    On Error Resume Next
    
    Application.CommandBars("QuickCommands").Delete
    Set cmdBar = Application.CommandBars.Add(Name:="QuickCommands", Position:=msoBarTop, Temporary:=False)
    
    AddButton cmdBar, "日期时间", "InsertDate"
    AddButton cmdBar, "加粗", "ToggleBold"
    AddButton cmdBar, "斜体", "ToggleItalic"
    AddButton cmdBar, "下划线", "ToggleUnderline"
    AddButton cmdBar, "居中", "AlignCenter"
    AddButton cmdBar, "左对齐", "AlignLeft"
    AddButton cmdBar, "右对齐", "AlignRight"
    AddButton cmdBar, "两端对齐", "AlignJustify"
    AddButton cmdBar, "高亮", "ToggleHighlight"
    AddButton cmdBar, "清格式", "ClearFormat"
    AddButton cmdBar, "删空段", "RemoveEmpty"
    AddButton cmdBar, "合并空格", "NormalizeSpaces"
    AddButton cmdBar, "页码", "InsertPageNumbers"
    AddButton cmdBar, "导出PDF", "ExportPDF"
    AddButton cmdBar, "标题1", "ApplyHeading1"
    AddButton cmdBar, "标题2", "ApplyHeading2"
    AddButton cmdBar, "正文", "ApplyNormal"
    AddButton cmdBar, "项目符号", "BulletList"
    AddButton cmdBar, "编号", "NumberList"
    
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
    Selection.TypeText Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Sub ToggleBold()
    Selection.Font.Bold = Not Selection.Font.Bold
End Sub

Sub ToggleItalic()
    Selection.Font.Italic = Not Selection.Font.Italic
End Sub

Sub ToggleUnderline()
    If Selection.Font.Underline = 0 Then
        Selection.Font.Underline = 1
    Else
        Selection.Font.Underline = 0
    End If
End Sub

Sub AlignCenter()
    Selection.ParagraphFormat.Alignment = 1
End Sub

Sub AlignLeft()
    Selection.ParagraphFormat.Alignment = 0
End Sub

Sub AlignRight()
    Selection.ParagraphFormat.Alignment = 2
End Sub

Sub AlignJustify()
    Selection.ParagraphFormat.Alignment = 3
End Sub

Sub ToggleHighlight()
    With Selection.Range
        If .HighlightColorIndex = 0 Then
            .HighlightColorIndex = 7
        Else
            .HighlightColorIndex = 0
        End If
    End With
End Sub

Sub ClearFormat()
    Selection.ClearFormatting
End Sub

Sub RemoveEmpty()
    Dim p As Object
    For Each p In ActiveDocument.Paragraphs
        If Len(Trim(p.Range.Text)) = 1 Then p.Range.Delete
    Next
End Sub

Sub NormalizeSpaces()
    Dim r As Object
    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " {2,}"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=2
    End With
End Sub

Sub InsertPageNumbers()
    ActiveDocument.Sections(1).Footers(1).PageNumbers.Add 1
End Sub

Sub DeleteBlankPages()
    Dim s As Object
    For Each s In ActiveDocument.Sections
        If s.Range.ComputeStatistics(0) = 0 Then s.Range.Delete
    Next
End Sub

Sub ExportPDF()
    Dim outPath As String
    outPath = Replace(ActiveDocument.FullName, ".docx", ".pdf")
    If outPath = ActiveDocument.FullName Then
        outPath = Replace(ActiveDocument.FullName, ".doc", ".pdf")
    End If
    ActiveDocument.ExportAsFixedFormat outPath, 17
End Sub

Sub ApplyHeading1()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
End Sub

Sub ApplyHeading2()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading2)
End Sub

Sub ApplyNormal()
    Selection.Style = ActiveDocument.Styles(wdStyleNormal)
End Sub

Sub BulletList()
    Selection.Range.ListFormat.ApplyBulletDefault
End Sub

Sub NumberList()
    Selection.Range.ListFormat.ApplyNumberDefault
End Sub

Sub RemoveList()
    Selection.Range.ListFormat.RemoveNumbers
End Sub

Sub MultiLevelList()
    Selection.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(1).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=0
End Sub
