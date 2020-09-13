Attribute VB_Name = "Module2"
Sub Refresh()
Attribute Refresh.VB_Description = "Macro recorded 3/19/99 by Clifton A. Burris"
Attribute Refresh.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Refresh Macro
' Macro recorded 3/19/99 by Clifton A. Burris
'

'
    Sheets("Form 219").Select
    ActiveSheet.PivotTables("PivotTable3").PivotSelect "", xlDataAndLabel
    ActiveSheet.PivotTables("PivotTable3").RefreshTable
    ActiveSheet.PivotTables("PivotTable3").PivotSelect "", xlLabelOnly
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Font.Bold = True
End Sub
Sub Clear()
Attribute Clear.VB_Description = "Macro recorded 3/19/99 by Clifton A. Burris"
Attribute Clear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear Macro
' Macro recorded 3/19/99 by Clifton A. Burris
'

'
    Sheets("Data Input").Select
    Range("A20").Select
    Selection.RemoveSubtotal
End Sub
Sub Transfer()
Attribute Transfer.VB_Description = "Macro recorded 3/19/99 by Clifton A. Burris"
Attribute Transfer.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Transfer Macro
' Macro recorded 3/19/99 by Clifton A. Burris
'

'
    Sheets("Form 219").Select
    ActiveWindow.SmallScroll ToRight:=-6
    ActiveSheet.PivotTables("PivotTable4").PivotSelect "", xlDataAndLabel
    ActiveSheet.PivotTables("PivotTable4").RefreshTable
    ActiveSheet.PivotTables("PivotTable4").PivotSelect "", xlLabelOnly
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Font.Bold = True
    ActiveSheet.PivotTables("PivotTable4").PivotSelect "", xlDataOnly
    Selection.NumberFormat = "$#,##0.00_);($#,##0.00)"
    Sheets("Data Input").Select
    Range("A20").Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(12), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    Range("A20:L71").Select
    Selection.Copy
    Sheets("Form 219").Select
    ActiveWindow.SmallScroll Down:=-53
    Range("A17").Select
    ActiveSheet.Paste
    Range("B17").Select
    Sheets("Data Input").Select
    ActiveWindow.SmallScroll Down:=-42
    ActiveWindow.SmallScroll ToRight:=-8
    Range("A11").Select
    Application.CutCopyMode = False
End Sub
Sub Clear219()
Attribute Clear219.VB_Description = "Macro recorded 3/19/99 by Clifton A. Burris"
Attribute Clear219.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear219 Macro
' Macro recorded 3/19/99 by Clifton A. Burris
'

'
    Sheets("Form 219").Select
    ActiveWindow.SmallScroll Down:=-37
    Range("A17:L68").Select
    Selection.ClearContents
End Sub
