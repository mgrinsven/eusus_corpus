Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_Description = "Macro recorded 4/22/03 by Maria Mays"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' Macro1 Macro
' Macro recorded 4/22/03 by Maria Mays
'
' Keyboard Shortcut: Ctrl+g
'
    Sheets("Sheet1").Select
    ActiveCell.Offset(-79, 0).Range("A1:H1,B:B").Select
    Sheets("Sheet2").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
End Sub
