Attribute VB_Name = "Module2"
Sub ProtectSheet()
Attribute ProtectSheet.VB_Description = "Macro recorded 2000-08-03 by George Smith"
Attribute ProtectSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ProtectSheet Macro
' Macro recorded 2000-08-03 by George Smith
'

'
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Sub UnprotectSheet()
Attribute UnprotectSheet.VB_Description = "Macro recorded 2000-08-03 by George Smith"
Attribute UnprotectSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UnprotectSheet Macro
' Macro recorded 2000-08-03 by George Smith
'

'
    ActiveSheet.Unprotect
End Sub
Sub DeleteRow()
Attribute DeleteRow.VB_Description = "Macro recorded 2000-08-03 by George Smith"
Attribute DeleteRow.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' DeleteRow Macro
' Macro recorded 2000-08-03 by George Smith
'
' Keyboard Shortcut: Ctrl+d
'
    Application.Run "TCCashFlow2.xls!UnprotectSheet"
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    Application.Run "TCCashFlow2.xls!ProtectSheet"
End Sub
