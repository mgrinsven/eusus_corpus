Attribute VB_Name = "Module3"
Sub InsertCalculatedRow()
Attribute InsertCalculatedRow.VB_Description = "Macro recorded 2000-08-04 by George Smith"
Attribute InsertCalculatedRow.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' InsertCalculatedRow Macro
' Macro recorded 2000-08-04 by George Smith
'
' Keyboard Shortcut: Ctrl+n
'
    Application.Run "TCCashFlow2.xls!UnprotectSheet"
    Application.Run "TCCashFlow2.xls!InsertNewDetailLine"
    Application.Run "TCCashFlow2.xls!ProtectSheet"
End Sub
