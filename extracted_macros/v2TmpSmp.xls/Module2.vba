Attribute VB_Name = "Module2"


Sub Surface()
Attribute Surface.VB_Description = "Macro recorded 7/6/99 by Audrey F. Borchardt"
Attribute Surface.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' Surface Macro
' Macro recorded 7/6/99 by Audrey F. Borchardt
'
' Keyboard Shortcut: Ctrl+a
'
    Range("B28").Select
    Selection.Copy
    Range("B28:AZ78").Select
    ActiveSheet.Paste
End Sub
Sub Testlocal()
Attribute Testlocal.VB_Description = "Macro recorded 7/6/99 by Audrey F. Borchardt"
Attribute Testlocal.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' Testlocal Macro
' Macro recorded 7/6/99 by Audrey F. Borchardt
'
' Keyboard Shortcut: Ctrl+b
'
    Range("B20").Select
    Selection.Copy
    Range("A19:C21").Select
    ActiveSheet.Paste
End Sub
