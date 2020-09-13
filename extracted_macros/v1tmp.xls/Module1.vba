Attribute VB_Name = "Module1"





Sub SlopeGraph()
Attribute SlopeGraph.VB_Description = "Macro recorded 7/2/99 by Audrey F. Borchardt"
Attribute SlopeGraph.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' SlopeGraph Macro
' Macro recorded 7/2/99 by Audrey F. Borchardt
'
' Keyboard Shortcut: Ctrl+d
'
    Range("B13").Select
    Selection.Copy
    Range("B14:B63").Select
    ActiveSheet.Paste
    Range("B13:B63").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B67:B117").Select
    ActiveSheet.Paste
    Range("E67:E117").Select
    ActiveSheet.Paste
End Sub
