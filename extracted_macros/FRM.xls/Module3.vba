Attribute VB_Name = "Module3"
Sub REINITIALIZE()
Attribute REINITIALIZE.VB_Description = "Macro recorded 09/20/2000 by greulich"
Attribute REINITIALIZE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' REINITIALIZE Macro
' Macro recorded 09/20/2000 by greulich
'

'
    Range("K4:K14").Select
    Selection.Copy
    Range("C4").Select
    ActiveSheet.Paste
    Range("K19:K57").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E19").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=117
    ActiveWindow.ScrollColumn = 1
    Range("D164").Select
    Application.CutCopyMode = False
End Sub
