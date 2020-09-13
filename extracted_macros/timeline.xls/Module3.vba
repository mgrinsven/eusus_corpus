Attribute VB_Name = "Module3"
Sub reverseprint()
Attribute reverseprint.VB_Description = "Macro recorded 6/21/99 by Deloitte & Touche"
Attribute reverseprint.VB_ProcData.VB_Invoke_Func = " \n14"
'
' reverseprint Macro
' Macro recorded 6/21/99 by Deloitte & Touche
'

'
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=2, Copies:=1, Collate _
        :=True
    Range("A1:G1").Select
    ActiveWorkbook.Save
End Sub
Sub chronoprint()
Attribute chronoprint.VB_Description = "Macro recorded 6/21/99 by Deloitte & Touche"
Attribute chronoprint.VB_ProcData.VB_Invoke_Func = " \n14"
'
' chronoprint Macro
' Macro recorded 6/21/99 by Deloitte & Touche
'

'
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=2, Copies:=1, Collate _
        :=True
    Range("A1:G1").Select
    ActiveWorkbook.Save
End Sub
