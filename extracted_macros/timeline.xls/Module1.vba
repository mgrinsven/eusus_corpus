Attribute VB_Name = "Module1"
Sub printall()
Attribute printall.VB_Description = "Macro recorded 6/17/99 by Deloitte & Touche"
Attribute printall.VB_ProcData.VB_Invoke_Func = " \n14"
'
' printall Macro
' Macro recorded 6/17/99 by Deloitte & Touche
'

'
    Range("B2:I46").Select
    Selection.PrintOut Copies:=1, Collate:=True
End Sub
