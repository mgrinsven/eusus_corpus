Attribute VB_Name = "Module1"
Sub RemoveHrsText()
Attribute RemoveHrsText.VB_Description = "Macro recorded 4/18/2001 by John Hail"
Attribute RemoveHrsText.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RemoveHrsText Macro
' Macro recorded 4/18/2001 by John Hail
'

'
    Columns("F:F").Select
    Selection.Replace What:=" hrs", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
End Sub
