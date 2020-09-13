Attribute VB_Name = "Module1"
Sub sort()
Attribute sort.VB_Description = "Macro recorded 3/18/99 by Clifton A. Burris"
Attribute sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sort Macro
' Macro recorded 3/18/99 by Clifton A. Burris
'

'
    Range("A20:L56").Select
    Selection.sort Key1:=Range("A11"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    End Sub
Sub Subtotal()
Attribute Subtotal.VB_Description = "Macro recorded 3/18/99 by Clifton A. Burris"
Attribute Subtotal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Subtotal Macro
' Macro recorded 3/18/99 by Clifton A. Burris
'

'
    Range("A11").Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(13), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
End Sub
