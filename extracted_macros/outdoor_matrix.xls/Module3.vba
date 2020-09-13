Attribute VB_Name = "Module3"
Sub Rank_Criteria()
Attribute Rank_Criteria.VB_Description = "Macro recorded 4/27/99 by Kamel S. Saidi"
Attribute Rank_Criteria.VB_ProcData.VB_Invoke_Func = " \r14"
'
' Rank_Criteria Macro
' Macro recorded 4/27/99 by Kamel S. Saidi
'

'
    Range("A2:B12").Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End Sub
