Attribute VB_Name = "Module5"
Sub GrabUn1()
Attribute GrabUn1.VB_Description = "Macro recorded 10/17/2002 by Tim Keely"
Attribute GrabUn1.VB_ProcData.VB_Invoke_Func = " \r14"
'
' GrabUn1 Macro
' Macro recorded 10/17/2002 by Tim Keely
'

'
    Range("E1:L142").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "BB1:BI4"), CopyToRange:=Range("BB6:BI6"), Unique:=False
    ActiveWindow.ScrollColumn = 42
End Sub
