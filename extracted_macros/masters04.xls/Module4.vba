Attribute VB_Name = "Module4"
Sub update()
Attribute update.VB_Description = "Macro recorded 10/17/2002 by Tim Keely"
Attribute update.VB_ProcData.VB_Invoke_Func = " \r14"
'
' update Macro
' Macro recorded 10/17/2002 by Tim Keely
'

'
    Sheets("course list").Select
    Application.Run "masters04.xls!GrabProj"
    Application.Run "masters04.xls!GrabElec"
    Application.Run "masters04.xls!GrabCore"
    Application.Run "masters04.xls!GrabConc"
    Sheets("Print").Select
    Range("A1").Select
End Sub
Sub GrabMasUnits()
Attribute GrabMasUnits.VB_Description = "Macro recorded 10/17/2002 by Tim Keely"
Attribute GrabMasUnits.VB_ProcData.VB_Invoke_Func = " \r14"
'
' GrabMasUnits Macro
' Macro recorded 10/17/2002 by Tim Keely
'

'
    Range("E1:L142").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "BB1:BI4"), CopyToRange:=Range("BB9:BI9"), Unique:=False
End Sub
