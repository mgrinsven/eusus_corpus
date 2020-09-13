Attribute VB_Name = "Module2"
Sub GrabElec()
Attribute GrabElec.VB_Description = "Macro recorded 10/16/2002 by Tim Keely"
Attribute GrabElec.VB_ProcData.VB_Invoke_Func = "e\r14"
'
' GrabElec Macro
' Macro recorded 10/16/2002 by Tim Keely
'
' Keyboard Shortcut: Option+Cmd+e
'
    Range("E1:L155").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "N1:U2"), CopyToRange:=Range("N4:U4"), Unique:=False
End Sub
Sub GrabConc()
Attribute GrabConc.VB_Description = "Macro recorded 10/16/2002 by Tim Keely"
Attribute GrabConc.VB_ProcData.VB_Invoke_Func = "t\r14"
'
' GrabConc Macro
' Macro recorded 10/16/2002 by Tim Keely
'
' Keyboard Shortcut: Option+Cmd+t
'
    Sheets("course list").Select
   
    Range("E1:L155").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "X1:AE2"), CopyToRange:=Range("X4:AE4"), Unique:=False
End Sub
Sub GrabCore()
Attribute GrabCore.VB_Description = "Macro recorded 10/16/2002 by Tim Keely"
Attribute GrabCore.VB_ProcData.VB_Invoke_Func = "r\r14"
'
' GrabCore Macro
' Macro recorded 10/16/2002 by Tim Keely
'
' Keyboard Shortcut: Option+Cmd+r

Sheets("course list").Select
    Range("E1:L155").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "AH1:AO2"), CopyToRange:=Range("AH4:AO4"), Unique:=False
    ActiveWindow.ScrollColumn = 11
End Sub
Sub GrabProj()
Attribute GrabProj.VB_Description = "Macro recorded 10/16/2002 by Tim Keely"
Attribute GrabProj.VB_ProcData.VB_Invoke_Func = "p\r14"
'
' GrabProj Macro
' Macro recorded 10/16/2002 by Tim Keely
'
' Keyboard Shortcut: Option+Cmd+p
'
    
    Sheets("course list").Select
    Range("E1:L155").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "AR1:AY7"), CopyToRange:=Range("AR9:AY9"), Unique:=False
    ActiveWindow.ScrollColumn = 33
End Sub
