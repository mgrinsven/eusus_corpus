Attribute VB_Name = "Module13"
Sub PersonnelDetail()
Attribute PersonnelDetail.VB_Description = "Macro recorded 12/5/2001 by Linda Nichols"
Attribute PersonnelDetail.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PersonnelDetail Macro
' Macro recorded 12/5/2001 by Linda Nichols
'

'
    Sheets("Results Detail").Select
    Range("D7:G7").Select
End Sub
Sub FloorspaceDetails()
Attribute FloorspaceDetails.VB_Description = "Macro recorded 12/5/2001 by Linda Nichols"
Attribute FloorspaceDetails.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FloorspaceDetails Macro
' Macro recorded 12/5/2001 by Linda Nichols
'

'
    Sheets("Results Detail").Select
    ActiveWindow.LargeScroll Down:=1
    Range("D45:G45").Select
End Sub
Sub EquipmentDetails()
Attribute EquipmentDetails.VB_Description = "Macro recorded 12/5/2001 by Linda Nichols"
Attribute EquipmentDetails.VB_ProcData.VB_Invoke_Func = " \n14"
'
' EquipmentDetails Macro
' Macro recorded 12/5/2001 by Linda Nichols
'

'
    Sheets("Results Detail").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("D20:G20").Select
End Sub
Sub ResultsSummary()
Attribute ResultsSummary.VB_Description = "Macro recorded 12/5/2001 by Linda Nichols"
Attribute ResultsSummary.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ResultsSummary Macro
' Macro recorded 12/5/2001 by Linda Nichols
'

'
    Sheets("Results Summary").Select
    Range("A1").Select
End Sub
