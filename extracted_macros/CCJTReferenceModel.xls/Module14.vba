Attribute VB_Name = "Module14"
Sub BackToPerformance()
Attribute BackToPerformance.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute BackToPerformance.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BackToPerformance Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    Sheets("Performance Assumptions").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("B5:D5").Select
End Sub
Sub BackToTestFire()
Attribute BackToTestFire.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute BackToTestFire.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BackToTestFire Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    Sheets("Performance Assumptions").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B28:D28").Select
End Sub
Sub BackToLocations()
Attribute BackToLocations.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute BackToLocations.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BackToLocations Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    Sheets("Performance Assumptions").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B49:D49").Select
End Sub
Sub PerformanceToBackground()
Attribute PerformanceToBackground.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute PerformanceToBackground.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PerformanceToBackground Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    Sheets("Background State Information").Select
    Range("B26:F26").Select
End Sub
Sub PrintParams()
Attribute PrintParams.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute PrintParams.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintParams Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
Sub PrintResultsSumm()
Attribute PrintResultsSumm.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute PrintResultsSumm.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintResultsSumm Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    ActiveSheet.Shapes("Button 18").Select
    Selection.Characters.Text = "PRINT"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("C4").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PrintResultsSumm"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ReviewAssumptions"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ReturnTop"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    Sheets("Results Detail").Select
    ActiveWorkbook.Save
    ActiveWindow.LargeScroll Down:=1
    Range("D25").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    ActiveWindow.LargeScroll Down:=1
    Range("D25").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    ActiveWindow.LargeScroll Down:=1
    Range("A24").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A1").Select
    Application.Goto Reference:="EquipmentDetails"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!FloorspaceDetails"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!EquipmentDetails"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    ActiveSheet.Shapes("Button 1").Select
    Selection.Copy
    Sheets("Results Detail").Select
    Range("D2").Select
    ActiveSheet.Buttons.Add(190.5, 8.25, 62.25, 39).Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft 10.5
    Selection.ShapeRange.IncrementTop -13.5
    Range("D3").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!SummaryToMap"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!MapToResultsSummary"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    ActiveSheet.Shapes("Button 18").Select
    Selection.Copy
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Range("E2").Select
    ActiveSheet.Buttons.Add(303, 10.5, 93.75, 40.5).Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft -213#
    Selection.ShapeRange.IncrementTop -13.5
    Range("F3").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!ResultsSummary"
    ActiveSheet.Shapes("Button 3").Select
    Selection.Copy
    Range("G4").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V30 with buttons.xls'!PersonnelDetail"
    Range("F3").Select
    ActiveSheet.Buttons.Add(474, 6.75, 63.75, 39.75).Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft -1.5
    Selection.ShapeRange.IncrementTop -24.75
    Range("F2").Select
    ActiveSheet.Shapes("Button 15").Select
    Selection.OnAction = "PrintResultsDetails"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
Sub PrintResultsDetails()
Attribute PrintResultsDetails.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute PrintResultsDetails.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintResultsDetails Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
Sub PrintComments()
Attribute PrintComments.VB_Description = "Macro recorded 12/6/2001 by Linda Nichols"
Attribute PrintComments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintComments Macro
' Macro recorded 12/6/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
