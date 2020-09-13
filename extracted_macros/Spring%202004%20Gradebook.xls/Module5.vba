Attribute VB_Name = "Module5"
Sub FormatAndProtect()
Attribute FormatAndProtect.VB_Description = "Macro recorded 2/11/2004 by Zahra and Paul"
Attribute FormatAndProtect.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatAndProtect Macro
' Macro recorded 2/11/2004 by Zahra and Paul
'

'
For i = 1 To Sheets.Count Step 2
    Sheets(i).Select
    Columns("A:E").Select
    Selection.EntireColumn.Hidden = True
    Cells.Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets(i + 1).Select
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Cells.Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
Next i
End Sub
