Attribute VB_Name = "Module3"
Sub UnprotectSheets()
Attribute UnprotectSheets.VB_Description = "Macro recorded 12/11/2003 by Instructional Computing"
Attribute UnprotectSheets.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' UnprotectSheets Macro
' Macro recorded 12/11/2003 by Instructional Computing
'
' Keyboard Shortcut: Ctrl+Shift+U
'
For i = 1 To Sheets.Count
    a$ = Sheets(i).Name
    Sheets(i).Select
    Cells.Select
    ActiveSheet.Unprotect
Next i
End Sub
Sub TransferGrades()
InSIDColumn = 3
InGradeColumn = 66
InstartRow = 2
InEndRow = 21
OutSIDColumn = 3
OutGradeColumn = 5
OutStartRow = 2
OutEndRow = 25
InWindow = "s04m04.xls"
InSheet = "Sheet1"
OutSheet = "CIS 105 1704 Spring 2004"
OutWindow = "Spring 2004 Gradebook.xls"
For i = InstartRow To InEndRow
    Windows(InWindow).Activate
    thesid = Cells(i, InSIDColumn)
    theGrade = Cells(i, InGradeColumn) * 100
    Windows(OutWindow).Activate
    For j = OutStartRow To OutEndRow
        If Cells(j, OutSIDColumn) = thesid Then
            Cells(j, OutGradeColumn) = theGrade
            Exit For
        End If
    Next j
Next i
End Sub
