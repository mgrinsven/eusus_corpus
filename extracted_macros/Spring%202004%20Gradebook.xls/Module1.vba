Attribute VB_Name = "Module1"
Dim SheetToSearch As Integer
Sub CompareWorkbooks()
For Count = 1 To ActiveWorkbook.Sheets.Count
    SheetToSearch = Count
    CompareTwoSheets
Next Count
End Sub
Sub CompareTwoSheets()
State = "are the same."
Workbooks(1).Activate
Sheets(SheetToSearch).Activate
NumberOfRows = ActiveSheet.UsedRange.Rows.Count
NumberOfColumns = ActiveSheet.UsedRange.Columns.Count
For i = 1 To NumberOfRows
    For j = 1 To NumberOfColumns
        If CellsDiffer(Int(i), Int(j)) Then
            Cells(i, j).Select
            With Selection.Interior
                .ColorIndex = 38
                .Pattern = xlSolid
            End With
            State = "differed as shown."
        End If
    Next j
Next i
MsgBox ("The sheets " & State)
End Sub
Function CellsDiffer(row As Integer, column As Integer) As Boolean
'
' CellsDiffer Macro
' Macro recorded 12/11/2003 by Zahra Ebrahimian
'
    CellsDiffer = False
    Sheets(SheetToSearch).Select
    string1$ = Cells(row, column).Value
    Workbooks(2).Activate
    Sheets(SheetToSearch).Select
    If (Cells(row, column).Value <> string1$) Then
        CellsDiffer = True
    End If
    Workbooks(1).Activate
End Function

