Attribute VB_Name = "Module1"


Sub HideColumnsRows()
'
' HideColumnsRows Macro
' Macro recorded 05/02/2004 by TIMBER_Consultant
'
'   This macro shows only the selected data (years & series/rows & columns)on worksheet -Data-
'   Attention: range for columns starts on cell F5 and range for rows starts on cell C18
    Worksheets("Data").Activate
    Range("F7").Select
    Unprotectsheet
'   Selecting series/columns
    Range(Selection, Selection.End(xlToRight)).Select
'   Showing only the selected series
    If TypeName(Selection) <> "Range" Then Exit Sub
    For Each Cell In Selection
        If Cell.Value = "N" Then
            Cell.EntireColumn.Hidden = True
        End If
    Next Cell
    Range("C18").Select
    Unprotectsheet
'   Selecting years/rows
    Range(Selection, Selection.End(xlDown)).Select
'   Showing only the selected years
    Dim StartYear As Integer
    Dim EndYear As Integer
'
    Range("A1").Select
    StartYear = Worksheets("Data").Range("D5")
    EndYear = Worksheets("Data").Range("D6")
    
    Range("D17").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter Field:=1, Criteria1:=">=" & StartYear, Operator:=xlAnd, _
        Criteria2:="<=" & EndYear
    
    
'    If TypeName(Selection) <> "Range" Then Exit Sub
'    For Each Cell In Selection
'        If Cell.Value = "N" Then
'        Cell.EntireRow.Hidden = True
'        End If
'     Next Cell
     Range("A3").Select
    ProtectSheet
'
End Sub


Sub UnhideColumnsRows()
'
' UnhideColumnsRows Macro
' Macro recorded 05/02/2004 by TIMBER_Consultant
'
'   This macro show all columns and rowns of worksheet -Data-
'   Please note that the data range is from column F to column IV
'   and the program expects that the first year is in cell C18
    Worksheets("Data").Activate
    Unprotectsheet
'   Unhiding all the columns
    Worksheets("Data").AutoFilterMode = False
    Columns("F:IV").Select
    Selection.EntireColumn.Hidden = False
    Range("A2").Select
    Worksheets("Data").Activate
'   Unhiding all rows - suppressed after adding AutoFilterMode line
'       Range("D17").Select
'       Selection.AutoFilter
'       Range("F18").Select
'       Range("C17").Select
'       Rows("18:65500").Select
'       Selection.EntireRow.Hidden = False
    Range("A3").Select
    ProtectSheet
End Sub


Sub Unprotectsheet()
'
' Unprotectsheet Macro
' Macro recorded 20/01/2004 by TIMBER_Consultant
'

'
    ActiveSheet.Unprotect
End Sub

Sub ProtectSheet()
'
' ProtectSheet Macro
' Macro recorded 12/02/2004 by TIMBER_Consultant
'

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
'
End Sub

