Attribute VB_Name = "Module2"
Sub SetSheetsUp()
Attribute SetSheetsUp.VB_Description = "Macro recorded 22/11/2001 by Michael LeRoy"
Attribute SetSheetsUp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' Macro recorded 22/11/2001 by Michael LeRoy
'

'
    Dim rang As String
    Dim Tot As Integer
    Tot = Range("TotalYears").Value
    
    Worksheets("Data Balance Sheet").Unprotect
    Worksheets("Data Income Statement").Unprotect
    Worksheets("Data Share Info").Unprotect
    Worksheets("Analysis Balance Sheet").Unprotect
    Worksheets("Analysis Income Statement").Unprotect
    Worksheets("Analysis Ratios").Unprotect
    
    Worksheets("Data Balance Sheet").Columns("C:H").Hidden = False
    Worksheets("Data Income Statement").Columns("C:H").Hidden = False
    Worksheets("Data Share Info").Columns("C:H").Hidden = False
    Worksheets("Analysis Balance Sheet").Columns("C:H").Hidden = False
    Worksheets("Analysis Income Statement").Columns("C:H").Hidden = False
    Worksheets("Analysis Ratios").Columns("C:H").Hidden = False
    Select Case Tot
        Case 1
            rang = "C:H"
        Case 2
            rang = "D:H"
        Case 3
            rang = "E:H"
        Case 4
            rang = "F:H"
        Case 5
            rang = "G:H"
        Case 6
            rang = "H:H"
    End Select
    If Tot <> 7 Then
        Worksheets("Data Balance Sheet").Columns(rang).Hidden = True
        Worksheets("Data Income Statement").Columns(rang).Hidden = True
        Worksheets("Data Share Info").Columns(rang).Hidden = True
        Worksheets("Analysis Balance Sheet").Columns(rang).Hidden = True
        Worksheets("Analysis Income Statement").Columns(rang).Hidden = True
        Worksheets("Analysis Ratios").Columns(rang).Hidden = True
    End If
    Worksheets("Data Balance Sheet").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("Data Income Statement").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("Data Share Info").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("Analysis Balance Sheet").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("Analysis Income Statement").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("Analysis Ratios").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
