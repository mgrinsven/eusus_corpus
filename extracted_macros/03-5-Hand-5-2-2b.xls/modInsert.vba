Attribute VB_Name = "modInsert"

' InsertRows, SetFormulas Macros
' Written by Alice M. Wenk 4/19/97

Sub InsertRows()
Attribute InsertRows.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim lastRow As Integer
    Dim numRows As Integer
    Dim InsertRows As Integer
    Dim diffRow As Integer
    Dim firstRow As Integer
    Dim endRow As Integer
    Dim iCount As Integer
    Dim dlgWhere As DialogSheet
    Dim section As String
    Dim optn As OptionButton
    Dim xFootFormula As String
    Dim sumFormula As String
    
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("Evaluation_of_Misstatements").Activate
    
' display dialog window, obtain section (place on worksheet) and number of rows to add
    Set dlgWhere = ThisWorkbook.DialogSheets("DialogInsert")

TryAgain:
    If dlgWhere.Show Then
        If dlgWhere.OptionButtons("optKnown").Value = xlOn Then
                section = "Known"
            ElseIf dlgWhere.OptionButtons("optLikely").Value = xlOn Then
                section = "Likely"
            ElseIf dlgWhere.OptionButtons("optCOKnown").Value = xlOn Then
                section = "CarryoverKnown"
            ElseIf dlgWhere.OptionButtons("optCOLikely").Value = xlOn Then
                section = "CarryoverLikely"
        End If
        If dlgWhere.EditBoxes("txtNumRows").Text = "" Or section = "" Then
            MsgBox "All required information has not been entered"
            GoTo TryAgain
        Else
             numRows = CInt(dlgWhere.EditBoxes("txtNumRows").Text)
        End If
    Else
        Exit Sub
    End If

    ' insert rows into appropriate section
    Select Case section
    Case "Known"
        lastRow = 1
        Do Until Cells(lastRow, 1).Text = "Total known misstatements"
            lastRow = lastRow + 1
        Loop
        Cells(lastRow, 1).Activate
        For InsertRows = 1 To numRows
            ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Copy
            Selection.Insert Shift:=xlDown
            ActiveCell.Range("A1:G1").ClearContents
        Next InsertRows
        Application.CutCopyMode = False

    Case "Likely"
        lastRow = 1
        Do Until Cells(lastRow, 1).Text = "Total likely misstatements"
            lastRow = lastRow + 1
        Loop
        Cells(lastRow, 1).Activate
        For InsertRows = 1 To numRows
            ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Copy
            Selection.Insert Shift:=xlDown
            ActiveCell.Range("A1:G1").ClearContents
        Next InsertRows
        Application.CutCopyMode = False
    
    Case "CarryoverKnown"
        lastRow = 1
        Do Until Cells(lastRow, 1).Text = "Total:"
            lastRow = lastRow + 1
        Loop
        Cells(lastRow, 1).Activate
        For InsertRows = 1 To numRows
            ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Copy
            Selection.Insert Shift:=xlDown
            ActiveCell.Range("A1:G1").ClearContents
        Next InsertRows
        Application.CutCopyMode = False
    
    Case "CarryoverLikely"
        lastRow = 1
        For iCount = 1 To 2
            Do Until Cells(lastRow, 1).Text = "Total:"
                lastRow = lastRow + 1
            Loop
            If iCount <> 2 Then lastRow = lastRow + 1
        Next iCount
        Cells(lastRow, 1).Activate
        For InsertRows = 1 To numRows
            ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Copy
            Selection.Insert Shift:=xlDown
            ActiveCell.Range("A1:G1").ClearContents
        Next InsertRows
        Application.CutCopyMode = False
    End Select
    
    SetFormulas
    
    ' reset dialog box (clean up prior to next activation)
    For Each optn In dlgWhere.OptionButtons
        optn.Value = xlOff
    Next
    dlgWhere.EditBoxes("txtNumRows").Text = ""
    
    ' activate protection and screen updating
    Cells(lastRow, 1).Select
    Application.ScreenUpdating = True
End Sub

'***************************************************************************

Sub SetFormulas()
Attribute SetFormulas.VB_ProcData.VB_Invoke_Func = " \n14"
        ThisWorkbook.Worksheets("Evaluation_of_Misstatements").Activate
  ' Insert appropriate formulas throughout that do not automatically update
        firstRow = 2
        Do Until Cells(firstRow - 1, 1).Text = "KNOWN MISSTATEMENTS"
            firstRow = firstRow + 1
        Loop
        lastRow = 1
        Do Until Cells(lastRow + 1, 1).Text = "Total known misstatements"
           lastRow = lastRow + 1
        Loop
        diffRow = (lastRow + 1) - firstRow
        ' set formula strings
        sumFormula = "=SUM(R[" & -diffRow & "]C:R[-1]C)"
        xFootFormula = "=XFoot(RC[-5]:RC[-1],R[" & -diffRow & "]C:R[-1]C)"
        Range("CY_knw_Assets").FormulaR1C1 = sumFormula
        Range("CY_knw_Liabs").FormulaR1C1 = sumFormula
        Range("CY_knw_RetEarn_bf").FormulaR1C1 = sumFormula
        Range("CY_knw_Equity").FormulaR1C1 = sumFormula
        Range("CY_knw_Income").FormulaR1C1 = sumFormula
        Range("Tot_knw_Xfoot").FormulaR1C1 = xFootFormula
        Range("CY_tx_knw_Liabs").FormulaR1C1 = sumFormula
        Range("CY_tx_knw_RetEarn_bf").FormulaR1C1 = sumFormula
        Range("CY_tx_knw_Equity").FormulaR1C1 = sumFormula
        Range("CY_tx_knw_Income").FormulaR1C1 = sumFormula
    
  ' Insert appropriate formulas throughout that do not automatically update
        firstRow = 2
        Do Until Cells(firstRow - 1, 1).Text = "LIKELY MISSTATEMENTS"
            firstRow = firstRow + 1
        Loop
        lastRow = 1
        Do Until Cells(lastRow + 1, 1).Text = "Total likely misstatements"
           lastRow = lastRow + 1
        Loop
        diffRow = (lastRow + 1) - firstRow
        ' set formula strings
        sumFormula = "=SUM(R[" & -diffRow & "]C:R[-1]C)"
        xFootFormula = "=XFoot(RC[-5]:RC[-1],R[" & -diffRow & "]C:R[-1]C)"
        Range("CY_lik_Assets").FormulaR1C1 = sumFormula
        Range("CY_lik_Liabs").FormulaR1C1 = sumFormula
        Range("CY_lik_RetEarn_bf").FormulaR1C1 = sumFormula
        Range("CY_lik_Equity").FormulaR1C1 = sumFormula
        Range("CY_lik_Income").FormulaR1C1 = sumFormula
        Range("Tot_lik_Xfoot").FormulaR1C1 = xFootFormula
        Range("CY_tx_lik_Liabs").FormulaR1C1 = sumFormula
        Range("CY_tx_lik_RetEarn_bf").FormulaR1C1 = sumFormula
        Range("CY_tx_lik_Equity").FormulaR1C1 = sumFormula
        Range("CY_tx_lik_Income").FormulaR1C1 = sumFormula

  ' Insert appropriate formulas throughout that do not automatically update
        firstRow = 2
        Do Until Cells(firstRow - 1, 1).Text = "Known Misstatements:"
            firstRow = firstRow + 1
        Loop
        lastRow = 1
        Do Until Cells(lastRow + 1, 1).Text = "Total:"
           lastRow = lastRow + 1
        Loop
        diffRow = (lastRow + 1) - firstRow
        ' set formula strings
        sumFormula = "=SUM(R[" & -diffRow & "]C:R[-1]C)"
        xFootFormula = "=XFoot(RC[-5]:RC[-1],R[" & -diffRow & "]C:R[-1]C)"
        Range("PY_knw_RetEarn").FormulaR1C1 = sumFormula
        Range("PY_knw_Income").FormulaR1C1 = sumFormula
        Range("PY_tot_knw_Xfoot").FormulaR1C1 = xFootFormula
        Range("PY_tx_knw_RetEarn").FormulaR1C1 = sumFormula
        Range("PY_tx_knw_Income").FormulaR1C1 = sumFormula

  ' Insert appropriate formulas throughout that do not automatically update
        firstRow = 2
        Do Until Cells(firstRow - 1, 1).Text = "Likely Misstatements:"
            firstRow = firstRow + 1
        Loop
        lastRow = firstRow
        Do Until Cells(lastRow + 1, 1).Text = "Total:"
           lastRow = lastRow + 1
        Loop
        diffRow = (lastRow + 1) - firstRow
        ' set formula strings
        sumFormula = "=SUM(R[" & -diffRow & "]C:R[-1]C)"
        xFootFormula = "=XFoot(RC[-5]:RC[-1],R[" & -diffRow & "]C:R[-1]C)"
        Range("PY_lik_RetEarn").FormulaR1C1 = sumFormula
        Range("PY_lik_Income").FormulaR1C1 = sumFormula
        Range("PY_tot_lik_Xfoot").FormulaR1C1 = xFootFormula
        Range("PY_tx_lik_RetEarn").FormulaR1C1 = sumFormula
        Range("PY_tx_lik_Income").FormulaR1C1 = sumFormula
End Sub
