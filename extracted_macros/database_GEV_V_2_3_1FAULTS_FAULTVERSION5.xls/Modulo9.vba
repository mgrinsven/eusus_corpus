Attribute VB_Name = "Modulo9"
Sub cancella_ultimo_inserimento()
Attribute cancella_ultimo_inserimento.VB_Description = "Macro registrata il 10/03/2002 da Antonello Lobianco"
Attribute cancella_ultimo_inserimento.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' cancella_ultimo_inserimento Macro
' Macro registrata il 10/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+c
'
    n_serv = Foglio6.Cells(60, 2)
    If (n_serv > 0) Then
    Sheets("SetPar").Select
    Range("A1").Select
    Sheets("db").Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    End If
    Sheets("immissione dati").Select
    Range("H24").Select
    Selection.ClearContents
    Range("E22:E33").Select
    Selection.ClearContents
    Sheets("calcoli").Select
    Range("J11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("L11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("AA11").Select
    ActiveCell.FormulaR1C1 = "1"
    Sheets("immissione dati").Select
    ActiveWindow.SmallScroll Down:=-13
    Range("D6").Select
    Selection.ClearContents
    Range("D6").Select
    
    
End Sub
Sub salva_ed_esci()
Attribute salva_ed_esci.VB_Description = "Macro registrata il 10/03/2002 da Antonello Lobianco"
Attribute salva_ed_esci.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' salva_ed_esci Macro
' Macro registrata il 10/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+s
'
    Sheets("Home").Select
    ActiveWorkbook.Save
    Application.ActiveWindow.Close
End Sub
