Attribute VB_Name = "Modulo8"
Sub inserimanto_dati()
Attribute inserimanto_dati.VB_Description = "Macro registrata il 10/03/2002 da Antonello Lobianco"
Attribute inserimanto_dati.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' inserimanto_dati Macro
' Macro registrata il 10/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+i
'
    Sheets("calcoli").Select
    Range("A15:AC15").Select
    Selection.Copy
    Range("A16").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Range("J11").Select
    Application.CutCopyMode = False
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
    Rows("16:16").Select
    Selection.Copy
    Sheets("db").Select
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Sheets("immissione dati").Select
    Application.CutCopyMode = False
    'Selection.ClearContents
    ActiveWindow.SmallScroll Down:=14
    Range("E22:E33").Select
    Selection.ClearContents
    Range("H24").Select
    Selection.ClearContents
    ActiveWindow.LargeScroll Down:=-1
    Range("D6").Select
    Selection.ClearContents
End Sub
