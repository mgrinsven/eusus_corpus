Attribute VB_Name = "Modulo6"
Sub inserisci_gev()
Attribute inserisci_gev.VB_Description = "Macro registrata il 16/03/2002 da Antonello Lobianco"
Attribute inserisci_gev.VB_ProcData.VB_Invoke_Func = "v\n14"
'
' inserisci_gev Macro
' Macro registrata il 16/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+v
'
    If (Foglio8.Cells(7, 5) = "") Then
    Else
    Range("E7").Select
    Selection.Copy
    Sheets("parametri").Select
    Range("A2").Select
    Selection.Insert Shift:=xlDown
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    Sheets("SetPar").Select
    Range("E7").Select

End Sub
