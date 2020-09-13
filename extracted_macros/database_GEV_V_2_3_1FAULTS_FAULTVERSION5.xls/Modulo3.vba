Attribute VB_Name = "Modulo3"
Sub stampa()
Attribute stampa.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute stampa.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' stampa Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+t
'
    Range("A1").Select
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
        :=True
End Sub
Sub go_inserimento()
Attribute go_inserimento.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute go_inserimento.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' go_inserimento Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+n
'
    Sheets("immissione dati").Select
    Range("D6").Select
End Sub
Sub go_visualizza_gev()
Attribute go_visualizza_gev.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute go_visualizza_gev.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' go_visualizza_gev Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+g
'
    Sheets("visualizza_singolo").Select
    Range("A1").Select
End Sub
Sub go_visualizza_gruppo()
Attribute go_visualizza_gruppo.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute go_visualizza_gruppo.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' go_visualizza_gruppo Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+r
'
    Sheets("visualizza_gruppo").Select
    Range("A1").Select
End Sub
Sub go_parametri()
Attribute go_parametri.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute go_parametri.VB_ProcData.VB_Invoke_Func = "o\n14"
'
' go_parametri Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+o
'
    Sheets("SetPar").Select
    Range("A1").Select
End Sub
Sub go_help()
Attribute go_help.VB_Description = "Macro registrata il 15/03/2002 da Antonello Lobianco"
Attribute go_help.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' go_help Macro
' Macro registrata il 15/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+e
'
    Sheets("Help").Select
    Range("A1").Select
End Sub

Sub cancella_gev()
Attribute cancella_gev.VB_Description = "Macro registrata il 16/03/2002 da Antonello Lobianco"
Attribute cancella_gev.VB_ProcData.VB_Invoke_Func = "l\n14"
'
' cancella_gev Macro
' Macro registrata il 16/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+l
'
    riga = Foglio6.Cells(52, 3)
    'riga2 = rig + 7
    n_gev = Foglio6.Cells(61, 2)
    ultima_riga = 2 + n_gev
    If (riga = 1) Then
    Else
      If (riga = ultima_riga) Then
      Else
         Foglio4.Cells(riga, 1).Delete Shift:=xlUp
       End If
    End If
    Sheets("SetPar").Select
End Sub
