Attribute VB_Name = "Modulo4"
Sub rimuovi_servizio()
Attribute rimuovi_servizio.VB_Description = "Macro registrata il 16/03/2002 da Antonello Lobianco"
Attribute rimuovi_servizio.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' rimuovi_servizio Macro
' Macro registrata il 16/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+z
'
    rig = Foglio6.Cells(56, 3)
    riga2 = rig + 2
    n_serv = Foglio6.Cells(60, 2)
    controllo = rig - n_serv
    If (controllo = 2) Then
    Else
       If (riga2 > 3) Then
    
       Foglio5.Rows(riga2).Delete Shift:=xlUp
       
       End If
    End If
    
    Sheets("SetPar").Select
    Range("A1").Select
End Sub
