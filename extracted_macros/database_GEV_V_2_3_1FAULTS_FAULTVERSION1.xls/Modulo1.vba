Attribute VB_Name = "Modulo1"
Sub visualizza_singolo()
Attribute visualizza_singolo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' visualizza_singolo Macro
' Macro registrata il 10/03/2002 da Antonello Lobianco
'
' Scelta rapida da tastiera: CTRL+v
'

 'Imposta i valori dei "prendi tutto"
 'Gev
 Foglio6.Cells(21, 4) = "FALSE"
 'Mese
  Foglio6.Cells(22, 4) = "FALSE"
 'Tipo di servizio
  Foglio6.Cells(23, 4) = "TRUE"
 'Macchina
  Foglio6.Cells(21, 5).Value = "TRUE"
  
  'Calcola il numero di servizi, ore di servizio e numero sanzioni per mese:
  'Imposta il mese
  Count = 1
  
  While (Count < 13)
    Foglio6.Cells(22, 2).Value = Count
    n_serv = Foglio6.Cells(27, 2)
    h_serv = Foglio6.Cells(28, 2)
    n_sanz = Foglio6.Cells(29, 2)
    cella_mese = Count + 2
    Foglio3.Cells(9, cella_mese) = n_serv
    Foglio3.Cells(10, cella_mese) = h_serv
    Foglio3.Cells(11, cella_mese) = n_sanz
    
    Count = Count + 1
  Wend
  
  'Ora ci occupiamo dei dati annuali
  'Prende tutti  i mesi
  Foglio6.Cells(22, 4) = "TRUE"
  
  Foglio3.Cells(17, 6) = Foglio6.Cells(33, 2)
  Foglio3.Cells(18, 6) = Foglio6.Cells(34, 2)
  Foglio3.Cells(19, 6) = Foglio6.Cells(35, 2)
  Foglio3.Cells(20, 6) = Foglio6.Cells(36, 2)
  Foglio3.Cells(21, 6) = Foglio6.Cells(37, 2)
  Foglio3.Cells(22, 6) = Foglio6.Cells(38, 2)
  Foglio3.Cells(23, 6) = Foglio6.Cells(39, 2)
  Foglio3.Cells(24, 6) = Foglio6.Cells(40, 2)
  Foglio3.Cells(25, 6) = Foglio6.Cells(41, 2)
  Foglio3.Cells(26, 6) = Foglio6.Cells(42, 2)
  Foglio3.Cells(27, 6) = Foglio6.Cells(43, 2)
  Foglio3.Cells(28, 6) = Foglio6.Cells(44, 2)
  Foglio3.Cells(23, 13) = Foglio6.Cells(45, 2)
  Foglio3.Cells(22, 13) = Foglio6.Cells(33, 4)
  
  'In particolare ora ci occupiamo dei tipi di servizi
  'Non prende più tutti i servizi contemporaneamente
  Foglio6.Cells(23, 4) = "FALSE"
  
  'Imposta il tipo di servizio
  Count = 1
  
  While (Count < 5)
    Foglio6.Cells(23, 2).Value = Count
    n_serv = Foglio6.Cells(33, 4)
    ore_serv = Foglio6.Cells(33, 7)
    cella_serv = Count + 16
    Foglio3.Cells(cella_serv, 13) = n_serv
    Foglio3.Cells(cella_serv, 14) = ore_serv
    Count = Count + 1
  Wend
  'Infine ora vediamo i km
  'Prendo nuovamente tutti i servizi, ma rendo attivo il filtro macchina
  Foglio6.Cells(23, 4) = "TRUE"
  Foglio6.Cells(21, 5).Value = "FALSE"
  Foglio3.Cells(24, 13) = Foglio6.Cells(35, 4)
  
End Sub
Sub visualizza_gruppo()
Attribute visualizza_gruppo.VB_ProcData.VB_Invoke_Func = " \n14"

'Imposta i valori dei "prendi tutto"
 'Gev
  Foglio6.Cells(21, 4) = "true"
 'Mese
  Foglio6.Cells(22, 4) = "FALSE"
 'Macchina
  Foglio6.Cells(21, 5).Value = "TRUE"
  
'Imposta il mese
  count_mese = 1
'Incomincia il ciclo del mese
  While (count_mese < 13)
     Foglio6.Cells(22, 2) = count_mese
     riga_serv = count_mese + 6
     riga_sanz = count_mese + 25
     'Imposta il tipo di servizio
     count_serv = 1
     'Tipo di servizio
     Foglio6.Cells(23, 4) = "false"
     While (count_serv < 5)

        Foglio6.Cells(23, 2).Value = count_serv
        n_serv_col = 3 * count_serv
        h_serv_col = 3 * count_serv + 1
        km_col = 3 * count_serv + 2
        Foglio7.Cells(riga_serv, n_serv_col) = Foglio6.Cells(33, 4)
        Foglio7.Cells(riga_serv, h_serv_col) = Foglio6.Cells(48, 2)
        Foglio7.Cells(riga_serv, km_col) = Foglio6.Cells(35, 4)
        count_serv = count_serv + 1
     Wend
     
     'Ora ci occupiamo delle sanzioni e delle segnalazioni
     'Tipo di servizio
     Foglio6.Cells(23, 4) = "true"
     
     Foglio7.Cells(riga_sanz, 3) = Foglio6.Cells(33, 2)
     Foglio7.Cells(riga_sanz, 4) = Foglio6.Cells(34, 2)
     Foglio7.Cells(riga_sanz, 5) = Foglio6.Cells(35, 2)
     Foglio7.Cells(riga_sanz, 6) = Foglio6.Cells(36, 2)
     Foglio7.Cells(riga_sanz, 7) = Foglio6.Cells(37, 2)
     Foglio7.Cells(riga_sanz, 8) = Foglio6.Cells(38, 2)
     Foglio7.Cells(riga_sanz, 9) = Foglio6.Cells(39, 2)
     Foglio7.Cells(riga_sanz, 10) = Foglio6.Cells(40, 2)
     Foglio7.Cells(riga_sanz, 11) = Foglio6.Cells(41, 2)
     Foglio7.Cells(riga_sanz, 12) = Foglio6.Cells(42, 2)
     Foglio7.Cells(riga_sanz, 13) = Foglio6.Cells(43, 2)
     Foglio7.Cells(riga_sanz, 14) = Foglio6.Cells(44, 2)
     
    
    
    count_mese = count_mese + 1
  Wend
End Sub
