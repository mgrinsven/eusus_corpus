Attribute VB_Name = "Module1"
Public Maxthrow As Integer
Public Maxperiod As Integer
Sub Playgame()

    Dim Throw As Integer
    Dim Startcell As Range
    Dim Endcell As Range
    Throw = Range("A24").Value
    Maxthrow = Range("R22").Value
    j = 0
    k = 0
    'If max # rounds to be played is larger than Maxperiod, the game can't be played
    If Maxthrow > Maxperiod Then
        response = MsgBox("The maximum number of periods is " & Maxperiod, vbOKOnly + vbExclamation, "Dice game")
        Range("R22").Select
    Else
        'If max # rounds have been played, the game is finished
        If Throw >= Maxthrow Then
            response = MsgBox("The game is finished", vbOKOnly + vbInformation, "Dice Game")
        'Calculate the inventories and numbers produced
        Else
            Range("A1:S24").Calculate
            Do While j <= 15 'copy last score to table
                Cells(26 + Throw, 2 + j).Value = Cells(4, 2 + j).Value
                j = j + 3
            Loop
            Range("A24").Value = Throw + 1
            Set Startcell = Range("B26").Offset(Throw, 0)
            Set Endcell = Startcell.Offset(0, 19)
            Range(Startcell, Endcell).Calculate
            Do While k <= 12 'copy the inventory values to the top
                Cells(6, 5 + k).Value = Cells(26 + Throw, 4 + k).Value
                k = k + 3
            Loop
            Cells(17, 19).Value = Cells(26 + Throw, 21).Value 'copy the avg WIP to the top
            Set Startcell = Range("B26").Offset(Throw, 0)
            Set Endcell = Startcell.Offset(0, 19)
            Range(Startcell, Endcell).Calculate
            Range("B8:S15").Calculate
        End If
    End If
End Sub
Sub ResetThrow()
Attribute ResetThrow.VB_Description = "Macro recorded 03/20/2002 by Faculteit ETEW"
Attribute ResetThrow.VB_ProcData.VB_Invoke_Func = " \n14"
    Maxperiod = 5000
    Maxthrow = Range("r22").Value
    Range("E6:T7,b26:b28,e26:e28,h26:h28,k26:k28,n26:n28,q26:q28").ClearContents
    Range("A24").Value = 0
    Range(Cells(29, 1), Cells(Maxperiod, 21)).ClearContents
    Range("A27:U28").AutoFill Destination:=Range(Cells(27, 1), Cells(Maxthrow + 25, 21)), Type:=xlFillDefault
    Range("A24").Select
    Calculate
End Sub

Sub ThrowAll()

    Dim Throw As Integer
    Maxperiod = 5000
    Throw = 0
    Maxthrow = Range("R22").Value
    'If max # rounds to be played is larger than Maxperiod, the game can't be played
    If Maxthrow > Maxperiod Then
        response = MsgBox("The maximum number of periods is " & Maxperiod, vbOKOnly + vbExclamation, "Dice game")
        Range("R22").Select
    Else
        response = MsgBox("The current game will be reset. Are you sure?", vbYesNoCancel + vbExclamation + vbDefaultButton2, "Dice Game")
        If response = vbYes Then
            ResetThrow
            Do While Throw < Maxthrow + 1
                Playgame
                Throw = Throw + 1
            Loop
        End If
        Calculate
    End If
End Sub
