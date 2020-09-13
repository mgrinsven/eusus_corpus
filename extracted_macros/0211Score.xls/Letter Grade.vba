Attribute VB_Name = "Letter Grade"



























































Function abc(Loc)
Attribute abc.VB_ProcData.VB_Invoke_Func = " \n14"
    
        If Loc >= 462 Then
            abc = "  A"
        ElseIf Loc >= 450 And Loc <= 461 Then
            abc = "  A-"
        ElseIf Loc >= 433 And Loc <= 449 Then
            abc = "  B+"
        ElseIf Loc >= 416 And Loc <= 432 Then
            abc = "  B"
        ElseIf Loc >= 400 And Loc <= 415 Then
            abc = "  B-"
        ElseIf Loc >= 383 And Loc <= 399 Then
            abc = "  C+"
        ElseIf Loc >= 366 And Loc <= 382 Then
            abc = "  C"
        ElseIf Loc >= 350 And Loc <= 365 Then
            abc = "  C-"
        ElseIf Loc >= 333 And Loc <= 349 Then
            abc = "  D+"
        ElseIf Loc >= 316 And Loc <= 332 Then
            abc = "  D"
        ElseIf Loc >= 300 And Loc <= 315 Then
            abc = "  D-"
        ElseIf Loc < 300 Then
            abc = "  F"
    End If

End Function



