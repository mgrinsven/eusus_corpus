Attribute VB_Name = "Module1"
Public Function LastNonZero(ThisRow As Integer, TstRange)
Attribute LastNonZero.VB_ProcData.VB_Invoke_Func = " \n14"
Const TopRow As Integer = 10
Const BotRow As Integer = 40
Const ColumnNum As Integer = 15     ' column O
Dim Text, counter As Integer
Text = 0
counter = ThisRow - 1
Do
    With ActiveSheet.Cells(counter, ColumnNum)
        If Application.WorksheetFunction.IsNumber(.Value) Then
            LastNonZero = .Value: Exit Function
            ElseIf .Value <> "" Then
                Text = 1
        End If
    End With
    counter = counter - 1
Loop Until counter < TopRow
If Text = 0 Then
    LastNonZero = ""
    Else
        LastNonZero = "Bad End"
End If
End Function
Public Function FirstNonZero(ThisRow As Integer, TstRange)
Attribute FirstNonZero.VB_ProcData.VB_Invoke_Func = " \n14"
Const TopRow As Integer = 10
Const BotRow As Integer = 40
Const ColumnNum As Integer = 2     ' column B
Dim Text, counter As Integer
Text = 0
counter = TopRow
Do
    With Cells(counter, ColumnNum)
        If Application.WorksheetFunction.IsNumber(.Value) Then
            FirstNonZero = .Value: Exit Function
            ElseIf .Value <> "" Then
                Text = 1
        End If
    End With
    counter = counter + 1
Loop Until counter > BotRow
If Text = 0 Then
    FirstNonZero = ""
    Else
        FirstNonZero = "No Start"
End If
End Function
Public Function Conclusion(LeakCheck As Integer, OverShort As Integer)
Attribute Conclusion.VB_ProcData.VB_Invoke_Func = " \n14"
With Range("A52").Font
If LeakCheck > OverShort Then
    Conclusion = "YES"
        .Name = "Tahoma"
        .Size = 16
        .Color = vbMagenta
    Else
    Conclusion = "No"
        .Name = "Arial"
        .Size = 12
        .Color = vbBlack
End If
End With
End Function
Public Function Conclusion2(LeakCheck As Integer, OverShort As Integer, ThisCell)
Attribute Conclusion2.VB_ProcData.VB_Invoke_Func = " \n14"
Worksheets(ActiveSheet).Activate
With Range(ThisCell).Font
If LeakCheck > OverShort Then
        .Name = "Tahoma"
        .Size = 16
        .Color = vbMagenta
    Conclusion = "YES"
    Else
        .Name = "Arial"
        .Size = 14
        .Color = vbBlack
    Conclusion = "No"
End If
End With
End Function
