Attribute VB_Name = "Module1"
Sub update()
    Sheets("Papers").Select
    Columns("A:K").Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=2, MatchCase:=False, Orientation:=xlTopToBottom

    Sheets("Tuesday").Select
    Call updateSub(1, 12, 14)
    Call updateSub(2, 16, 19)
    Sheets("Wednesday").Select
    Call updateSub(1, 10, 12)
    Call updateSub(2, 14, 16)
    Call updateSub(3, 18, 20)
    Call updateSub(4, 22, 25)
    Sheets("Thursday").Select
    Call updateSub(1, 11, 13)
    Call updateSub(2, 15, 18)
    
    Sheets("Papers").Select
    Columns("A:K").Sort Key1:=Range("E2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=2, MatchCase:=False, Orientation:=xlTopToBottom
    Columns("A:K").Sort Key1:=Range("D2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=2, MatchCase:=False, Orientation:=xlTopToBottom
    Columns("A:K").Sort Key1:=Range("C2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=2, MatchCase:=False, Orientation:=xlTopToBottom
    Columns("A:K").Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=2, MatchCase:=False, Orientation:=xlTopToBottom
     
End Sub

Sub updateSub(session, row1, row2)
    Dim paperCell As Range
    sheetName = Mid(ActiveSheet.Name, 1, 3)
    For col = 6 To 11
        track = col - 5
        paper = 0
        For row = row1 To row2
            paper = paper + 1
            paperID = Cells(row, col)
            If TypeName(Cells(row, col).Value) = "Double" Then
                Debug.Print sheetName, row, col, session, paper, paperID, TypeName(paperID)
                Set paperRow = Range("papers!a:a").Find(paperID)
                If Not paperRow Is Nothing Then
                    Debug.Print paperRow.row, paperRow.Value
                    chair = nameLookup(sheetName, track, session)
                    paperRow.Offset(0, 1) = sheetName
                    paperRow.Offset(0, 2) = track
                    paperRow.Offset(0, 3) = session
                    paperRow.Offset(0, 4) = paper
                    paperRow.Offset(0, 5) = chair
                End If
            End If
        Next row
    Next col
End Sub

Function nameLookup(day, track, session)
    If day = "Tue" Then
        nameLookup = Range("'Session Chairs'!a:l").Cells(3 * session + 1, track + 1).Value
    ElseIf day = "Wed" Then
        nameLookup = Range("'Session Chairs'!a:l").Cells(3 * session + 2 * 3 + 1, track + 1).Value
    ElseIf day = "Thu" Then
        nameLookup = Range("'Session Chairs'!a:l").Cells(3 * session + 6 * 3 + 1, track + 1).Value
    End If
End Function
