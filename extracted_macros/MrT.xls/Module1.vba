Attribute VB_Name = "Module1"
Option Explicit
Sub GetAccount()
Attribute GetAccount.VB_Description = "Macro recorded 9/27/98 by Alastair Bor"
Attribute GetAccount.VB_ProcData.VB_Invoke_Func = " \n14"
'(C)1998 Alastair M. Bor

'This software may be freely distributed so long as the copyright information is unchanged
'This software is not warranted in any way

    Application.ScreenUpdating = False
    Range("A2").Copy
    ActiveSheet.Paste
    Selection.Font.ColorIndex = 1
    Selection.Interior.ColorIndex = xlNone
End Sub
Sub GenerateT()
Attribute GenerateT.VB_Description = "Macro recorded 9/27/98 by Alastair Bor"
Attribute GenerateT.VB_ProcData.VB_Invoke_Func = " \n14"
'(c)1998 Alastair M. Bor

'This software may be freely distributed so long as the copyright information is unchanged
'This software is not warranted in any way

Dim inputdata(101, 3) As Variant  '100 records, 4 attributes (trans, amt, acct)
Dim accounts(60, 2) As Variant      '37 different accounts, 2 attributes (type, name)
Dim RowA As Integer
Dim RowAs As Integer
Dim RowLi As Integer
Dim RowSt As Integer
Dim ActSumL As Single
Dim ActSumR As Single
Dim offset As Integer
Dim i As Integer
Dim TraNum As Integer
Dim AccNum As Integer
Dim AccNam As String
Dim AccTyp As String
Dim AccTypOffset As Integer
Dim Amt As Single
Dim OldNum As Integer

Application.ScreenUpdating = False

Sheets("Input").Select
If Not (Cells(4, 13) = 0) Then
    MsgBox ("Warning: Audit reveals an error")
    End
End If

Range("A6:M106").Sort Key1:=Range("B6"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

Sheets("Categories").Select
For i = 1 To 60
    accounts(i, 1) = Cells(i + 1, 4) 'Name
    accounts(i, 2) = Cells(i + 1, 3) 'Type
Next

Sheets("Input").Select
For i = 1 To 100
    inputdata(i, 1) = Cells(i + 5, 1)   'Transaction number
    inputdata(i, 2) = Cells(i + 5, 2)   'Account number
    inputdata(i, 3) = Cells(i + 5, 3)   'Amount
Next

Application.ScreenUpdating = True

'===================SET UP TITLES==========================

Sheets("TTables").Select        'Switch to output screen
Cells.Delete Shift:=xlUp        'Clear screen
    
Range("A1").ColumnWidth = 2
Range("B1").FormulaR1C1 = "'Assets"
Range("D1").ColumnWidth = 2
Range("F1").ColumnWidth = 2
Range("G1").FormulaR1C1 = "'Liabilities"
Range("I1").ColumnWidth = 2
Range("K1").ColumnWidth = 2
Range("L1").FormulaR1C1 = "'Stockholders Equity"
Range("N1").ColumnWidth = 2
    
Range("L1:M1").Select
With Selection
    .HorizontalAlignment = xlCenter
    .MergeCells = True
    .Font.Bold = True
End With
 
Range("G1:H1").Select
With Selection
    .HorizontalAlignment = xlCenter
    .MergeCells = True
    .Font.Bold = True
End With
    
Range("B1:C1").Select
With Selection
    .HorizontalAlignment = xlCenter
    .MergeCells = True
    .Font.Bold = True
End With

'===================CREATE INDIVIDUAL T TABLES==========================
RowA = 0
RowAs = 3
RowLi = 3
RowSt = 3
OldNum = 0
ActSumL = 0
ActSumR = 0

For i = 1 To 100
    TraNum = inputdata(i, 1)            'Transaction Number
    AccNum = inputdata(i, 2)            'Account Number
    AccNam = accounts(AccNum, 1)        'Account Name
    AccTyp = accounts(AccNum, 2)        'Account type (Asset, Liability, SE)
    If AccTyp = "A" Then
        AccTypOffset = 0
        RowA = RowAs
    End If
    If AccTyp = "L" Then
        AccTypOffset = 5
        RowA = RowLi
    End If
    If AccTyp = "S" Then
        AccTypOffset = 10
        RowA = RowSt
    End If
    Amt = inputdata(i, 3)               'Transaction amount
    Application.ScreenUpdating = False
        If OldNum < AccNum Then         'Its a new account, make a heading
            Cells(RowA, 2 + AccTypOffset) = AccNam
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 3 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            Cells(RowA, 2 + AccTypOffset).HorizontalAlignment = xlCenter
            Cells(RowA, 2 + AccTypOffset).Font.Italic = True
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 3 + AccTypOffset)).Merge
            RowA = RowA + 1
            If AccTypOffset = 0 Then
                Cells(RowA, 2 + AccTypOffset) = "Dr. (+)"
                Cells(RowA, 2 + AccTypOffset).HorizontalAlignment = xlCenter
                Cells(RowA, 3 + AccTypOffset) = "Cr. (-)"
                Cells(RowA, 3 + AccTypOffset).HorizontalAlignment = xlCenter
            End If
            If AccTypOffset > 0 Then
                Cells(RowA, 2 + AccTypOffset) = "Dr. (-)"
                Cells(RowA, 2 + AccTypOffset).HorizontalAlignment = xlCenter
                Cells(RowA, 3 + AccTypOffset) = "Cr. (+)"
                Cells(RowA, 3 + AccTypOffset).HorizontalAlignment = xlCenter
            End If
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 3 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            RowA = RowA + 1
        End If
        If (AccTyp = "A" And Amt >= 0) Or ((AccTyp = "L" Or AccTyp = "S") And Amt < 0) Then
        'It goes on LEFT
            If Cells(RowA, 1 + AccTypOffset) > 0 Then RowA = RowA + 1
            Cells(RowA, 1 + AccTypOffset) = TraNum
            Cells(RowA, 2 + AccTypOffset) = Abs(Amt)
            OldNum = AccNum
            ActSumL = ActSumL + Abs(Amt)
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 2 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            Selection.Style = "Comma"
            Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
            Application.ScreenUpdating = True
            End With
        End If
        If (AccTyp = "A" And Amt < 0) Or ((AccTyp = "L" Or AccTyp = "S") And Amt >= 0) Then
        'It goes on Right
            If Cells(RowA, 4 + AccTypOffset) > 0 Then RowA = RowA + 1
            Cells(RowA, 3 + AccTypOffset) = Abs(Amt)
            Cells(RowA, 4 + AccTypOffset) = TraNum
            OldNum = AccNum
            ActSumR = ActSumR + Abs(Amt)
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 2 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            Range(Cells(RowA, 3 + AccTypOffset), Cells(RowA, 3 + AccTypOffset)).Select
            Selection.Style = "Comma"
            Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
            Application.ScreenUpdating = True
        End If
        If i = 100 Or inputdata(i + 1, 2) > AccNum Then 'Add them up 'cause this acct is done
            RowA = RowA + 1
            Cells(RowA, 2 + AccTypOffset) = ActSumL
            Cells(RowA, 3 + AccTypOffset) = ActSumR
            ActSumL = 0
            ActSumR = 0
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 3 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            Selection.Style = "Comma"
            Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .Weight = xlThick
            End With
            Range(Cells(RowA, 2 + AccTypOffset), Cells(RowA, 2 + AccTypOffset)).Select
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            RowA = RowA + 4
        End If
        If AccTyp = "A" Then RowAs = RowA
        If AccTyp = "L" Then RowLi = RowA
        If AccTyp = "S" Then RowSt = RowA
Next

End Sub
Sub CleanInput()
Attribute CleanInput.VB_Description = "Macro recorded 9/28/98 by Alastair Bor"
Attribute CleanInput.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Msg, Style, Title, Response
    Msg = "Are you sure you want to erase all data?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Clear All Data"
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then    ' User chose Yes.
        Range("A6:C106").ClearContents
    Else
        End
    End If
End Sub
