Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 63, 4, MSForms, CommandButton"
Attribute VB_Control = "SORTTHEDATA, 60, 5, MSForms, CommandButton"
Attribute VB_Control = "Find_Boats, 59, 6, MSForms, CommandButton"
Attribute VB_Control = "ListBox3, 40, 7, MSForms, ListBox"
Attribute VB_Control = "ListBox2, 38, 8, MSForms, ListBox"



Private Sub CommandButton1_Click()
    Sheets("INPUTS").Select
    Range("A1").Select
    ActiveSheet.CALCULATE
    Rem Sheets("data base").Select
    Rem Range("A1").Select
    Rem ActiveSheet.CALCULATE
    Rem  Sheets("inputs").Select
    Rem Range("A1").Select
End Sub

Private Sub Find_Boats_Click()
    Sheets("data base").Select
    ActiveSheet.CALCULATE
    Sheets("inputs").Select
    Range("L13:N310").Select ' clears first 300 rows
    Selection.ClearContents
    y = 8
    For x = 8 To 1300
    If Sheets("data base").Cells(x, 3).value > 0 Then
         Sheets("inputs").Cells(y + 5, 14).value = Sheets("data base").Cells(x, 3).value
         Sheets("inputs").Cells(y + 5, 13).value = Sheets("data base").Cells(x, 2).value
         Sheets("inputs").Cells(y + 5, 12).value = Sheets("data base").Cells(x, 1).value
         y = y + 1
    End If
    Next x
    Sheets("INPUTS").Select
    Range("L12").Select
End Sub

Private Sub ListBox3_Click()

End Sub

Private Sub SORTTHEDATA_Click()
 Range("L13:N314").Select
    Selection.Sort Key1:=Range("N13"), Order1:=xlDescending, Header:=xlGuess _
        , OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("L9").Select
End Sub
