Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_Description = "Macro recorded 5/29/2003 by JW Lewis"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' Macro recorded 5/29/2003 by JW Lewis
'

'
   With ActiveWindow
        .Top = 4
        .Left = 2.5
    End With
    Application.Left = 136.75
    Application.Top = 13.75
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Range("F6").Select
    Application.CommandBars("Stop Recording").Visible = True
End Sub
