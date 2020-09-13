Attribute VB_Name = "Module1"
Sub MakeCards()
Attribute MakeCards.VB_Description = "Macro recorded 29/01/2002 by Odds"
Attribute MakeCards.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' MakeCards Macro
' Macro recorded 29/01/2002 by Odds
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'goto gards and delete existing cards
Application.ScreenUpdating = False
    Sheets("Cards").Select
    Columns("A:A").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.ResetAllPageBreaks
    ActiveWindow.View = xlNormalView
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("A1").Select
'Go to first ref no of Database
    Sheets("Database").Select
        Range("B2").Select
'set up check for ref no exists
Do While ActiveCell <> ""
    ActiveCell.Offset(0, -1).Select
    'check if reference selected
    Do While ActiveCell <> ""
    'copy, paste and format details in Cards
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(0, 1).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(1, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C&"" (""&R[-1]C[1]&"")"""
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(-1, 0).Range("A1:B1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    ActiveCell.Offset(3, 0).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(0, 1).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(0, 1).Range("A1").Select
    Sheets("Database").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Cards").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(1, -2).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C&"", ""&R[-1]C[1]&"", ""&R[-1]C[2]"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset(-1, 0).Range("A1:C1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    ActiveCell.Offset(-3, 0).Range("A1:A5").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    ActiveCell.Offset(5, 0).Range("A1").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    Sheets("Database").Select
    'goto next cell in Mark? column
    ActiveCell.Offset(1, -7).Range("A1").Select
    Loop
'nothing in mark cell so go to ref no and check out
    ActiveCell.Offset(1, 1).Select
    Loop
'finish up
    Sheets("Cards").Select
    ActiveSheet.DrawingObjects.Select
    Selection.Cut
    Range("C1").Select
    ActiveSheet.Paste
    Range("A1").Select
Application.ScreenUpdating = True
End Sub
