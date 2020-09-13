Attribute VB_Name = "Module1"
Sub InsertNewDetailLine()
Attribute InsertNewDetailLine.VB_Description = "Macro recorded 2000-07-31 by George Smith"
Attribute InsertNewDetailLine.VB_ProcData.VB_Invoke_Func = " \n14"
'
' InsertNewDetailLine Macro
' Macro recorded 2000-07-31 by George Smith
'
' Keyboard Shortcut: Ctrl+i
'
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "(Replace with detail)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "0"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Range("A1:J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 10).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Range("A1:J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 10).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Range("A1:J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 10).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    ActiveCell.Offset(0, -41).Range("A1").Select
End Sub
