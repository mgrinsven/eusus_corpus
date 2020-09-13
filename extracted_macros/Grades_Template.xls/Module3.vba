Attribute VB_Name = "Module3"
Sub HWInsert()
Attribute HWInsert.VB_Description = "Macro recorded 2/4/2004 by Scott Bartuska"
Attribute HWInsert.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("HWInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("HWInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub

Sub LabInsert()
    Range("LabInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("LabInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub

Sub TestInsert()
    Range("TestInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("TestInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub

Sub MidInsert()
    Range("MidInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("MidInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub

Sub FinalInsert()
    Range("FinalInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("FinalInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub

Sub QuizInsert()
    Range("QuizInsert").Select
    Selection.EntireRow.Insert
    ActiveWindow.LargeScroll ToRight:=1
    Range("P1:T1").Select
    Selection.Copy
    ActiveWindow.LargeScroll ToRight:=-1
    Range("QuizInsert").Select
    ActiveCell.Offset(-1, 0).Select
    ActiveSheet.Paste
End Sub
