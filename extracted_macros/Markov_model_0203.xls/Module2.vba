Attribute VB_Name = "Module2"
Sub Spinner3_Change()
If Range("H22").Value = 99 Then Range("G22").Value = Range("G22").Value - 0.001
If Range("H22").Value = 101 Then Range("G22").Value = Range("G22").Value + 0.001
Range("H22").Value = 100
End Sub
Sub Spinner4_Change()
If Range("H24").Value = 99 Then Range("G24").Value = Range("G24").Value - 0.001
If Range("H24").Value = 101 Then Range("G24").Value = Range("G24").Value + 0.001
Range("H24").Value = 100
End Sub

Sub Button6_Click()
Attribute Button6_Click.VB_Description = "Macro recorded 3/5/99 by Philip Batty"
Attribute Button6_Click.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Button6_Click Macro
' Macro recorded 3/5/99 by Philip Batty
'

'
    ActiveSheet.Next.Select
    UserForm1.Show
End Sub
