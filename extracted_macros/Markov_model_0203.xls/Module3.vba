Attribute VB_Name = "Module3"
Sub Button7_Click()
Range("G13").Value = 0
Range("G22").Value = 0
Range("G24").Value = 0
Range("deltaretent").Value = "=sheet3!G13"
Range("deltainput").Value = "=sheet3!G22"
Range("deltainputb").Value = "=sheet3!G24"
End Sub
Sub Macro1()
Attribute Macro1.VB_Description = "Macro recorded 3/8/99 by Philip Batty"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' Macro recorded 3/8/99 by Philip Batty
'

'
    ActiveSheet.Next.Select
    ActiveWindow.LargeScroll Down:=0
    Range("F26").Select
    ActiveCell.FormulaR1C1 = "=Sheet3!R[-13]C[1]"
    Range("F27").Select
    ActiveWindow.LargeScroll Down:=1
    Range("F46").Select
    ActiveCell.FormulaR1C1 = "=Sheet3!R[-30]C[1]"
    Range("F47").Select
End Sub
