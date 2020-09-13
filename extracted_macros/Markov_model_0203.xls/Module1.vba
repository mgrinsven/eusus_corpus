Attribute VB_Name = "Module1"
Sub Spinner2_Change()
If Range("H13").Value = 99 Then Range("G13").Value = Range("G13").Value - 0.001
If Range("H13").Value = 101 Then Range("G13").Value = Range("G13").Value + 0.001
Range("H13").Value = 100
End Sub
