Attribute VB_Name = "Module2"
Sub AllBuffer()

Dim buffer As Integer
buffer = Range("D20").Value
Range("G20").Value = buffer
Range("J20").Value = buffer
Range("M20").Value = buffer
Range("P20").Value = buffer

End Sub

Sub AllInv()

Dim inv As Integer
inv = Range("D22").Value
Range("G22").Value = inv
Range("J22").Value = inv
Range("M22").Value = inv
Range("P22").Value = inv

End Sub
