Attribute VB_Name = "Module2"

'
' PrintFRSMacro Macro
' Macro recorded 1/9/99 by Kenneth G. Johnson
'
Sub PrintFRSMacro()
    Application.Goto Reference:="Copy1"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Copy2"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Copy3"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Copy4"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Menu"
    Application.Goto Range("c4")
End Sub


'
'PrintWorksheet Macro
'Macro Recorded 1/9/1999 by Kenneth G. Johnson
'
'
Sub PrintWorksheet()
    Application.Goto Reference:="Worksheet"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Menu"
    Application.Goto Range("c4")
End Sub
'
'DataInput Macro
'Macro Recorded 1/9/1999 by Kenneth G. Johnson
'
'
Sub DataInput()
    Application.Goto Reference:="Copy1"
    Application.Goto Range("b14")
End Sub

