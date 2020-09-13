Attribute VB_Name = "Module3"
Sub PrintSupplement()
Attribute PrintSupplement.VB_Description = "Macro recorded 12/5/1999 by Kenneth G. Johnson"
Attribute PrintSupplement.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintSupplement Macro
' Macro recorded 12/5/1999 by Kenneth G. Johnson
'
Printall:
    Application.Goto Reference:="Copy5"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Copy6"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Copy7"
    Selection.PrintOut Copies:=1
    Application.Goto Reference:="Menu"
    Application.Goto Range("c4")

End Sub
    
    
   

