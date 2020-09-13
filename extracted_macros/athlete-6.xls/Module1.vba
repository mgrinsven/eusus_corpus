Attribute VB_Name = "Module1"
Sub Header()
Attribute Header.VB_Description = "Enters Your Name in left section of header."
Attribute Header.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Header Macro
' Enters Your Name in left section of header.
'
    ActiveSheet.PageSetup.LeftHeader = "Your Name"
End Sub
