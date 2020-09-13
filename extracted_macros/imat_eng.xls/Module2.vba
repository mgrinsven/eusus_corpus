Attribute VB_Name = "Module2"
Sub view_graph1()
Attribute view_graph1.VB_Description = "Macro recorded 9/3/98 by ER"
Attribute view_graph1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' view_graph1 Macro
' Macro recorded 9/3/98 by ER
'
    Sheets("E. Graphs").Select
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$54"
    ActiveWindow.SelectedSheets.PrintPreview
End Sub
