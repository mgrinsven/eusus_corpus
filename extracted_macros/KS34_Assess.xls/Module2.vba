Attribute VB_Name = "Module2"


'
' chartresults Macro
' Macro recorded 12/09/97 by manager
'
'
Sub chartresults()
Attribute chartresults.VB_Description = "Macro recorded 12/09/97 by manager"
Attribute chartresults.VB_ProcData.VB_Invoke_Func = " \n0"
    ActiveCell.Offset(2, 0).Range("A1:J4").Select
    ActiveSheet.ChartObjects.Add(48.75, 115.5, 264.75, 215.25).Select
    Application.CutCopyMode = False
    ActiveChart.ChartWizard Source:=Range("A3:J6"), Gallery:=xlLine, _
        Format:=4, PlotBy:=xlRows, CategoryLabels:=1, SeriesLabels _
        :=1, HasLegend:=1, Title:="", CategoryTitle:="", ValueTitle _
        :="", ExtraTitle:=""
End Sub
'
' chartresults2 Macro
' Macro recorded 12/09/97 by manager
'
'
Sub chartresults2()
Attribute chartresults2.VB_Description = "Macro recorded 12/09/97 by manager"
Attribute chartresults2.VB_ProcData.VB_Invoke_Func = " \n0"
    Range("A3:J6").Select
    ActiveSheet.ChartObjects.Add(48.75, 114.75, 209.25, 165).Select
    Application.CutCopyMode = False
    ActiveChart.ChartWizard Source:=Range("A3:J6"), Gallery:=xlLine, _
        Format:=4, PlotBy:=xlRows, CategoryLabels:=1, SeriesLabels _
        :=1, HasLegend:=1, Title:="", CategoryTitle:="", ValueTitle _
        :="", ExtraTitle:=""
End Sub
