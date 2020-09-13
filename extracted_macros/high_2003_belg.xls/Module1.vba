Attribute VB_Name = "Module1"
            'This program is the property of District Audit
            'It may not be copied without permission
                                    'Chris Raspin
                                    'District Audit
                                    'May 1999
Option Explicit
Sub Auto_Open()
Sheets("Welcome").Select
Application.Calculation = xlAutomatic
End Sub
Sub ShowDialogBox()                         ' Keyboard Shortcut: Ctrl+a
Attribute ShowDialogBox.VB_Description = "Macro recorded 26/08/98 by Computer Manager"
Attribute ShowDialogBox.VB_ProcData.VB_Invoke_Func = "a\n14"
    Sheets("Dialogbox").Select
End Sub
Sub ShowScatter()                           'Shows the scatter plot
Attribute ShowScatter.VB_Description = "Macro recorded 26/08/98 by Computer Manager"
Attribute ShowScatter.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
        Histogram                           'Updates histogram in case user subsequently selects histogram tab
        Sheets("Scatter").Select
    Application.ScreenUpdating = True
End Sub
Sub ShowHistogram()                         'Shows the histogram
    Application.ScreenUpdating = False
        Histogram                           'updates histogram
        Sheets("Histogram").Select
        ActiveChart.Deselect
    Application.ScreenUpdating = True
End Sub
Sub UpdateTypeList()                        'Updates the list of types for use in drop down menus
    Application.ScreenUpdating = False
        Sheets("menus").Select
        ActiveSheet.PivotTables("PivotTable1").RefreshTable
        Sheets("dialogbox").Select
    Application.ScreenUpdating = True
End Sub
Sub Histogram()
Sheets("Selection").Select                  'Sorts the selected data
Rows("2:501").Select
Selection.Sort Key1:=Range("k2"), Order1:=xlDescending, Key2:=Range("d2") _
    , Order2:=xlDescending
Dim datapoints As Integer                   'Scales the chart for number of sites selected
datapoints = Sheets("Selection").Range("k1").Value
Sheets("histogram").SetSourceData Source:=Sheets("selection").Range("f1:j" & datapoints + 1), PlotBy:=xlColumns
End Sub

Sub ShowRanking()                         ' Shows Ranking sheet
    Sheets("Ranking").Select
End Sub
Sub Button3_Click()
Application.Goto Reference:="Country_Name"
End Sub
Sub Button4_Click()
Application.Goto Reference:="Data_types"
End Sub
Sub ShowHelp()
Application.Goto Reference:="Help"
End Sub
Sub ShowStructure()
Application.Goto Reference:="Structure"
End Sub
Sub ShowExample()
Application.Goto Reference:="Example"
End Sub
Sub ShowChange()
Application.Goto Reference:="Change_Names"
End Sub
