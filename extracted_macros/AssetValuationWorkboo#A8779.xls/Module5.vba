Attribute VB_Name = "Module5"
'
' PrintInstructions Macro
' Macro recorded 11/2/99
'
'
Sub PrintInstructions()
Attribute PrintInstructions.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Instructions").Select
    Range("A1:J57").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
'
' PrintInput Macro
' Macro recorded 11/2/99
'
'
Sub PrintInput()
Attribute PrintInput.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Input").Select
    Range("A1:J79").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
'
' PrintShockValuesFixed15Year Macro
' Macro recorded 11/2/99
'
'
Sub PrintShockValuesFixed15Year()
Attribute PrintShockValuesFixed15Year.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Shock Values").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B24").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B47").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B75:L135").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
'
' PrintShockValuesARMS Macro
' Macro recorded 11/2/99
'
'
Sub PrintShockValuesARMS()
Attribute PrintShockValuesARMS.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Shock Values").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B98").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B121").Select
    ActiveWindow.LargeScroll Down:=1
    Range("A138:L162").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
'
' PrintShockValuesFixed30Years Macro
' Macro recorded 11/2/99
'
'
Sub PrintShockValuesFixed30Years()
Attribute PrintShockValuesFixed30Years.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Shock Values").Select
    Range("B1:L72").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
'
' PrintSummary Macro
' Macro recorded 11/2/99
'
'
Sub PrintSummary()
Attribute PrintSummary.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Shock Summary").Select
    Range("A3:H74").Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
