Attribute VB_Name = "Module2"
Sub Macro_x__Erases_codes_exrates_and_two_dollar_amounts()
Attribute Macro_x__Erases_codes_exrates_and_two_dollar_amounts.VB_Description = "Macro recorded 9/21/99 by College of Business"
Attribute Macro_x__Erases_codes_exrates_and_two_dollar_amounts.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' Macro_x__Erases_codes_exrates_and_two_dollar_amounts Macro
' Macro recorded 9/21/99 by College of Business
'
' Keyboard Shortcut: Ctrl+x
'
    Application.Goto Reference:="CODES"
    Selection.ClearContents
    Application.Goto Reference:="RATES"
    Selection.ClearContents
    Application.Goto Reference:="REBOY_DOLLARS"
    Selection.ClearContents
    Application.Goto Reference:="REMEASUREMENTGAINLOSS"
    Selection.ClearContents
    Application.Goto Reference:="APHOME"
End Sub
