Attribute VB_Name = "Module1"
Sub Macro_a__Loads_P19_2_data()
Attribute Macro_a__Loads_P19_2_data.VB_Description = "Macro recorded 9/21/99 by College of Business"
Attribute Macro_a__Loads_P19_2_data.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' Macro_a__Loads_P19_2_data Macro
' Macro recorded 9/21/99 by College of Business
'
' Keyboard Shortcut: Ctrl+a
'
    Application.Goto Reference:="SALES1"
    Selection.Copy
    Application.Goto Reference:="SALES"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="INVENTORY_BEG1"
    Selection.Copy
    Application.Goto Reference:="INVENTORY_BEG"
    ActiveSheet.Paste
    Application.Goto Reference:="PURCHASES1"
    Application.CutCopyMode = False
    Selection.Copy
    Application.Goto Reference:="PURCHASES"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="INVENTORY_END1"
    Selection.Copy
    Application.Goto Reference:="INVENTORY_END"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DE_OE1"
    Selection.Copy
    Application.Goto Reference:="DE_OE"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="TAXEXPENSE1"
    Selection.Copy
    Application.Goto Reference:="TAXEXPENSE"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="REBOY1"
    Selection.Copy
    Application.Goto Reference:="REBOY"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DIVIDENDS1"
    Selection.Copy
    Application.Goto Reference:="DIVIDENDS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS1"
    Selection.Copy
    Application.Goto Reference:="ASSETS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB1"
    Selection.Copy
    Application.Goto Reference:="LIAB."
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="COMMONSTOCK1"
    Selection.Copy
    Application.Goto Reference:="COMMONSTOCK"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DIVIDENDS1"
    Selection.Copy
    Application.Goto Reference:="DIVIDENDS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME"
End Sub
Sub Macro_B__Loads_P19_3_data()
Attribute Macro_B__Loads_P19_3_data.VB_Description = "Macro recorded 9/21/99 by College of Business"
Attribute Macro_B__Loads_P19_3_data.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' Macro_B__Loads_P19_3_data Macro
' Macro recorded 9/21/99 by College of Business
'
' Keyboard Shortcut: Ctrl+b
'
Application.Goto Reference:="SALES2"
    Selection.Copy
    Application.Goto Reference:="SALES"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="INVENTORY_BEG2"
    Selection.Copy
    Application.Goto Reference:="INVENTORY_BEG"
    ActiveSheet.Paste
    Application.Goto Reference:="PURCHASES2"
    Application.CutCopyMode = False
    Selection.Copy
    Application.Goto Reference:="PURCHASES"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="INVENTORY_END2"
    Selection.Copy
    Application.Goto Reference:="INVENTORY_END"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DE_OE2"
    Selection.Copy
    Application.Goto Reference:="DE_OE"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="TAXEXPENSE2"
    Selection.Copy
    Application.Goto Reference:="TAXEXPENSE"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="REBOY2"
    Selection.Copy
    Application.Goto Reference:="REBOY"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DIVIDENDS2"
    Selection.Copy
    Application.Goto Reference:="DIVIDENDS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS2"
    Selection.Copy
    Application.Goto Reference:="ASSETS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB2"
    Selection.Copy
    Application.Goto Reference:="LIAB."
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="COMMONSTOCK2"
    Selection.Copy
    Application.Goto Reference:="COMMONSTOCK"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="DIVIDENDS2"
    Selection.Copy
    Application.Goto Reference:="DIVIDENDS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME"
End Sub
Sub Macro_w__Erases_data_in_column_B()
Attribute Macro_w__Erases_data_in_column_B.VB_Description = "Macro recorded 9/21/99 by College of Business"
Attribute Macro_w__Erases_data_in_column_B.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Macro_w__Erases_data_in_column_B Macro
' Macro recorded 9/21/99 by College of Business
'
' Keyboard Shortcut: Ctrl+w
'
    Application.Goto Reference:="SALES"
    Selection.ClearContents
    Application.Goto Reference:="INVENTORY_BEG"
    Selection.ClearContents
    Application.Goto Reference:="PURCHASES"
    Selection.ClearContents
    Application.Goto Reference:="INVENTORY_END"
    Selection.ClearContents
    Application.Goto Reference:="DE_OE"
    Selection.ClearContents
    Application.Goto Reference:="TAXEXPENSE"
    Selection.ClearContents
    Application.Goto Reference:="REBOY"
    Selection.ClearContents
    Application.Goto Reference:="DIVIDENDS"
    Selection.ClearContents
    Application.Goto Reference:="ASSETS"
    Selection.ClearContents
    Application.Goto Reference:="LIAB"
    Selection.ClearContents
    Application.Goto Reference:="COMMONSTOCK"
    Selection.ClearContents
    Application.Goto Reference:="APHOME"
End Sub
