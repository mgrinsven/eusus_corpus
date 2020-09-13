Attribute VB_Name = "Module1"
Sub Macro_a__Loads_P7_1_data_into_PDWS()
Attribute Macro_a__Loads_P7_1_data_into_PDWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_a__Loads_P7_1_data_into_PDWS.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' Macro_a__Loads_P7_1_data_into_PDWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+a
'
    Application.Goto Reference:="C_ASSETS1_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS1_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB1_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY1_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_b__Loads_P7_2_data_into_PDWS()
Attribute Macro_b__Loads_P7_2_data_into_PDWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_b__Loads_P7_2_data_into_PDWS.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' Macro_b__Loads_P7_2_data_into_PDWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+b
'
Application.Goto Reference:="C_ASSETS2_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS2_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB2_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY2_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_c__Loads_P7_3_data_into_PDWS()
Attribute Macro_c__Loads_P7_3_data_into_PDWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_c__Loads_P7_3_data_into_PDWS.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' Macro_c__Loads_P7_3_data_into_PDWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+c
'
Application.Goto Reference:="C_ASSETS3_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS3_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB3_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY3_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_d__Loads_P7_4_data_into_PDWS_and_CONWS()
Attribute Macro_d__Loads_P7_4_data_into_PDWS_and_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_d__Loads_P7_4_data_into_PDWS_and_CONWS.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' Macro_d__Loads_P7_4_data_into_PDWS_and_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+d
'
    Application.Goto Reference:="C_ASSETS4_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS4_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB4_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY4_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS4_P"
    Selection.Copy
    Application.Goto Reference:="ASSETS_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIABEQ4"
    Selection.Copy
    Application.Goto Reference:="LIABEQ_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_p__Posts_the_basic_elimination_entry()
Attribute Macro_p__Posts_the_basic_elimination_entry.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_p__Posts_the_basic_elimination_entry.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Macro_p__Posts_the_basic_elimination_entry Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+p
'
    Application.Goto Reference:="BEE_EQUITY"
    Selection.Copy
    Application.Goto Reference:="BEE_EQUITY_EEC"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="BEE_INVESTMENT"
    Selection.Copy
    Application.Goto Reference:="BEE_INVESTMENT_EEC"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME2"
End Sub
Sub Macro_q__Posts_the_excess_cost_elimination_entry()
Attribute Macro_q__Posts_the_excess_cost_elimination_entry.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_q__Posts_the_excess_cost_elimination_entry.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro_q__Posts_the_excess_cost_elimination_entry Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+q
'
    Application.Goto Reference:="ECE_NOTERECEIVABLEINVENTORY"
    Selection.Copy
    Application.Goto Reference:="ECE_NOTERECEIVABLEINVENTORY_AEC"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ECE_L_BE_AD_P"
    Selection.Copy
    Application.Goto Reference:="ECE_L_BE_AD_P_AEC"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ECE_GOODWILL"
    Selection.Copy
    Application.Goto Reference:="ECE_GOODWILL_AEC"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_w__Erases_data_in_column_B_of_CONWS()
Attribute Macro_w__Erases_data_in_column_B_of_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_w__Erases_data_in_column_B_of_CONWS.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Macro_w__Erases_data_in_column_B_of_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+w
'
    Application.Goto Reference:="ASSETS_P_CONWS"
    Selection.ClearContents
    Application.Goto Reference:="LIABEQ_P_CONWS"
    Selection.ClearContents
    Application.Goto Reference:="APHOME2"
End Sub
Sub Macro_u__Erases_data_in_column_C_of_PDWS()
Attribute Macro_u__Erases_data_in_column_C_of_PDWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_u__Erases_data_in_column_C_of_PDWS.VB_ProcData.VB_Invoke_Func = "u\n14"
'
' Macro_u__Erases_data_in_column_C_of_PDWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+u
'
    Application.Goto Reference:="ASSETS_S_PDWS"
    Selection.ClearContents
    Application.Goto Reference:="LIAB_S_PDWS"
    Selection.ClearContents
    Application.Goto Reference:="EQUITY_S_PDWS"
    Selection.ClearContents
    Application.Goto Reference:="ASSETS_AEC"
    Selection.ClearContents
    Application.Goto Reference:="LIABEQ_AEC"
    Selection.ClearContents
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_e__Loads_P7_5_data_into_PDWS_and_CONWS()
Attribute Macro_e__Loads_P7_5_data_into_PDWS_and_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_e__Loads_P7_5_data_into_PDWS_and_CONWS.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Macro_e__Loads_P7_5_data_into_PDWS_and_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+e
'
Application.Goto Reference:="C_ASSETS5_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS5_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB5_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY5_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS5_P"
    Selection.Copy
    Application.Goto Reference:="ASSETS_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIABEQ5"
    Selection.Copy
    Application.Goto Reference:="LIABEQ_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_f__Loads_P7_6_data_into_PDWS_and_CONWS()
Attribute Macro_f__Loads_P7_6_data_into_PDWS_and_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_f__Loads_P7_6_data_into_PDWS_and_CONWS.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' Macro_f__Loads_P7_6_data_into_PDWS_and_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+f
'
Application.Goto Reference:="C_ASSETS6_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS6_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB6_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY6_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS6_P"
    Selection.Copy
    Application.Goto Reference:="ASSETS_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIABEQ6"
    Selection.Copy
    Application.Goto Reference:="LIABEQ_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_g__Loads_P7_7_data_into_PDWS_and_CONWS()
Attribute Macro_g__Loads_P7_7_data_into_PDWS_and_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_g__Loads_P7_7_data_into_PDWS_and_CONWS.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' Macro_g__Loads_P7_7_data_into_PDWS_and_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+g
'
Application.Goto Reference:="C_ASSETS7_S"
    Selection.Copy
    Application.Goto Reference:="C_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="NC_ASSETS7_S"
    Selection.Copy
    Application.Goto Reference:="NC_ASSETS_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIAB7_S"
    Selection.Copy
    Application.Goto Reference:="LIAB_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="EQUITY7_S"
    Selection.Copy
    Application.Goto Reference:="EQUITY_S_PDWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="ASSETS7_P"
    Selection.Copy
    Application.Goto Reference:="ASSETS_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="LIABEQ7"
    Selection.Copy
    Application.Goto Reference:="LIABEQ_P_CONWS"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="APHOME1"
End Sub
Sub Macro_x__Erases_postings_in_CONWS()
Attribute Macro_x__Erases_postings_in_CONWS.VB_Description = "Macro recorded 9/23/99 by College of Business"
Attribute Macro_x__Erases_postings_in_CONWS.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' Macro_x__Erases_postings_in_CONWS Macro
' Macro recorded 9/23/99 by College of Business
'
' Keyboard Shortcut: Ctrl+x
'
    Application.Goto Reference:="ASSETS_EEC"
    Selection.ClearContents
    Application.Goto Reference:="LIABEQ_EEC"
    Selection.ClearContents
    Application.Goto Reference:="APHOME2"
End Sub
