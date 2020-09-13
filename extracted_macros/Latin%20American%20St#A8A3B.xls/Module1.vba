Attribute VB_Name = "Module1"
Sub Country()
Attribute Country.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute Country.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Country Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Sheets("ALL").Select
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B3:Q466").Select
    Selection.Sort Key1:=Range("F3"), Order1:=xlAscending, Key2:=Range("B3") _
        , Order2:=xlAscending, Key3:=Range("C3"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub
Sub School()
Attribute School.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute School.VB_ProcData.VB_Invoke_Func = " \n14"
'
' School Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B3:Q436").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("B3"), Order1:=xlAscending, Key2:=Range("F3") _
        , Order2:=xlAscending, Key3:=Range("C3"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub
Sub LastName()
Attribute LastName.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute LastName.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LastName Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B3:Q39").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("C3"), Order1:=xlAscending, Key2:=Range("F3") _
        , Order2:=xlAscending, Key3:=Range("B3"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub
Sub ProfessionalExp()
Attribute ProfessionalExp.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute ProfessionalExp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ProfessionalExp Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B3:Q363").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("O3"), Order1:=xlAscending, Key2:=Range("C3") _
        , Order2:=xlAscending, Key3:=Range("B3"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub
Sub Interest()
Attribute Interest.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute Interest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Interest Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("P3"), Order1:=xlAscending, Key2:=Range("C3") _
        , Order2:=xlAscending, Key3:=Range("B3"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub
Sub FiltersON()
Attribute FiltersON.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute FiltersON.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FiltersON Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Range("A1").Select
End Sub
Sub FiltersOFF()
Attribute FiltersOFF.VB_Description = "Macro recorded 1/10/2003 by Patrick Cassereau"
Attribute FiltersOFF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FiltersOFF Macro
' Macro recorded 1/10/2003 by Patrick Cassereau
'

'
    Range("A1").Select
    Selection.AutoFilter
    Range("A1").Select
End Sub
